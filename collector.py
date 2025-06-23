import streamlit as st
import json
import csv
import os
import pandas as pd
import s3fs
import numpy as np
import xlsxwriter
import io
from PIL import Image
import plotly.express as px
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode
from pyxlsb import open_workbook as open_xlsb
from streamlit_extras.customize_running import center_running
import time
from datetime import datetime

def app():

    with st.sidebar.expander("ABOUT THE APP"):
        st.markdown("""
                Each form submitted during Audit or Construction activities is
                automatically saved in the cloud.
                This app retrieves collected data in a shape of single dataframe.
                Afterwards, it can be downloaded in csv or xlsx, depending on the preferences.
                To do so, in the sidebar menu:  
                    1. Choose the category: Audits, Construction or leave it blank for entire library of forms (not recommended).  
                    2. Choose Project Number or leave it blank for All in the selected category.  
                    3. Choose Form name or leave it blank for All in the selected project.  
                    4. Choose Download type. "Complete" will pull all fields, "Faults & Materials" will focus only on those 2 categories.  
                    5. Use "Collect Data" button.  
                    6. Download to csv or xlsx with the dedicated buttons.  
                    

                Please note that ID column refers to different objects, depending on
                form purpose. For EHT circuits it will be EHT circuit number, for Panel accordingly
                it will be Panel name. However, for e.g. Acceptance Form it shall be form name.
                """)


    fs = s3fs.S3FileSystem(anon=False)

    @st.cache_data(ttl=600)
    def list_files(search_item):
        return fs.find(search_item)

    @st.cache_data(ttl=600)
    def read_file(filename):
        with fs.open(filename) as f:
            return f.read().decode("utf-8")

    if "load_state" not in st.session_state:
       st.session_state.load_state = False

    @st.cache_data(ttl=1800)
    def get_csvsource(file_name):
        file_content = read_file(file_name)
        data = pd.read_csv(io.StringIO(file_content))
        return data


    @st.cache_data(ttl=1800)
    def collect_data(directory,form_name):

        searched_dir = f"{directory}{form_name}/"
        #list_of_files = fs.find(searched_dir)
        list_of_files = list_files(searched_dir)
        searched_files = []

        form_df = None

        for s3_file in list_of_files:
            if s3_file.split('.')[-1] == 'json':
                searched_files.append(s3_file)

        ###ADDED
        id_corr_dict = {}
        id_date_dict = {}
        for file in searched_files:
            unit_id = file[0:-24]
            if '/Audits/' in unit_id:
                split_1 = unit_id.split('Audits/')[1].split('/',1)
            else:
                split_1 = unit_id.split('Construction/')[1].split('/',1)
            uniq_proj = split_1[0]
            form_pref = split_1[1].split(' - ')[0]    
            idno = split_1[1].split('/',1)[1].split('_',1)[-1].split('_NF')[0]
            unique_id =  uniq_proj + '_' + form_pref + '_' + idno   
            pfdate = file[-24:-5]
            #st.write(file)
            #st.write(unit_id)
            #st.write(unique_id)
            pfdate_format = datetime.strptime(pfdate,'%Y-%m-%d_%H_%M_%S')
            #if unit_id not in id_date_dict.keys():
            if unique_id not in id_corr_dict.keys():
                id_corr_dict[unique_id] = unit_id
                id_date_dict[unit_id] = pfdate
            else:
                #st.write(id_date_dict)
                #st.write(id_date_dict[unit_id])
                existing_form = id_date_dict[id_corr_dict[unique_id]]
                exist_date = datetime.strptime(existing_form,'%Y-%m-%d_%H_%M_%S')
                if pfdate_format > exist_date:
                    id_date_dict[unit_id] = pfdate
                    id_corr_dict[unique_id] = unit_id
        list_of_ujsons = []
        for unitid, udate in id_date_dict.items():
            list_of_ujsons.append(unitid+udate+'.json')       
        ###ADDED


        form_df = pd.DataFrame({'Project':[],'Form':[],'ID':[], 'Section':[], 'Item':[],
                    'Question_label':[],'Question_name':[],
                        'Question_subname':[],'Answer':[],'Status':[],'Data_type':[]})


        ###for jejson in searched_files:
        for jejson in list_of_ujsons:
            #st.write(jejson)
            file_content = read_file(jejson)
            data = json.loads(file_content)
            #eht_cct_no = data['pages'][0]['sections'][0]['answers'][4]['values'][0]
            if '/Audits/' in jejson:
                proj_split = jejson.split('Audits/')[1].split('/',1)
            else:
                proj_split = jejson.split('Construction/')[1].split('/',1)   
            project_no = proj_split[0]
            form_split = proj_split[1].split('/',1)
            form_no = form_split[0]
            id_no = form_split[1].split('_',1)[-1].split('_NF')[0]
            (project_id,form_id,unit_id,section_name,question_label,question_name,question_subname,
            answer_value,branch_no,data_type,exception_type) = ([] for i in range(11))

            fault_colors = ['#F6E2DF','#C0392B']
            for item in data['pages'][0]['sections']:
                #if item['type'] == 'Flow':
                if item['type'] != 'Repeat':
                    for answer in item['answers']:
                        section_name.append(item['label'])
                        question_label.append(answer['label'])
                        question_name.append(answer['question'])
                        question_subname.append('')
                        data_type.append(answer['dataType'])
                        branch_no.append(1)
                        unit_id.append(id_no)
                        project_id.append(project_no)
                        form_id.append(form_no)
                        try:
                            answer_value.append(answer['values'][0])
                        except:
                            answer_value.append('')
                        try:
                            if answer['valuesMetadata'][0]['exception']['backgroundColor'] in fault_colors:
                                exception_type.append('Fault')
                            else:
                                exception_type.append('OK')  
                        except:
                            exception_type.append('')
                if item['type'] == 'Repeat':
                    branch_count = 1
                    for branch in item['rows']:
                        for elem in branch['pages'][0]['sections'][0]['answers']:
                            section_name.append(item['label'])
                            question_label.append(elem['label'])
                            question_name.append(item['name'])
                            try:
                                question_subname.append(elem['question'])
                            except:
                                question_subname.append('')
                            data_type.append(elem['dataType'])
                            branch_no.append(branch_count)
                            unit_id.append(id_no)
                            project_id.append(project_no)
                            form_id.append(form_no)
                            try:
                                answer_value.append(elem['values'][0])
                            except:
                                answer_value.append('')
                            try:
                                if elem['valuesMetadata'][0]['exception']['backgroundColor'] in fault_colors:
                                    exception_type.append('Fault')
                                else:
                                    exception_type.append('OK')
                            except:
                                exception_type.append('') 
                        branch_count += 1


            form_df = pd.concat([form_df,
                                pd.DataFrame({
                                    'Project':project_id,
                                    'Form':form_id,
                                    'ID':unit_id,
                                    'Section':section_name,
                                    'Item':branch_no,
                                    'Question_label':question_label,
                                    'Question_name':question_name,
                                    'Question_subname':question_subname,
                                    'Answer':answer_value,
                                    'Status':exception_type,
                                    'Data_type':data_type
                                            })
                                            ])

        form_df = form_df.astype({"Item": int})
        form_df = form_df[form_df['Data_type'] != 'Image']
        form_df.drop(columns='Data_type',inplace=True)
        
        form_df['Question_name'] = form_df['Question_name'].astype(str)
        form_df['Question_subname'] = form_df['Question_subname'].astype(str)

        form_df['Question_name'] = form_df['Question_name'].apply(lambda x: x.strip())
        form_df['Question_subname'] = form_df['Question_subname'].apply(lambda x: x.strip())

        form_df['Question_name'] = form_df['Question_name'].replace(regex=r'\s+([?.!"])', value=r'\1')
        form_df['Question_subname'] = form_df['Question_subname'].replace(regex=r'\s+([?.!"])', value=r'\1')

        form_df['Question_name'] = form_df['Question_name'].replace(regex=['  '], value=' ')
        form_df['Question_subname'] = form_df['Question_subname'].replace(regex=['  '], value=' ')

        return form_df

    #@st.cache_data()
    # def to_excel(df):
    #     output = io.BytesIO()
    #     writer = pd.ExcelWriter(output, engine='xlsxwriter')
    #     tab_name = form_select.split('-')[0]
    #     df.to_excel(writer, index=False, sheet_name=tab_name)
    #     workbook = writer.book
    #     worksheet = writer.sheets[tab_name]
    #     format1 = workbook.add_format({'num_format': '0.00'}) 
    #     worksheet.set_column('A:A', None, format1)  
    #     writer.close()
    #     processed_data = output.getvalue()
    #     return processed_data

       
    def to_excel(df,fname):
        output = io.BytesIO()
        # workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        # worksheet = workbook.add_worksheet()
        # worksheet.write('A1', 'Hello')
        # workbook.close()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        #tab_name = form_select.split('-')[0]
        df.to_excel(writer, index=False, sheet_name=fname)
        workbook = writer.book
        worksheet = writer.sheets[fname]
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    #@st.cache_data() 
    def convert_results(df):
        return df.to_csv().encode('utf-8')

    #if st.sidebar.button('Collect Data'):
    with st.form('input'):
        proj_dir = "s3-nvent-prontoforms-data/"
        form_select = ''
        with st.sidebar:
            category_select = st.sidebar.selectbox('Select Category',('','Audits','Construction'),key='category_selection')
            cat_dir = f"s3-nvent-prontoforms-data/{category_select}/"
            #projects_list = fs.find(cat_dir)
            projects_list = list_files(cat_dir)
            project_nos = []
            for pfile in projects_list:
                try:
                    project_nos.append(pfile.split('/')[2])
                except:
                    continue
            if category_select:
                project_select = st.sidebar.selectbox('Select a Project',np.unique(project_nos).tolist(),key='project_selection')
                proj_dir = f"{cat_dir}{project_select}/"
                #st.write(proj_dir)
                #forms_list = fs.find(proj_dir)
                forms_list = list_files(proj_dir)
                #st.write(forms_list)
                form_nos = [""]
                for ffile in forms_list:
                    try:
                        form_nos.append(ffile.split('/')[3])
                    except:
                        continue
                if project_select:
                    form_select = st.sidebar.selectbox('Select a Form',np.unique(form_nos).tolist(),key='form_selection')

            download_type = st.sidebar.radio(
                "***Download type***",
                ["Complete", "Faults & Materials"]
            )
            st.write('<style>div.row-widget.stRadio > div{flex-direction:row;}</style>', unsafe_allow_html=True)


            submit_button = st.form_submit_button('Collect Data')
        
        map_insulation = {'Code 1 Count':'Missing/Damaged Sealant',
                        'Code 2 Count':'Damaged Insulation/Cladding',
                        'Code 3 Count':'Missing Cladding',
                        'Code 4 Count':'Missing Insulation and Cladding and/or Blankets',
                        'Code 5 Count':'Other'
                        }


        if submit_button:
            center_running()
            time.sleep(2)
        #if submit_button or st.session_state.load_state:
            st.session_state.load_state = True

            translate_csv = get_csvsource('s3-nvent-prontoforms-data/Data_sources/translation.csv')

            #st.table(translate_csv)

            form_df = collect_data(proj_dir,form_select)
            #form_df['Question_subname'] = form_df['Question_subname'].map(map_insulation)
            form_df['Question_subname'] = form_df['Question_subname'].replace(map_insulation)
            form_df['Question_name'] = form_df['Question_name'].replace(translate_csv['Translation'].tolist(),translate_csv['Question'].tolist())
            form_df['Question_subname'] = form_df['Question_subname'].replace(translate_csv['Translation'].tolist(),translate_csv['Question'].tolist())
            #form_df['Section'] = form_df['Section'].replace(translate_csv['Translation'],translate_csv['Question'])
            #st.table(form_df)
            #st.session_state.load_state = False

            gb_form_df = GridOptionsBuilder.from_dataframe(form_df)
            gb_form_df.configure_default_column(enablePivot=True, enableValue=True, enableRowGroup=True) # , groupable=True, editable=True
            gb_form_df.configure_selection(
                selection_mode='mulitple',use_checkbox=False,groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
            gb_form_df.configure_side_bar()
            #gb_form_df.configure_pagination(paginationAutoPageSize=True)
            #paginationAutoPageSize=True,
            gridOptions_form_df = gb_form_df.build()
            #gridOptions_form_df['suppressHorizontalScroll'] = False

            grid_return_form_df = AgGrid(
                form_df,
                gridOptions=gridOptions_form_df,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED, 
                update_mode = GridUpdateMode.VALUE_CHANGED | GridUpdateMode.SELECTION_CHANGED | GridUpdateMode.FILTERING_CHANGED | GridUpdateMode.SORTING_CHANGED,
                fit_columns_on_grid_load=False,
                theme='streamlit',
                enable_enterprise_modules=True,
                height=350, 
                width='100%',
                reload_data=True,
                columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS
            )
            gb_form_df_data = grid_return_form_df['data']
            
            metric_col1, metric_col2, metric_col3 = st.columns(3)
            metric_col1.metric('# Projects',gb_form_df_data['Project'].nunique())
            metric_col2.metric('# Form Types',gb_form_df_data['Form'].nunique())
            metric_col3.metric('# Forms',len(gb_form_df_data[['Project','Form','ID']].drop_duplicates()))

            fault_df = gb_form_df_data[gb_form_df_data['Status']=='Fault']
            fault_df = fault_df.groupby(['Project','Form','ID','Question_name','Question_subname']).agg({'Status':'count'}).reset_index()
            fault_df.rename(columns={'Status':'Quantity'},inplace=True)
            fault_df_graph = fault_df[['Project','Form','Question_name','Question_subname','Quantity']]
            fault_df_graph.Question_name[fault_df_graph.Question_name=='Field Insulation Inspections'] = fault_df_graph.Question_subname
            fault_df_graph = fault_df_graph.groupby(['Project','Form','Question_name','Question_subname']).agg({'Quantity':'sum'}).reset_index()

            mat_df_type = gb_form_df_data[
                (gb_form_df_data['Section'].str.contains("_Material")) & (gb_form_df_data['Question_subname'].str.startswith("Required Material")) ]
            mat_df_type = mat_df_type[mat_df_type['Answer']!=""]
            #mat_df_type = mat_df_type[~mat_df_type['Answer'].isnull()]
            mat_df_qty = gb_form_df_data[
                (gb_form_df_data['Section'].str.contains("_Material")) & (gb_form_df_data['Question_subname'].str.startswith("Quantity")) ]
            mat_df_qty['Answer'] = mat_df_qty['Answer'].replace("",1)
            mat_df_qty.rename(columns={'Answer':'Quantity'},inplace=True)
            #mat_df_qty = mat_df_qty[mat_df_qty['Answer']!=""]
            #st.write(len(mat_df_type), len(mat_df_qty))                
            mat_df_type.reset_index(inplace=True)
            mat_df_qty.reset_index(inplace=True)
            #mat_df_type['Quantity'] = mat_df_qty['Answer']
            #st.dataframe(mat_df_type)
            #st.dataframe(mat_df_qty)
            mat_df_type = pd.merge(mat_df_type, mat_df_qty, how='left', on=['Project','Form','ID','Section','Item','Question_name'])
            #mat_df = gb_form_df_data[gb_form_df_data['Section'].str.contains("_Material")]
            #mat_df_type = mat_df[(mat_df['Question_subname'].str.startswith("Required material")) & (~mat_df['Answer'].isnull())]




            #st.table(mat_df_type)
            #st.table(mat_df_qty)
            mat_df_type['Quantity'] = mat_df_type['Quantity'].astype(int)
            mat_df_type = mat_df_type.groupby(['Project','Form','Question_name','ID','Answer']).agg({'Quantity':'sum'}).reset_index()
            mat_df_type.rename(columns={'Answer':'Material'},inplace=True)
            mat_df_type_graph = mat_df_type.groupby(['Project','Form','Material']).agg({'Quantity':'sum'}).reset_index()
            # fig1 = px.bar(fault_df, x="Quantity", y="Question_name", color="Project",
            #         barmode='group',template='gridon', title = "Reported Faults", orientation='h', 
            #         #category_orders={'TaskName':task_list},
            #         height=500)
            # st.plotly_chart(fig1,use_container_width=True)           

            if len(fault_df)!=0:
                fig2 = px.scatter(fault_df_graph, x="Form", y='Question_name', color="Project", size='Quantity',
                template='gridon', title = "Reported Faults",
                #category_orders={'TaskName':task_list},
                height=500)
                st.plotly_chart(fig2,use_container_width=True)
            else:
                st.write('NO FAULTS RECORDED')

            if len(mat_df_type_graph)!=0:
                fig3 = px.scatter(mat_df_type_graph, x="Form", y="Material", color="Project", size='Quantity',
                template='gridon', title = "Reported Replacement Materials", 
                #category_orders={'TaskName':task_list},
                height=500)
                st.plotly_chart(fig3,use_container_width=True)
            else:
                st.write('NO MATERIAL REPLACEMENTS PROPOSED')


            if category_select!="":
                def_project = project_select
            else:
                def_project = "project"

            if form_select:
                def_form = form_select.split('-')[0]
            else:
                def_form = "form"

            user_input_full = st.sidebar.text_input("Name your file: ", max_chars = 30,value = def_project + "_" + def_form)



            if user_input_full and download_type == "Complete":  

                xl_to_csv_full = convert_results(form_df)
                #xl_to_csv_fault = convert_results(fault_df)

                #st.sidebar.markdown("Complete Download")

                st.sidebar.download_button(
                    label="Download to CSV - All",
                    data=xl_to_csv_full,
                    file_name=user_input_full + "_complete.csv"
                    )

                df_xlsx_full = to_excel(form_df,user_input_full)
                #xl_name = f"{category_select}_{project_select}_{form_select}.xlsx"
                st.sidebar.download_button(label='Download to EXCEL - All',
                                                data=df_xlsx_full ,
                                                file_name= user_input_full + "_complete.xlsx"
                                                #mime="application/vnd.ms-excel"
                                                )
            #file_name= f"{category_select}_{project_select}_{form_select}.xlsx"

            if user_input_full and download_type == "Faults & Materials":  
                
                xl_to_csv_fault = convert_results(fault_df)
                xl_to_csv_mat = convert_results(mat_df_type)              
                #st.sidebar.markdown("Faults Download")

                st.sidebar.download_button(
                    label="Download to CSV - Faults",
                    data=xl_to_csv_fault,
                    file_name=user_input_full + "_faults.csv"
                    )

                df_xlsx_fault = to_excel(fault_df,user_input_full)
                #xl_name = f"{category_select}_{project_select}_{form_select}.xlsx"
                st.sidebar.download_button(label='Download to EXCEL - Faults',
                                                data=df_xlsx_fault ,
                                                file_name= user_input_full + "_faults.xlsx"
                                                #mime="application/vnd.ms-excel"
                                                )

                st.sidebar.download_button(
                    label="Download to CSV - Materials",
                    data=xl_to_csv_mat,
                    file_name=user_input_full + "_materials.csv"
                    )

                df_xlsx_mat = to_excel(mat_df_type,user_input_full)
                #xl_name = f"{category_select}_{project_select}_{form_select}.xlsx"
                st.sidebar.download_button(label='Download to EXCEL - Materials',
                                                data=df_xlsx_mat ,
                                                file_name= user_input_full + "_materials.xlsx"
                                                #mim
                                                )
    