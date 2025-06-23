import streamlit as st
import json
import csv
import os
import pandas as pd
import s3fs
import numpy as np
#import xlsxwriter
import io
from PIL import Image
import plotly.express as px
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode
from base64 import b64encode
import requests
import datetime
import pytz
import time
from streamlit_extras.customize_running import center_running

def app():

    with st.sidebar.expander("ABOUT THE APP"):
        st.markdown("""
                    Easy form dispatch to selected Auditor.

                    1. Verify you have dispatch priviledges by entering your EID (e.g. E2009911)
                    2. Select Recipient (truecontext app user)  
                    3. Select desired form type  
                    4. Select project  
                    5. Populate values and attach PDFs to appear in the dispatched form  

                    Please note: Files will not be transferred as attachments with the submission.
                                 The purpose of dispatching them is mainly to enable easy access to drawings while performing the audit.
                                 In case there is a need to illustrate a fault on e.g. isometric, recommented workflow is to:  
                                    -	Annotate it accordingly on pdf (then save the change or not, depending on further actions)  
                                    -	Take a screenshot and save it on a device  
                                    -	Attach the picture to the subject question  
                """)



    fs = s3fs.S3FileSystem(anon=False)

    @st.cache_data(ttl=600)
    def read_file(filename):
        return fs.open(filename)

    @st.cache_data(ttl=3600)
    def get_csvsource(file_name):
        file_content = read_file(file_name)
        data = pd.read_csv(file_content, encoding="cp1252")
        return data

    def get_csvlog(file_name):
        file_content = read_file(file_name)
        data = pd.read_csv(file_content, encoding="cp1252")
        return data

    def dwg_base64(byte_file):
        file_content = byte_file.read()
        base64_file = b64encode(file_content)
        return base64_file.decode()

    def save_csv(df,filename):
        with fs.open(filename, 'w') as file:
            return df.to_csv(file, index=False)


    def form_373(formid,userid,username,note,duedate,priority,project,
                auditscope,
                cct,panel,
                iso,iso_dwg,
                pid,pid_dwg,
                sld,sld_dwg,
                pcl,pcl_dwg,
                tmaint,tlim,iop,
                cbnum,supply,thmethod):


        form_dict = {
                "formId": formid,
                "userId": userid,
                "username": username,
                "dispatchToDraft": "false",
                "suppressPush": "true",        
                "metadata": 
                {
                "notes": note,
                "dueDate": duedate,
                "priority": priority
                }
                }
        
        map_power_fr = {'Single Phase':'Mnophasé',
                        'Three Phase Delta': 'Triphasé Triangle',
                        'Three Phase Star': 'Triphasé Etoile'}
        
        map_power_de = {'Single Phase':'Einphasig',
                        'Three Phase Delta': 'Dreieck-Schaltung, 3-phasig',
                        'Three Phase Star': 'Stern-Schaltung, 3-phasig'}
        
        map_control_fr = {'Ambient Sensing - Field Controller': 'Détection T°C  Ambiante - Controleur Local',
                        'Ambient Sensing - Panel Controller': 'Détection T°C  Ambiante - Controleur en Armoire',
                        'Line Sensing - Field Controller': 'Détection T°C Ligne/Équipement - Controleur Local',
                        'Line Sensing - Panel Controller with Field Sensor': 'Détection T°C ligne/Équipement - Controleur en Armoire & Sonde Locale',
                        'Uncontrolled': 'Non Controlé (Sans Régulation de T°C)',
                        'Other': 'Autre'}

        map_control_de = {'Ambient Sensing - Field Controller': 'Feld-Thermostat/Umgebungsfühler',
                        'Ambient Sensing - Panel Controller': 'Schaltschrank-Thermostat/Umgebungsfühler',
                        'Line Sensing - Field Controller': 'Feld-Thermostat/Anlegefühler',
                        'Line Sensing - Panel Controller with Field Sensor': 'Schaltschrank-Thermostat/Anlegefühler',
                        'Uncontrolled': 'Ungeregelt',
                        'Other': 'Andere'}

        if formid == '1421926165':
            pr_no_method = "Automatic"
            auditscope = auditscope
        elif formid == '1425061023':
            pr_no_method = "Automatique"
            auditscope = auditscope.replace('Level', 'Niveau')
            supply = supply.replace(supply, map_power_fr[supply])
            thmethod = thmethod.replace(thmethod, map_control_fr[thmethod])
        elif formid == '1424996084':
            pr_no_method = "Automatisch"
            auditscope = auditscope.replace('Level', 'Stufe')
            supply = supply.replace(supply, map_power_de[supply])
            thmethod = thmethod.replace(thmethod, map_control_de[thmethod])

        data_lst = [
                    {"label":"AuditScope",
                    "answer":auditscope
                    },        
                    {"label":"ProjectNoMethod",
                    "answer":pr_no_method
                    },
                    {"label":"ProjectNo_Auto",
                    "answer":project
                    },
                    {"label":"nvt_ProjectNo",
                    "answer":project
                    },
                    {"label":"EHT_CircuitTagNo",
                    "answer":cct
                    },
                    {"label":"EHT_PanelTagNo",
                    "answer":panel
                    },
                    {"label":"EHTIsoNum",
                    "answer":iso
                    },                               
                    {"label":"PIDNum",
                    "answer":pid
                    },                 
                    {"label":"SLDNum",
                    "answer":sld
                    },                 
                    {"label":"CableListNum",
                    "answer":pcl
                    },                 
                    {"label":"CBNum",
                    "answer":cbnum
                    },
                    {"label":"PowerSupplySystem",
                    "answer":supply
                    },
                    {"label":"ThCtrlMethodDoc",
                    "answer":thmethod
                    }
                    ]

        temp_maint = {"label":"MaintainTempC",
                    "answer":tmaint
                    }

        temp_limit = {"label":"Ins_ThFIeldLimiter",
                    "answer":tlim
                    }

        current_op = {"label":"OperatingCurrentAmp",
                    "answer":iop
                    }

        iso_attach = {"label":"Related_ISO",
                    "answer":{
                    "contentType": "application/pdf",
                    "bytes":iso_dwg,
                    #"filename": "iso_dwg.pdf"
                            }
                    }

        pid_attach = {"label":"Related_PID",
                    "answer":{
                    "contentType": "application/pdf",
                        "bytes":pid_dwg,
                    #"filename": "pid.pdf"
                            }
                    }

        sld_attach = {"label":"Related_SLD",
                    "answer":{
                    "contentType": "application/pdf",
                        "bytes":sld_dwg,
                    #"filename": "sld.pdf"
                            }
                    }

        pcl_attach = {"label":"Related_PCL",
                    "answer":{
                    "contentType": "application/pdf",
                        "bytes":pcl_dwg,
                    #"filename": "pcl.pdf"
                    }
                    }

        if tmaint is not None:
            data_lst.append(temp_maint)
        if tlim is not None:
            data_lst.append(temp_limit)
        if iop is not None:
            data_lst.append(current_op)        
        if iso_dwg is not None:
            data_lst.append(iso_attach)
        if pid_dwg is not None:
            data_lst.append(pid_attach)
        if sld_dwg is not None:
            data_lst.append(sld_attach)
        if pcl_dwg is not None:
            data_lst.append(pcl_attach)

        form_dict['data'] = data_lst

        get_type = 'api/1.1/data/dispatch.json'
        url = f"https://api.prontoforms.com/{get_type}"
        response = requests.post(url,json=form_dict, auth=(st.secrets['pf_username'], st.secrets['pf_password']))

        return response

    users_df = get_csvsource('s3-nvent-prontoforms-data/Data_sources/users.csv')
    forms_df = get_csvsource('s3-nvent-prontoforms-data/Data_sources/forms.csv')
    projects_df = get_csvsource('s3-nvent-prontoforms-data/Data_sources/SAP_projects.csv')


    with st.form('input'):
    #with st.form_submit_button('input'):

        with st.sidebar:

            # def dispatcher_check():
            #     if dispatcher_select in users_df.user_name.tolist():
            #         st.session_state['dispatcher_status'] = True
            #     else:
            #         st.session_state['dispatcher_status'] = False
            #dispatcher_select = st.sidebar.selectbox('Select Dispatcher',users_df.full_name.tolist(),key='disp_selection')
            dispatcher_select = st.text_input('Enter Dispatcher EID', key='disp_selection')
            print(f"Test: {users_df}")
            if dispatcher_select.lower() in users_df.user_name.tolist():
                dispatcher_name = users_df.loc[users_df.user_name == dispatcher_select.lower(),'full_name'].iloc[0]
            #st.write(dispatcher_select)
            # if st.session_state['dispatcher_status']:
            #     #dispatcher_status = True
            #     dispatcher_name = users_df.loc[users_df.user_name == dispatcher_select,'full_name'].iloc[0]
            #     st.markdown(f"""👋 Welcome {dispatcher_name}""")
            # else:
            #     #dispatcher_status = False
            #     st.warning("Not a valid Dispatcher EID!")
            #dispatcher_name = ''
            #form_id = None
            #dispatcher_name = None
            if st.form_submit_button('Verify Dispatcher'):
                if dispatcher_select.lower() not in users_df.user_name.tolist():
                    st.session_state['dispatcher_status'] = False
                    st.warning("Not a valid Dispatcher EID!")
                else:
                    st.session_state['dispatcher_status'] = True
                    st.success('Dispatcher verified!')
                    #dispatcher_name = users_df.loc[users_df.user_name == dispatcher_select.lower(),'full_name'].iloc[0]
                    st.markdown(f"""👋 Welcome {dispatcher_name}""")
            

            user_list = users_df.full_name.tolist()
            # if dispatcher_name !='':
            #     default_index = user_list.index(dispatcher_name)
            try:
                default_user = user_list.index(dispatcher_name)
            except:
                default_user = 0


            user_select = st.sidebar.selectbox('Select Recipient',user_list,index=default_user,key='user_selection')
            user_id = users_df.loc[users_df.full_name == user_select,'user_id'].iloc[0] 
            #st.write(user_id)
            form_select = st.sidebar.selectbox('Select Form',np.unique(forms_df.form_name).tolist(),key='form_selection')
            lang_select = st.sidebar.selectbox('Select Language',forms_df.language.tolist(),key='lang_selection')
            form_id = forms_df.loc[(forms_df.form_name == form_select) & (forms_df.language == lang_select),'form_id'].iloc[0]
            #st.write(form_id)
            project_select = st.sidebar.selectbox('Select Project',projects_df.Dropdown.tolist(),key='project_selection')
            project_id = projects_df.loc[projects_df.Dropdown == project_select,'Project Definition'].iloc[0]
            #st.write(project_id)
            submit_button = st.form_submit_button('Dispatch Form')

        # CCT
        if form_id in (1421926165, 1425061023, 1424996084) and len(user_select) > 0 and len(project_select) > 0 and 'dispatcher_status' in st.session_state and st.session_state['dispatcher_status']:

            col_01, col_02, col_03, col_04, col_05, col_06 = st.columns([2,1,1,1,1,1])
            with col_01:
                notes = st.text_input('Dispatch note:',key='note_373')
            with col_02:
                auditscope = st.selectbox('Audit scope:',
                                            ('Level 1','Level 2','Level 3'),
                                            key='auditscope_373')
            with col_03:
                due_date = st.date_input('Due date',key='due_date_373')
            with col_04:
                priority = st.selectbox('Priority',('Low','Medium','High'),key='priority_selection_373')
            with col_05:
                eht_cct = st.text_input('EHT Circuit Tag Number:',key='cct_tag_373')
            with col_06:
                panel_no = st.text_input('EHT Panel Tag Number:',key='panel_tag_373')

            col_11, col_12, col_13, col_14 = st.columns([1,1,1,1])
            with col_11:     
                iso_no = st.text_input('EHT Isometric number:',key='iso_no_373')
            with col_12:
                pid_no = st.text_input('P&ID Number:',key='pidno_373')
            with col_13:
                sld_no = st.text_input('Single Line Diagram Number:',key='sldno_373')
            with col_14:
                pcl_no = st.text_input('Cable List Number:',key='pcl_373')

            col_21, col_22, col_23, col_24 = st.columns([1,1,1,1])
            with col_21:
                uploaded_iso = st.file_uploader("Attach ISO drawing", accept_multiple_files=False,type="pdf")
                if uploaded_iso is not None:
                    iso_dwg = dwg_base64(uploaded_iso)
                else:
                    iso_dwg = None
            with col_22:
                uploaded_pid = st.file_uploader("Attach P&ID document", accept_multiple_files=False,type="pdf")
                if uploaded_pid is not None:
                    pid_dwg = dwg_base64(uploaded_pid)
                else:
                    pid_dwg = None
            with col_23:
                uploaded_sld = st.file_uploader("Attach Single Line Diagram", accept_multiple_files=False,type="pdf")
                if uploaded_sld is not None:
                    sld_dwg = dwg_base64(uploaded_sld)
                else:
                    sld_dwg = None
            with col_24:
                uploaded_pcl = st.file_uploader("Attach Cable List", accept_multiple_files=False,type="pdf")
                if uploaded_pcl is not None:
                    pcl_dwg = dwg_base64(uploaded_pcl)
                else:
                    pcl_dwg = None
            #col_41, col_42 = st.columns([1,1])

            col_21, col_22, col_23, col_24, col_25, col_26 = st.columns([1,1,1,1,1,2])
            with col_21:
                #temp_maintain = st.number_input('Maintain Temperature [°C]:',min_value = -500.0, max_value = 500.0, value = 5.0, step = 0.1)
                temp_maintain = st.text_input('Maintain Temperature [°C]:',key='tm_373')
            with col_22:
                temp_limiter = st.text_input('Limiter Setpoint [°C]:',key='tl_373')
                #temp_limiter = st.number_input('Limiter Setpoint [°C]:',min_value = -500.0, max_value = 500.0, value = 5.0, step = 0.1)
            with col_23:       
                #current_op = st.number_input('Operating Current [A]:', min_value = 0.1, max_value = 150.0, value = 10.0, step = 0.1)
                current_op = st.text_input('Operating Current [A]:',key='Ia_373')
            #col_31, col_32, col_33, col_34 = st.columns([1,1,2,2])
            with col_24:
                cb_no = st.text_input('CB No. in EHT Panel:',key='cbno_373')
            with col_25:
                supply_system = st.selectbox('Power Supply System',('Single Phase','Three Phase Delta','Three Phase Star'),key='supply_selection_373')
            with col_26:
                control_method = st.selectbox(
                    'Temperature Control Method',(
                        'Ambient Sensing - Field Controller',
                        'Ambient Sensing - Panel Controller',
                        'Line Sensing - Field Controller',
                        'Line Sensing - Panel Controller with Field Sensor',
                        'Uncontrolled',
                        'Other'),
                        key='control_selection_373')
            

            if submit_button:
                center_running()
                time.sleep(2)
                user_id_sub = int(user_id)
                due_date_sub = str(due_date)
                form_id_sub = str(form_id)

                temp_check = 1
                if temp_maintain != '':
                    try:
                        temp_maintain_sub = float(temp_maintain.replace(',','.'))
                    except:
                        temp_check = 0
                else:
                    temp_maintain_sub = None
                if temp_limiter != '':
                    try:
                        temp_limiter_sub = float(temp_limiter.replace(',','.'))
                    except:
                        temp_check = 0
                else:
                    temp_limiter_sub = None
                if current_op != '':
                    try:
                        current_op_sub = float(current_op.replace(',','.'))
                    except:
                        temp_check = 0
                else:
                    current_op_sub = None

                if temp_check == 0:
                    st.sidebar.warning("Temperature and Current have to be numerical!")
                else:
                    if len(eht_cct) > 0:
                        st.session_state.load_state = True
                        submit_373 = form_373(form_id_sub,user_id_sub,user_select,notes,
                                        due_date_sub, priority,project_id,auditscope,eht_cct,panel_no,iso_no,iso_dwg,
                                        pid_no,pid_dwg,sld_no,sld_dwg,pcl_no,pcl_dwg,
                                        temp_maintain_sub,temp_limiter_sub,current_op_sub,cb_no,supply_system,control_method)
                    
                        #st.write(submit_373.status_code)
                        #st.write(submit_373.content)
                        if submit_373.status_code in [200,201]:
                            st.sidebar.success("Form Dispatched!")

                            cet_timezone = pytz.timezone('Europe/Paris')
                            cet_time = datetime.datetime.now(pytz.utc).astimezone(cet_timezone)
                            
                            log_row = {'timestamp':cet_time,'language':lang_select,'dispatcher':dispatcher_select,
                                        'recipient':user_select,'project':project_id,
                                        'form':form_id,'duedate':due_date_sub,'priority':priority,'scope':auditscope,
                                        'cct':eht_cct,'panel':panel_no,'iso':iso_no,'pid':pid_no,
                                        'sld':sld_no,'pcl':pcl_no,'tm':temp_maintain_sub,'tl':temp_limiter_sub,
                                        'io':current_op_sub,'cb':cb_no,'supply':supply_system,'control':control_method,
                                        'note':notes
                                        }
                            log_df = get_csvlog('s3-nvent-prontoforms-data/Logs/NF373.csv')
                            log_df.loc[len(log_df)] = log_row
                            save_csv(log_df,'s3-nvent-prontoforms-data/Logs/NF373.csv')

                        else:
                            st.sidebar.warning("Dispatch Failed!")
                            st.sidebar.write(submit_373.content)
                            st.sidebar.write(submit_373.headers)
                    else:
                        st.sidebar.warning("Dispatch Failed!")                
                        st.warning("Please enter EHT Circuit Tag Number as a minimum!")