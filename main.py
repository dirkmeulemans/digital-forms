import streamlit as st
from streamlit_option_menu import option_menu
from PIL import Image
import collector, dispatcher, generator

image = Image.open('nVent_Logo_RGB_rev_F2.png')
new_image = image.resize((150, 100))
icon_img = Image.open('nVent_Icon_Red.png')
st.set_page_config(layout="wide",page_title="nVent Life Cycle Services",page_icon = icon_img)

hide_default_format = """
       <style>
       #MainMenu {visibility: hidden; }
       footer {visibility: hidden;}
       </style>
       """

remove_white_spaces = """
        <style>
               .block-container {
                    padding-top: 1rem;
                    padding-bottom: 0rem;
                    padding-left: 5rem;
                    padding-right: 5rem;
                }
        </style>
        """

st.markdown(hide_default_format, unsafe_allow_html=True)
st.markdown(remove_white_spaces, unsafe_allow_html=True)

image = Image.open('nVent_Logo_RGB_rev_F2.png')
new_image = image.resize((150, 100))

col_h1, col_h2 = st.columns([1,3])

with col_h1:
    st.image(new_image)

# with col_h2:
#     st.markdown("""
#             # ProntoForms Dispatcher
#             """)




# st.set_page_config(
#     page_title="nVent Life Cycle Services"
# )


def check_password():
    """Returns `True` if the user had the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False
    
    if "password_correct" not in st.session_state:
        # First run, show input for password.
        col_pass1, col_pass2, col_pass3 = st.columns([1,3,1])
        with col_pass2:
            st.text_input(
                "Password", type="password", on_change=password_entered, key="password"
            )
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        col_pass1, col_pass2, col_pass3 = st.columns([1,3,1])
        with col_pass2:
            st.text_input(
                "Password", type="password", on_change=password_entered, key="password"
            )
        col_passerr1, col_passerr2, col_passerr3 = st.columns([1,3,1])
        with col_passerr2:
            st.error("😕 Password incorrect")
        return False
    else:
        return True

if check_password():

    class MultiApp:

        def __init__(self):
            self.apps = []
        def add_app(self, title, function):
            self.apps.append({
                "title": title,
                "function": function
            })
        def run():
            #with st.sidebar:
            app = option_menu(
                menu_title = "Tools",
                options = ["Dispatcher", "Collector","Generator"],
                icons = ["send", "database-down", "pen"],
                menu_icon = "gear",
                default_index = 0, 
                orientation="horizontal",
                styles = {
                    "container": {"padding": "5!important", "background-color": "black"},
                    "icon": {"color": "white", "font-size": "23px"},
                    "nav-link": {"color":"white","font-size": "20px", "text-align": "left", "margin":"0px", "--hover-color": "grey"},
                    "nav-link-selected": {"background-color": "#ae0001"}
                }

            )

            if app == "Dispatcher":
                dispatcher.app()
            if app == "Collector":
                collector.app()
            if app == "Generator":
                generator.app()

        run()


