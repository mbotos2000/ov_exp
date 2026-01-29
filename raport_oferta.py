from __future__ import print_function
from io import BytesIO
from datetime import *
import streamlit as st
import pandas as pd
from pandas import *
from docx2python import docx2python
import os
import base64
import time
import ftplib
from mailmerge import MailMerge
from difflib import get_close_matches
import pickle
import string
def load_ftp_file():
    # Establish FTP connection
    #ftp_server = ftplib.FTP("users.utcluj.ro", st.secrets['u'], st.secrets['p'])
    ftp_server = ftplib.FTP_TLS("users.utcluj.ro")
    ftp_server.login(user=st.secrets['u'], passwd=st.secrets['p'])
    ftp_server.prot_p()
    ftp_server.encoding = "utf-8"  # Force UTF-8 encoding
    ftp_server.cwd('./public_html')

    # Download CSV files
    
    # Download DOCX templates
    docx_files = {}
    for filename in [
        "template.docx"]:
        file_data = BytesIO()
        ftp_server.retrbinary(f"RETR {filename}", file_data.write)
        file_data.seek(0)  # Reset file pointer to the start
        docx_files[filename] = file_data

    # Close FTP connection
    ftp_server.quit()

    # Return downloaded files
    return ( 
        docx_files["template.docx"]  )
# Use a session state flag to control cache invalidation

if "refresh_data" not in st.session_state:
    st.session_state.refresh_data = False

if st.button("🔄 Refresh FTP Data (apasa doar daca nu s-a actualizat baza de date!!!)"):
    st.session_state.refresh_data = True
def find_closest_match_index(word, word_list, cutoff=0.6):
    word = preprocess(word)
    word_list = [preprocess(w) for w in word_list]
    
    closest_matches = get_close_matches(word, word_list, n=1, cutoff=cutoff)
    if closest_matches:
        return word_list.index(closest_matches[0])
    return 0
	
def clean_value(value):
    if pd.isna(value):  # Replaces NaN or None with an empty string
        return ''
    elif isinstance(value, bool):  # Convert boolean values to strings
        return str(value)
    elif isinstance(value, (int, float, str)):  # Keep numbers and strings as they are
        return value
    else:
        return 'Unknown object'  # Handle unrecognized objects by converting them to a string
def fix_encoding(text):
    if isinstance(text, str):
        try:
            return text.encode('latin1').decode('utf-8')  # Fix incorrectly decoded text
        except UnicodeEncodeError:
            return text  # Return text unchanged if no encoding issues
    return text  # If it's not a string, return as is
def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'
    return href

def format_eu_number(value):
    # Convert input to integer
    n = int(value)

    # Format using Python's standard formatting
    formatted = f"{n:,.2f}"

    # Swap separators: , ↔ .
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")

    return formatted
@st.cache_data
def schimba_val_inc_nd(new):
    st.session_state['val_inc_nd'] = str(new)
@st.cache_data
def schimba_nr_contract(new):
    st.session_state['nr_contract'] = str(new)
@st.cache_data
def schimba_data_contract(new):
    st.session_state['data_contract'] = str(new)
@st.cache_data
def schimba_beneficiar(new):
    st.session_state['beneficiar'] = str(new)
@st.cache_data
def schimba_cerere(new):
    st.session_state['cerere'] = str(new)
@st.cache_data
def schimba_numec(new):
    st.session_state['numec'] = str(new)
@st.cache_data
def schimba_val_ET(new):
    st.session_state['val_ET'] = str(new)
@st.cache_data
def schimba_ore_et(new):
    st.session_state['ore_et'] = str(new)
@st.cache_data
def schimba_tarif_et(new):
    st.session_state['tarif_et'] = str(new)
@st.cache_data
def schimba_zimax_et(new):
    st.session_state['zimax_et'] = str(new)
@st.cache_data
def schimba_zimin_et(new):
    st.session_state['zimin_et'] = str(new)
@st.cache_data
def schimba_val_a_3d(new):
    st.session_state['val_a_3d'] = str(new)
@st.cache_data
def schimba_val_a_3d(new):
    st.session_state['val_a_3d'] = str(new)
@st.cache_data
def schimba_val_a_rel(new):
    st.session_state['val_a_rel'] = str(new)
@st.cache_data
def schimba_zimax_a(new):
    st.session_state['zimax_a'] = str(new)
@st.cache_data
def schimba_zimin_a(new):
    st.session_state['zimin_a'] = str(new)
@st.cache_data
def schimba_zimax_IND(new):
    st.session_state['zimax_IND'] = str(new)
@st.cache_data
def schimba_zimin_IND(new):
    st.session_state['zimin_IND'] = str(new)
@st.cache_data
def schimba_val_bet(new):
    st.session_state['val_bet'] = str(new)    
@st.cache_data
def schimba_val_geo(new):
    st.session_state['val_geo'] = str(new)

@st.cache_data
def schimba_val_geo(new):
    st.session_state['val_geo'] = str(new)  

@st.cache_data
def schimba_val_dezveliri(new):
    st.session_state['val_dezveliri'] = str(new) 


@st.cache_data
def schimba_nr_dezveliri(new):
    st.session_state['nr_dezveliri'] = str(new)
    
@st.cache_data
def schimba_zimax_geo(new):
    st.session_state['zimax_geo'] = str(new)

@st.cache_data
def schimba_zimin_geo(new):
    st.session_state['zimin_geo'] = str(new) 

@st.cache_data
def schimba_val_et_finisaje(new):
    st.session_state['val_et_finisaje'] = str(new) 

@st.cache_data
def schimba_val_rel_struct(new):
    st.session_state['val_rel_struct'] = str(new)

@st.cache_data
def schimba_val_et_actualizat(new):
    st.session_state['val_et_actualizat'] = str(new)
@st.cache_data
def schimba_zimin_rel(new):
    st.session_state['zimin_rel'] = str(new)
@st.cache_data
def schimba_zimax_et_rel(new):
    st.session_state['zimax_et_rel'] = str(new)
@st.cache_data
def schimba_termen_predare(new):
    st.session_state['zimin_et_rel'] = str(new)
@st.cache_data
def schimba_termen_predare(new):
    st.session_state['termen_predare'] = str(new)
@st.cache_data
def schimba_termen_val(new):
    st.session_state['termen_val'] = str(new)
@st.cache_data
def schimba_semnatura(new):
    st.session_state['semnatura'] = str(new)

if "step" not in st.session_state:
    st.session_state.step = 1

st.set_page_config(layout="wide", initial_sidebar_state="auto")

for key in ["val_inc_nd","nr_contract","data_contract","beneficiar","cerere","numec","val_ET","ore_et","tarif_et","zimax_et","zimin_et",
    "val_a_3d","val_a_rel","zimax_a","zimin_a","zimax_IND","zimin_IND","val_bet","val_geo","val_dezveliri","nr_dezveliri",
    "zimax_geo","zimin_geo","val_et_finisaje","val_rel_struct","val_et_actualizat","zimin_rel","zimax_et_rel","termen_predare","termen_val","semnatura"]:
    st.session_state.setdefault(key, '')
for key in ['zimax_et','zimax_a' ,'zimax_IND','zimax_geo','zimin_geo','zimin_a','zimax_et_rel',"zimin_IND",'zimin_rel','zimax_rel','zimin_et_rel',]:
    st.session_state.setdefault(key, int(60.0))
keys_none=['cap2','cap3','cap4','resetare' ,'file']
for key in keys_none:
    st.session_state.setdefault(key, None)

st.session_state['file'] = st.file_uploader("Incarca centralizatorul in excel", type="xlsx")
        
if st.session_state['file']!=None:
  if st.session_state['file']:
        df = pd.read_excel(st.session_state['file'], header=None)
        st.dataframe(df)
        st.success("Excel loaded")

  st.title("Generare oferta")
  st.write('{:%d-%b-%Y}'.format(date.today()))
  with st.form('Inregistrare cerere'):
    st.header('Inregistrare cerere')
    if st.session_state.step >= 1:
          st.write('Oferta expertiza')
          st.text_area('Numar oferta',key='nr_contract')
          d_com=st.date_input("Data ofertei",date.today())
          st.session_state['data_contract']=str(d_com)     
    if st.session_state.step >= 2:
                st.write('Date despre beneficiar si cererea depusa:')
                st.text_area('Beneficiar',key='beneficiar')
                st.text_area('Denumire contract',key='numec')
                #schimba_beneficiar(beneficiar)
                st.text_area('Numar cerere pentru care se face oferta',key='cerere')
                #schimba_cerere(cerere)
    if st.session_state.step >= 3:
                st.write('1. Expertiză tehnică')
                st.text_area('Valoare expertiza tehnica',value=str(format_eu_number(df.iloc[113, 8])), key='val_ET')
                #schimba_val_ET(format_eu_number(a))
                st.text_area('Numar ore necesar verificare',key='ore_et')
                st.text_area('Tarif verificare',key='tarif_et')           
                st.selectbox('Durata de realizare a expertizei tehnice: ',range(1, 60),key='zimax_et')
                st.write('Numai putin de:')
                st.selectbox('Nu mai putin de: ',range(1, 60),key='zimin_et')
    if st.session_state.step >= 4:
                st.text_area('2.1 Scan 3D și generare nor de puncte: ',value=str(format_eu_number(df.iloc[115, 8])), key='val_a_3d')
                st.text_area('2.2 Elaborare releveu arhitectural al construcției : ',value=str(format_eu_number(df.iloc[113, 8])), key='val_a_rel')       
                st.selectbox('Durata de realizare a releveului: ',range(1, 60),key='zimax_a')
                st.selectbox('Nu mai putin de: ',range(1, 60),key='zimin_a')
    if st.session_state.step >= 5:
                st.write('3. Investigații prin încercări nedistructive la elementele structurale în vederea determinării modului de alcătuire și armare ')
                st.text_area('3. Investigații prin încercări nedistructive : ',value=str(format_eu_number(df.iloc[115, 8])), key='val_inc_nd')
                st.selectbox('Durata de realizare a releveului: ',range(1, 60), index=25,key='zimax_IND')
                st.selectbox('Nu mai putin de: ',range(1, 60),index=25,key='zimin_IND')
    if st.session_state.step >= 6:
                st.write('4. Teste pe betonul pus în operă prin extragere și testare carote ')
                st.text_area('4. Teste pe betonul pus în operă  : ',value=str(format_eu_number(df.iloc[118, 8])), key='val_bet')
    if st.session_state.step >= 7:
                st.write('5. Studiu Geotehnic și dezveliri la nivelul fundațiilor')
                st.text_area(' Studiu Geotehnic : ',value=str(format_eu_number(df.iloc[119, 8])), key='val_geo') 
                st.text_area(' Dezveliri : ',value=str(format_eu_number(df.iloc[119, 8])), key='val_dezveliri')
                st.selectbox('Numarul minim de dezveliri: ',range(1, 60),index=8, key='nr_dezveliri')
                st.selectbox('Durata de realizare a studiului geotehnic: ',range(1, 60),index=30, key='zimax_geo')
                st.write('Numai putin de:')
                st.selectbox('Nu mai putin de: ',range(1, 60),index=25,key='zimin_geo')
    if st.session_state.step >= 8:
                st.text_area(' Realizare lucrări de decopertare finisaje interioare : ',value=str(format_eu_number(df.iloc[121, 8])), key='val_et_finisaje') 
                st.text_area(' Elaborare releveu structural al construcției : ',value=str(format_eu_number(df.iloc[116, 8])), key='val_rel_struct') 
                st.text_area(' Actualizare expertiză tehnică   : ',value=str(format_eu_number(df.iloc[122, 4])), key='val_et_actualizat') 
                schimba_val_a_rel(format_eu_number(df.iloc[115, 9]))
                st.selectbox('Durata de realizare a releveului structural este de maxim: ',range(1, 60),index=30, key='zimax_rel')
                st.selectbox('Nu mai putin de: ',range(1, 60),index=25,key='zimin_rel')          
                st.selectbox('Durata de realizare a actualizării expertizei tehnice : ',range(1, 60),index=30, key='zimax_et_rel')
                st.selectbox('Nu mai putin de: ',range(1, 60),index=25,key='zimin_et_rel')
                st.selectbox('Termen predare: ',range(1, 60),index=20, key='termen_predare')
                st.selectbox('Termen valabilitate',range(1, 60),index=8, key='termen_val')
    if st.session_state.step >= 9:	
      template=load_ftp_file()
      keys_to_merge=["val_inc_nd",
    "nr_contract",
    "data_contract",
    "beneficiar",
    "cerere",
    "numec",
    "val_ET",
    "ore_et",
    "tarif_et",
    "zimax_et",
    "zimin_et",
    "val_a_3d",
    "val_a_rel",
    "zimax_a",
    "zimin_a",
    "zimax_IND",
    "zimin_IND",
    "val_bet",
    "val_geo",
    "val_dezveliri",
    "nr_dezveliri",
    "zimax_geo",
    "zimin_geo",
    "val_et_finisaje",
    "val_rel_struct",
    "val_et_actualizat",
    "zimin_rel",
    "zimax_et_rel",
    "termen_predare",
    "termen_val",
    "semnatura"]

      document=MailMerge(template)
        #st.write(document.get_merge_fields())
      for key in keys_to_merge:
                    document.merge(**{key: st.session_state[key]})
      document.write("oferta.docx")
      st.markdown(get_binary_file_downloader_html("oferta.docx", 'Word document'), unsafe_allow_html=True)
    submitted = st.form_submit_button("Next")

 # Logic AFTER the form
  if submitted:
    st.session_state.step += 1

        





