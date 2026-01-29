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
def schimba_nume_contract(new):
    st.session_state['nume_contract'] = str(new)
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

for key in ['val_inc_nd','nr_contract','nr_contract','data_contract','beneficiar','cerere',
            'nume_contract','ore_et','val_ET','tarif_et','zimax_et' ,'zimin_et','val_a_3d' ,'val_a_rel', 'val_bet','val_geo','nr_dezveliri',
            'val_et_finisaje','val_rel_struct','val_et_actualizat','termen_predare','termen_val','semnatura']:
    st.session_state.setdefault(key, '')
for key in ['zimax_et','zimax_a' ,'zimax_IND','zimax_geo','zimin_geo','zimax_et_rel']:
    st.session_state.setdefault(key, int(60.0))
keys_none=['cap2','cap3','cap4','resetare' ,'file']
for key in keys_none:
    st.session_state.setdefault(key, None)

st.session_state['file'] = st.file_uploader("Incarca centralizatorul in excel", type="xlsx")
        
if st.session_state['file']!=None:
  if st.session_state['file']:
        df = pd.read_excel(st.session_state['file'], header=None)
        st.dataframe(df)
        schimba_val_ET(df.iloc[113, 8])
        schimba_val_a_3d(df.iloc[115, 8])
        schimba_val_a_rel(df.iloc[115, 9])
        schimba_val_inc_nd(df.iloc[115, 8])
        schimba_val_bet(df.iloc[118, 8])
        schimba_val_et_actualizat(df.iloc[122, 4])
        schimba_val_et_finisaje(df.iloc[121,4])
        st.success("Excel loaded")

  st.title("Generare oferta")
  st.write('{:%d-%b-%Y}'.format(date.today()))
  with st.form('Inregistrare cerere'):
    st.header('Inregistrare cerere')
    if st.session_state.step >= 1:
          st.write('Oferta expertiza')
          st.text_area('Numar oferta',key='Nume_contract')
          d_com=st.date_input("Data ofertei",date.today())
          st.session_state['data_contract']=str(d_com)     
    if st.session_state.step >= 2:
                st.write('Date despre beneficiar si cererea depusa:')
                st.text_area('Beneficiar',key='beneficiar')
                st.text_area('Numar cerere pentru care se face oferta',key='cerere')
    if st.session_state.step >= 3:
                st.write('1. Expertiză tehnică')
                st.text_area('Valoare expertiza tehnica',value=str(df.iloc[113, 8]), key='val_et')
                st.text_area('Numar ore necesar verificare',key='ore_et')
                st.text_area('Tarif verificare verificare',key='tarif_et')           
                st.selectbox('Durata de realizare a expertizei tehnice: ',
                    range(1, 60),key='zimax_et')
                st.write('Numai putin de:')
                st.selectbox('Nu mai putin de: ',
                    range(1, 60),key='zimin_et')
    #a=st.button('Treci la capitolul 4')
    if st.session_state.step >= 4:
                st.write('2.Scanare 3D de înaltă precizie a construcției și elaborare releveu arhitectural al acesteia')
                st.text_area('2.1 Scan 3D și generare nor de puncte: ',value=str(df.iloc[113, 8]), key='val_a_3d')
                st.text_area('2.2 Elaborare releveu arhitectural al construcției : ',value=str(df.iloc[113, 8]), key='val_a_rel')       
                st.selectbox(
                    'Durata de realizare a releveului: ',
                    range(1, 60),key='zimax_a')
                st.write('Numai putin de:')
                st.selectbox(
                    'Nu mai putin de: ',
                    range(1, 60),key='zimin_a')
    if st.session_state.step >= 5:
                st.write('3. Investigații prin încercări nedistructive la elementele structurale în vederea determinării modului de alcătuire și armare ')
                st.text_area('3. Investigații prin încercări nedistructive : ',value=str(df.iloc[113, 8]), key='val_inc_nd')
                st.selectbox(
                    'Durata de realizare a releveului: ',
                    range(1, 60),value=30, key='zimax_IND')
                st.write('Numai putin de:')
                st.selectbox(
                    'Nu mai putin de: ',
                    range(1, 60),value=25,key='zimin_IND')
    if st.session_state.step >= 6:
                st.write('4.	Teste pe betonul pus în operă prin extragere și testare carote ')
                st.text_area('4.	Teste pe betonul pus în operă  : ',value=str(df.iloc[113, 8]), key='val_bet')
    if st.session_state.step >= 7:
                st.write('5.	Studiu Geotehnic și dezveliri la nivelul fundațiilor')
                st.text_area(' Studiu Geotehnic : ',value=str(df.iloc[113, 8]), key='val_geo') 
                st.text_area(' Dezveliri : ',value=str(df.iloc[113, 8]), key='val_dezveliri')
                st.selectbox(
                    'Numarul minim de dezveliri: ',
                    range(1, 60),value=8, key='nr_dezveliri')
                st.selectbox(
                    'Durata de realizare a studiului geotehnic: ',
                    range(1, 60),value=30, key='zimax_geo')
                st.write('Numai putin de:')
                st.selectbox(
                    'Nu mai putin de: ',
                    range(1, 60),value=25,key='zimin_geo')
    if st.session_state.step >= 8:
                st.text_area(' Realizare lucrări de decopertare finisaje interioare  : ',value=str(df.iloc[113, 8]), key='val_et_finisaje') 
                st.text_area(' Elaborare releveu structural al construcției   : ',value=str(df.iloc[113, 8]), key='val_rel_struct') 
                st.text_area(' Actualizare expertiză tehnică   : ',value=str(df.iloc[113, 8]), key='val_et_actualizat') 
  
                st.selectbox(
                    'Durata de realizare a releveului structural este de maxim: ',
                    range(1, 60),value=30, key='zimax_rel')
                st.write('Numai putin de:')
                st.selectbox(
                    'Nu mai putin de: ',
                    range(1, 60),value=25,key='zimin_rel')
                
                st.selectbox(
                    'Durata de realizare a actualizării expertizei tehnice : ',
                    range(1, 60),value=30, key='zimax_et_rel')
                st.write('Numai putin de:')
                st.selectbox(
                    'Nu mai putin de: ',
                    range(1, 60),value=25,key='zimin_et_rel')
                st.selectbox(
                    'Termen predare: ',
                    range(1, 60),value=20, key='termen_predare')
                st.selectbox(
                    'Termen valabilitate',
                    range(1, 60),value=8, key='termen_val')

    submitted = st.form_submit_button("Next")

 # Logic AFTER the form
  if submitted:
    st.session_state.step += 1


    

  if st.session_state['cap4']!=None:
    with st.form('capitolul 4'):
      
      d_dep='04.09.2022'
      d_fac='21.09.2022'
    submitted= st.form_submit_button("finalizeaza")
    if submitted:
       
    
        document = MailMerge(template)
        #st.write(document.get_merge_fields())
        if key in st.session_state:
                    document.merge(**{key: st.session_state[key]})
        
        #st.write(st.session_state)
        file_name=st.session_state['M_1_8']+'_FD_an'+st.session_state['M_2_4']+'_s'+st.session_state['M_2_5']+'_'+pres[st.session_state['M_1_6']]+'_'+st.session_state['M_2_1']+'_24-23.docx'
        #st.write(st.session_state['M_1_6'])
       # try:
        current_datetime = datetime.now()    
        document.write(file_name)
        st.markdown(get_binary_file_downloader_html(file_name, 'Word document'), unsafe_allow_html=True)
        st.session_state['denumirefisa']=file_name
        st.session_state['dataintocmire']=str(current_datetime)
        #os.startfile(file_name)
        





