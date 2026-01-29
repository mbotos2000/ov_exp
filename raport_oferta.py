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



st.set_page_config(layout="wide", initial_sidebar_state="auto")

if 'val_inc_nd' not in st.session_state:
    st.session_state['val_inc_nd']=''
if 'nr_contract' not in st.session_state:
    st.session_state['nr_contract']=''
if 'data_contract' not in st.session_state:
    st.session_state['data_contract']=''
if 'beneficiar' not in st.session_state:
    st.session_state['beneficiar']=''
if 'cerere' not in st.session_state:
    st.session_state['cerere']=''    
if 'nume_contract' not in st.session_state:
    st.session_state['nume_contract']=''                
if 'ore_et' not in st.session_state:
    st.session_state['ore_et']=''

if 'val_ET' not in st.session_state:
    st.session_state['val_ET']=''
if 'tarif_et' not in st.session_state:
    st.session_state['tarif_et']=''
if 'zimax_et' not in st.session_state:
    st.session_state['zimax_et']=''
if 'zimin_et' not in st.session_state:
    st.session_state['zimin_et']=''
if 'val_a_3d' not in st.session_state:
    st.session_state['val_a_3d']=''
if 'val_a_rel' not in st.session_state:
    st.session_state['val_a_rel']=''
if 'zimax_a' not in st.session_state:
    st.session_state['zimax_a']=''
if 'zimin_a' not in st.session_state:
    st.session_state['zimin_a']=''
if 'zimax_IND' not in st.session_state:
    st.session_state['zimax_IND']=''
if 'zimin_IND' not in st.session_state:
    st.session_state['zimin_IND']=''
if 'val_bet' not in st.session_state:
    st.session_state['val_bet']=''
if 'val_geo' not in st.session_state:
    st.session_state['val_geo']='-'

if 'val_dezveliri' not in st.session_state:
    st.session_state['val_dezveliri']='-'
    
 
if 'nr_dezveliri' not in st.session_state:
    st.session_state['nr_dezveliri']='-'
 
if 'zimax_geo' not in st.session_state:
    st.session_state['zimax_geo']='-'

if 'zimin_geo' not in st.session_state:
    st.session_state['zimin_geo']='-'

if 'val_et_finisaje' not in st.session_state:
    st.session_state['val_et_finisaje']='-'

if 'val_rel_struct' not in st.session_state:
    st.session_state['val_rel_struct']='-'
 
 
if 'val_et_actualizat' not in st.session_state:
    st.session_state['val_et_actualizat']=''
if 'zimin_rel' not in st.session_state:
    st.session_state['zimin_rel']=''
if 'zimax_et_rel' not in st.session_state:
    st.session_state['zimax_et_rel']=''
if 'zimin_et_rel' not in st.session_state:
    st.session_state['zimin_et_rel']=''
if 'termen_predare' not in st.session_state:
    st.session_state['termen_predare']=''
if 'termen_val' not in st.session_state:
    st.session_state['termen_val']=''
if 'semnatura' not in st.session_state:
    st.session_state['semnatura']=''



if 'file' not in st.session_state:
    st.session_state['file']=None




st.session_state['file'] = st.file_uploader("Incarca centralizatorul in excel", type="xlsx")
        
if st.session_state['file']!=None:
  excel_data = {}

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
  with st.form('Oferta expertiza'):
    st.header('Inregistrare cerere')
    submitted = st.form_submit_button("Treceti la alegerea specializarii")
    if submitted:
      st.text_area('Numar oferta',key='Nume_contract')
      d_com=st.date_input("Data ofertei",date.today())
      st.session_state['data_contract']=str(d_com)
      st.session_state['cap2']='1'     
  if st.session_state['cap2']!=None:
    with st.form('Date despre beneficiar si cererea depusa:'):
        st.text_area('Beneficiar',key='beneficiar')
        st.text_area('Numar cerere pentru care se face oferta',key='cerere')
        st.session_state['cap2']='2'
  if st.session_state['cap2']=='2':
    with st.form('1. Expretiza tehnica'):

        

                st.session_state['cap3']='1'
 

  if st.session_state['cap3']!=None:
    st.write('Distribuția fondului de timp (ore pe semestru)')
    #st.session_state['M_3_8']=str(data1['orestud'].loc[(data1['specializare']==st.session_state['M_1_6'])&(data1['nume_disciplina']==st.session_state['M_2_1']) & (data1['curs']=='CURS      ')].values[0])
    tosi=38

    #st.write('Total ore studiu individual ', tosi)

    slide_37a=0
    slide_37b=0
    slide_37c=0
    slide_37d=0
    slide_37e=0
    slide_37f=0
    st.write('Distribuția fondului de timp:')
    slide_37a=st.slider(
      '(a) Studiul după manual, suport de curs, bibliografie şi notițe',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
     
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    slide_37b=st.slider(
      '(b) Documentare suplimentară în bibliotecă, pe platforme electronice de specialitate şi pe teren',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    slide_37c=st.slider(
      'c) Pregătire seminarii / laboratoare, teme, referate, portofolii şi eseuri',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    slide_37d=st.slider(
      '(d) Tutoriat',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
    slide_37e=st.slider(
      'e) Examinări',
      min_value=0, max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
      help='Completati de la a spre f. Suma orelor de studiu individual este blocata pe valoarea din planurile de invatamant')
    sd=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f))
    if not(sd<=0):
        slide_37f=st.slider(
          '(f) Alte activități:',
           max_value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
      
          value=int(tosi-int(slide_37a)-int(slide_37b)-int(slide_37c)-int(slide_37d)-int(slide_37e)-int(slide_37f)),
          help='Completati de la a spre f. Suma orelor de studiu individual este cea din planurile de invatamant')
    else:
            st.write('(f) Alte activități: 0 ore')
            slide_37f=0
            slide_37e+=-sd
    a=st.button('Treci la capitolul 4')
    if a:
      st.write('Capitolul 4')
      schimba_M_3_7_a(slide_37a)
      schimba_M_3_7_b(slide_37b)
      schimba_M_3_7_c(slide_37c)
      schimba_M_3_7_d(slide_37d)
      schimba_M_3_7_e(slide_37e)
      schimba_M_3_7_f(slide_37f)
      
      st.session_state['cap4']='1'
    

  if st.session_state['cap4']!=None:
    with st.form('capitolul 4'):
      
      d_dep='04.09.2022'
      d_fac='21.09.2022'
      submitted= st.form_submit_button("finalizeaza")
      if submitted:
       
    
        document = MailMerge(template)
        #st.write(document.get_merge_fields())
        document.merge(da_cu=st.session_state['d_com'])
        document.merge(M_8_2_14=st.session_state['M_8_2_14'])
        document.merge(M_2_2_1=st.session_state['M_2_2_1'])
        document.merge(M_8_2_1=st.session_state['M_8_2_1'])
        document.merge(M_3_3_p=st.session_state['M_3_3_p'])
        document.merge(M_8_1_14=st.session_state['M_8_1_14'])
        document.merge(M_8_2_9=st.session_state['M_8_2_9'])
        document.merge(M_8_1_o1=st.session_state['M_8_1_o1'])
        document.merge(M_8_1_mp=st.session_state['M_8_1_mp'])
        document.merge(M_8_1_mp1=st.session_state['M_8_1_mp1'])
        document.merge(M_8_1_1=st.session_state['M_8_1_1'])
        document.merge(M_3_3_s=st.session_state['M_3_3_s'])
        document.merge(M_7_2=st.session_state['M_7_2'])
        document.merge(data_dep=st.session_state['data_dep'])
        document.merge(tip=st.session_state['tip'])
        document.merge(dir_dep=st.session_state['dir_dep'])
        document.merge(M_8_2_5=st.session_state['M_8_2_5'])
        document.merge(M_3_7_e=st.session_state['M_3_7_e'])
        document.merge(M_2_1=st.session_state['M_2_1'])
        document.merge(M_10_2_c=st.session_state['M_10_2_c'])
        document.merge(M_8_1_12=st.session_state['M_8_1_12'])
        document.merge(M_1_2=st.session_state['M_1_2'])
        document.merge(M_10_6=st.session_state['M_10_6'])
        document.merge(M_9=st.session_state['M_9'])
        document.merge(M_2_3=st.session_state['M_2_3'])
        document.merge(M_10_3_a=st.session_state['M_10_3_a'])
        document.merge(M_1_1=st.session_state['M_1_1'])
        document.merge(M_8_1_13=st.session_state['M_8_1_13'])
        document.merge(M_3_4=st.session_state['M_3_4'])
        document.merge(M_3_3_l=st.session_state['M_3_3_l'])
        document.merge(M_8_1_5=st.session_state['M_8_1_5'])
        document.merge(M_8_2_6=st.session_state['M_8_2_6'])
        document.merge(M_3_5=st.session_state['M_3_5'])
        document.merge(M_4_2=st.session_state['M_4_2'])
        document.merge(da_cu=st.session_state['da_cu'])
        document.merge(M_8_2_7=st.session_state['M_8_2_7'])
        document.merge(M_8_2_2=st.session_state['M_8_2_2'])
        document.merge(M_8_2_8=st.session_state['M_8_2_8'])
        document.merge(M_3_2=st.session_state['M_3_2'])
        document.merge(M_10_3_c=st.session_state['M_10_3_c'])
        document.merge(M_3_6_l=st.session_state['M_3_6_l'])
        document.merge(M_1_8=st.session_state['M_1_8'])
        document.merge(M_10_2_a=st.session_state['M_10_2_a'])
        document.merge(decan=st.session_state['decan'])
        document.merge(M_8_1_10=st.session_state['M_8_1_10'])
        document.merge(Biblio_c=st.session_state['Biblio_c'])
        document.merge(M_4_1=st.session_state['M_4_1'])
        document.merge(M_7_1=st.session_state['M_7_1'])
        document.merge(fac=st.session_state['fac'])
        document.merge(M_3_7_f=st.session_state['M_3_7_f'])
        document.merge(M_2_5=st.session_state['M_2_5'])
        document.merge(M_8_1_8=st.session_state['M_8_1_8'])
        document.merge(M_3_7_b=st.session_state['M_3_7_b'])
        document.merge(M_3_7_a=st.session_state['M_3_7_a'])
        document.merge(M_2_2=st.session_state['M_2_2'])
        document.merge(M_5_2=st.session_state['M_5_2'])
        document.merge(M_8_1_4=st.session_state['M_8_1_4'])
        document.merge(M_2_7_1=st.session_state['M_2_7_1'])
        document.merge(M_8_1_7=st.session_state['M_8_1_7'])
        document.merge(M_8_2_3=st.session_state['M_8_2_3'])
        document.merge(M_3_7_d=st.session_state['M_3_7_d'])
        document.merge(M_8_2_12=st.session_state['M_8_2_12'])
        document.merge(M_3_9=st.session_state['M_3_9'])
        document.merge(M_3_7_c=st.session_state['M_3_7_c'])
        document.merge(M_6_ct=st.session_state['M_6_ct'])
        document.merge(M_8_1_2=st.session_state['M_8_1_2'])
        document.merge(M_8_1_3=st.session_state['M_8_1_3'])
        document.merge(dep=st.session_state['dep'])
        document.merge(M_3_6_p=st.session_state['M_3_6_p'])
        document.merge(M_10_1_a=st.session_state['M_10_1_a'])
        document.merge(M_2_4=st.session_state['M_2_4'])
        document.merge(M_2_6=st.session_state['M_2_6'])
        document.merge(Biblio_a=st.session_state['Biblio_a'])
        document.merge(data_fac=st.session_state['data_fac'])
        document.merge(M_8_1_o=st.session_state['M_8_1_o'])
        document.merge(M_1_6=st.session_state['M_1_6'])
        document.merge(M_3_1=st.session_state['M_3_1'])
        document.merge(M_6_cp=st.session_state['M_6_cp'])
        document.merge(M_3_6_s=st.session_state['M_3_6_s'])
        document.merge(M_1_4=st.session_state['M_1_4'])
        document.merge(M_5_1=st.session_state['M_5_1'])
        document.merge(M_8_1_6=st.session_state['M_8_1_6'])
        document.merge(M_8_2_4=st.session_state['M_8_2_4'])
        document.merge(M_8_2_13=st.session_state['M_8_2_13'])
        document.merge(M_8_2_10=st.session_state['M_8_2_10'])
        document.merge(M_2_7_2=st.session_state['M_2_7_2'])
        document.merge(M_8_1_9=st.session_state['M_8_1_9'])
        document.merge(M_1_3=st.session_state['M_1_3'])
        document.merge(M_1_5=st.session_state['M_1_5'])
        document.merge(M_8_1_11=st.session_state['M_8_1_11'])
        document.merge(M_10_1_c=st.session_state['M_10_1_c'])
        document.merge(M_3_8=st.session_state['M_3_8'])
        document.merge(M_2_3_1=st.session_state['M_2_3_1'])
        document.merge(M_3_11=st.session_state['M_3_11'])
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
        





