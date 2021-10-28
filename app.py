import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pandas as pd
import numpy as np
import re
from datetime import date
import streamlit as st
import json
import glob
from PIL import Image
import smtplib
import docx2pdf 
import shutil
import zipfile
from datetime import datetime
import platform
def User_validation():
    
    
    
        
    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    
    now=datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M")
    time_past=datetime.strptime(past['Acceso']["Hora"], "%d/%m/%Y %H:%M")
    
    timesince = now - time_past
    Time_min= int(timesince.total_seconds() / 60)
    
    
   
    bool_negate = Time_min<120
  
    if not bool_negate:
        past['Acceso'].update({"Estado":"Negado"})
        str_json_0=json.dumps(past, indent=4)
        J_f=open("Validation/Validation.json","w")
        J_f.write(str_json_0)
        J_f.close()
    
    

    
    
    
    
    
    bool_aprove= past['Acceso']["Estado"]=="Aprovado"
    
    if not bool_aprove:
        
        colums= st.columns([1,2,1])
        
        with colums[1]:
            #st.image("Imagenes/Escudo_unal.png")
            st.subheader("Ingrese el usuario y contraseña")
            Usuario=st.text_input("Usuario")
            Clave=st.text_input("Contraseña",type="password")
            
        Users=["Juan Pablo"]
        bool_user = Usuario in Users
        bool_clave = (Clave)==("1234")
    

        bool_user_email = past['Acceso']["User"] == Usuario
    
        bool_time2 = Time_min<1000
        
        bool_1 =  bool_time2 and bool_user_email
        
        
        bool_2 = bool_user and bool_clave
        
        
        if not bool_user_email and  bool_2:
            
            past['Acceso'].update({"User":Usuario,"Estado":"Aprovado","Hora":dt_string})
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
            
        
        
        if not bool_2:
            if (Usuario != "") and (Clave!=""):
                with colums[1]:
                    st.warning("Usuario o contraseña incorrectos.\n\n Por favor intente nuevamente.")
            
        elif bool_2 and not bool_1:
            
            past['Acceso'].update({"User":Usuario,"Estado":"Aprovado","Hora":dt_string})
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
    
            
    
            EMAIL_ADDRESS = 'ASPproyectos2021@gmail.com'
            EMAIL_PASSWORD = 'Samuelmamahuevo2021'
            
            
        
            try:
                
                with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                    smtp.ehlo()
                    smtp.starttls()
                    smtp.ehlo()
                    
                    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                    
                    
                    
                    subject = 'Acceso aplicacion Julia' 
                    body = 'Acceso usuario ' + Usuario +' el '+dt_string 
                    
                    msg = f'Subject: {subject}\n\n{body}'
                    smtp.sendmail(EMAIL_ADDRESS, EMAIL_ADDRESS, msg)
            except:
                pass
            with colums[1]:   
                st.button("Acceder a la aplicación")
        elif bool_2:
            
            past['Acceso'].update({"Estado":"Aprovado","Hora":dt_string,"User":Usuario})
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
            

            
            with colums[1]:   
                st.button("Acceder a la aplicación")
                
    return bool_aprove
def Num_dias(leng):
    if leng==1:
        return "1 día"
    else:
        return str(leng) + " días"
def day_week(dia):
    if dia ==0:
        Dia="Lunes"
    elif dia ==1:
        Dia="Martes"
    elif dia ==2:
        Dia="Miércoles"
    elif dia ==3:
        Dia="Jueves"
    elif dia ==4:
        Dia="Viernes"
    elif dia ==5:
        Dia="Sábado"
    elif dia ==6:
        Dia="Domingo-Festivo"
        
    return Dia
def remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)        
 
def Range_fecha(dates):
   
    if len(dates)==1:
        return pd.to_datetime(dates[0]).strftime('%Y-%m-%d')
    else:
        return pd.to_datetime(dates[0]).strftime('%Y-%m-%d')+" hasta "+ pd.to_datetime(dates[-1]).strftime('%Y-%m-%d')

        
    
def dt_fechas(data,data_user,Fechas):
    dt_Final=pd.DataFrame(columns=["Dia","Fecha","Requerimiento","Respaldo"])
    for dia in Fechas:
        data_fecha=data_user[data_user["Fecha"]== dia]
        data_dia_todos=data[data["Fecha"]==dia]
        d_week=pd.Series(data=dia).dt.dayofweek.to_numpy()[0]
        df=pd.DataFrame([[day_week(d_week),dia,data_dia_todos["CANTIDAD"].sum(),data_fecha["CANTIDAD"].sum()]],columns=["Dia","Fecha","Requerimiento","Respaldo"])
        dt_Final=dt_Final.append(df, ignore_index=True)
    
    return dt_Final
        

def Mes_espa(mes):
    
    if mes =="01":
        Mes="Enero"
    elif mes =="02":
        Mes="Febrero"
    elif mes =="03":
        Mes="Marzo"
    elif mes =="04":
        Mes="Abril"
    elif mes =="05":
        Mes="Mayo"
    elif mes =="06":
        Mes="Junio"
    elif mes =="07":
        Mes="Julio"
    elif mes =="08":
        Mes="Agosto"
    elif mes =="09":
        Mes="Septiembre"
    elif mes =="10":
        Mes="Octubre"
    elif mes =="11":
        Mes="Noviembre"
    elif mes =="12":
        Mes="Diciembre"
        
    return Mes    

def F_Liq_pag(mes,ano):
    
    if mes%12 ==1:
        Fecha ="Enero"
        
    elif mes%12 ==2:
        Fecha ="Febrero"
        
    elif mes%12 ==3:
        Fecha ="Marzo"
        
    elif mes%12 ==4:
        Fecha ="Abril"
        
    elif mes%12 ==5:
        Fecha ="Mayo"
        
    elif mes%12 ==6:
        Fecha ="Junio"
        
    elif mes%12 ==7:
        Fecha ="Julio"
        
    elif mes%12 ==8:
        Fecha="Agosto"
        
    elif mes%12 ==9:
        Fecha="Septiembre"
        
    elif mes%12 ==10:
        Fecha="Octubre"
        
    elif mes%12 ==11:
        Fecha="Noviembre"
        
    elif mes%12 ==0:
        Fecha="Diciembre"
    
    if mes > 12:
        Fecha += " "+ str(ano+1)
    else:
        Fecha += " "+ str(ano)
    return Fecha       
    
        
    
def mes_espa(mes):
    
    if mes =="01":
        Mes="enero"
    elif mes =="02":
        Mes="febrero"
    elif mes =="03":
        Mes="marzo"
    elif mes =="04":
        Mes="abril"
    elif mes =="05":
        Mes="mayo"
    elif mes =="06":
        Mes="junio"
    elif mes =="07":
        Mes="julio"
    elif mes =="08":
        Mes="agosto"
    elif mes =="09":
        Mes="septiembre"
    elif mes =="10":
        Mes="octubre"
    elif mes =="11":
        Mes="noviembre"
    elif mes =="12":
        Mes="diciembre"
    
    
    return Mes

def dia_esp(dia):
    
    if dia =="01":
        Dia="1"
    elif dia =="02":
        Dia="2"
    elif dia =="03":
        Dia="3"
    elif dia =="04":
        Dia="4"
    elif dia =="05":
        Dia="5"
    elif dia =="06":
        Dia="6"
    elif dia =="07":
        Dia="7"
    elif dia =="08":
        Dia="8"
    elif dia =="09":
        Dia="9"
    else :
        Dia = dia
 
    
    
    return Dia

def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)
              

if User_validation():
    
    
    Opciones1=("Proyecto 1","Proyecto 2","Proyecto 3")
    eleccion=st.sidebar.selectbox('Seleccione el proyecto',Opciones1)
    
    
    if eleccion==Opciones1[0]:
        
        
        
        colums= st.columns([1,1,1])
        with colums[0]:                
            uploaded_file_1 = st.file_uploader("Suba el consolidado base")
        with colums[1]:                
            uploaded_file_2 = st.file_uploader("Suba la plantilla del documento")
        with colums[2]:                
            uploaded_file_3 = st.file_uploader("Suba el excel adicional")
        if (uploaded_file_1 is not None) and (uploaded_file_2 is not None) and (uploaded_file_3 is not None):
                    
            data=pd.read_excel(uploaded_file_1)
            data["Fecha"]=data["FECHAINI"].dt.to_pydatetime()
            
            
            
            if data["USUARIO"].isnull().values.any():
                st.warning("Revisar archivo de consolidado base, usuario no encontrado.")   
                data.dropna(subset = ["USUARIO"], inplace=True)
                Users=pd.unique(data["USUARIO"])
            else:
                Users=pd.unique(data["USUARIO"])
        
            
            Extras=pd.read_excel(uploaded_file_3,sheet_name="Usuarios")
            
            template_file_path = uploaded_file_2
            
            today =  date.today()
            fecha=dia_esp(today.strftime("%d")) +" de "+ mes_espa(today.strftime("%m")) +" de "+ today.strftime("%Y")
            
        
            colums= st.columns([1,4,1])
            with colums[1]:
                
                st.subheader("Introducción de las variables")
                
                P_bolsa=st.text_input("Introduzca el Precio de Escasez de Activación",value="10.00")
                P_contrato=st.text_input("Introduzca el precio del contrato [USD]",value="10.00")
                P_TMR=st.text_input("Introduzca el valor de la TRM",value="3,950.00")
                F_TRM = st.date_input("Seleccione la fecha del valor de la TRM:",value=today).strftime("%Y-%m-%d")
                
            columns_2 = st.columns([1,2,2,1])
            
            Opciones2=("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
            Opciones3=("I","II","III","IV","V")
            
            with columns_2[1]:
                eleccion2=st.selectbox('Seleccione el mes de la OFR',Opciones2)
            with columns_2[2]:
                eleccion3=st.selectbox('Selecciona la semana de la OFR',Opciones3)
                
            columns_3 = st.columns([2,1,2])

            with columns_3[1]:
                if platform.system()=='Windows':
                    b=st.checkbox("PDF")
                else:
                    
                    b=False
                a=st.button("Crear los documentos")
                
            Ruta="Documentos/"+eleccion2 +"/"+ eleccion3
            Ruta_x="Documentos/"
        
            if a:
                
                try:
                    path1 = os.path.join(Ruta)
                    shutil.rmtree(path1)
                    os.makedirs(Ruta, exist_ok=True)
                except:
                    os.makedirs(Ruta, exist_ok=True)
                
                
                Ruta_word=Ruta+"/Word"
                Ruta_pdf=Ruta+"/PDF"    
                
                Info ={"Ruta": Ruta, 
                      "File_names": None
                    } 
                   
                File_names=[]
                           
                os.makedirs(Ruta_word, exist_ok=True)
                if b:
                    
                    os.makedirs(Ruta_pdf, exist_ok=True)
                
                zf = zipfile.ZipFile(
                    "Resultado.zip", "w", zipfile.ZIP_DEFLATED)
                my_bar=st.progress(0)
                steps=len(Users)
                steps_done=0
            
                for usuario in Users:
                    
                    data_user=data.copy()
                    data_user=data_user[data_user["USUARIO"]== usuario] 
                    Empresas = pd.unique(data_user["agente1"])
                    
                    Respaldo = data[data["USUARIO"]== usuario]["CANTIDAD"].sum()
                    Fechas = pd.unique(data_user["Fecha"])
                
                    R_fechas = Range_fecha(Fechas)
                    
                    Data_frame_fechas=dt_fechas(data.copy(),data_user,Fechas)
                    
                    
                    try:
                        Email = str(Extras[Extras["USUARIO"] == usuario]["CORREO"].values)
                        Porc_come = Extras[Extras["USUARIO"] == usuario]["MARGEN"].values[0]
                    except:

                        Email = ""
                        Porc_come = 0.1
                        st.warning("No hay coincidencia en el Excel de usuarios para: "+usuario)   
                        
                    Email = re.sub("\[|\]|\'|0","",Email)
                    
                    
                    tx_empresas=""
                    for idx ,val in enumerate(Empresas):
                        
                        if len(Empresas)<4:
                            val_2=val[0:3]
                            tx_empresas += val_2
                            if idx==len(Empresas)-1:
                                pass
                            else:
                                tx_empresas +=", "
                            
                        else:
                            tx_empresas += "Los Generadores"
                        
                    P_kwh=float(re.sub(",","",P_TMR))*float(P_contrato)/1000
                    Ingreso=int(P_kwh*Respaldo)
                    C_comer=int(Ingreso*Porc_come)
                    C_GMS=int(Ingreso*4/1000)
                    I_NETO=Ingreso-C_comer-C_GMS
                    
                    if len(Data_frame_fechas.index.values)>13:
                        Enter="\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"
                    else:
                        Enter=""
                        
                    variables = {
                        "${FECHA}": fecha,
                        "${MES}": eleccion2,
                        "${AGENTES}": tx_empresas,
                        "${USUARIO}": usuario,
                        "${PRECIO_BOLSA}": P_bolsa,
                        "${PRECIO_CONTRATO}": P_contrato,
                        "${FECHA_TRM}": F_TRM,
                        "${PRECIO_TRM}": P_TMR,
                        "${EMAIL_USUARIO}": Email,
                        "${PRECIO_PKWH}":str(round(P_kwh,2)),
                        "${PORC_COMER}":str(int(Porc_come*100))+"%",
                        "${RESPALDO_TOT}":f'{Respaldo:,}',
                        "${INGRESO}":f'{Ingreso:,}',
                        "${COST_COME}":f'{C_comer:,}',
                        "${COST_GMS}":f'{C_GMS:,}',
                        "${INGRESO_NETO}": f'{I_NETO:,}',
                        "${NUM_DIAS}":Num_dias(len(Fechas)),
                        "${RANGO_FECHAS_1}": R_fechas,
                        "${ENTER}": Enter,
                        "${MES_LIQUIDACION}": F_Liq_pag(Opciones2.index(eleccion2)+2,int(today.strftime("%Y"))),
                        "${MES_PAGO}":  F_Liq_pag(Opciones2.index(eleccion2)+3,int(today.strftime("%Y"))),
                        "${INDICADOR}": eleccion3
                    }
                    
                    template_document = docx.Document(template_file_path)
                    
                    for variable_key, variable_value in variables.items():
    
                        for section in template_document.sections:
                            for paragraph in section.header.paragraphs:
                                replace_text_in_paragraph(paragraph, variable_key, variable_value)
    
                        for paragraph in template_document.paragraphs:
                    
                            replace_text_in_paragraph(paragraph, variable_key, variable_value)
                
                        for table in template_document.tables:
                            for col in table.columns:
                                for cell in col.cells:
                                    for paragraph in cell.paragraphs:
                                        replace_text_in_paragraph(paragraph, variable_key, variable_value)
                    
                    rows = template_document.tables[1].rows
                    index_1=Data_frame_fechas.index.values
                    Acum_Req=0
                    Acum_Res=0
                    for idx in index_1:
    
                        
                        rows[int(idx)+1].cells[0].text = Data_frame_fechas.iloc[idx]["Dia"]
                        
                        rows[int(idx)+1].cells[1].text = Data_frame_fechas.iloc[idx]["Fecha"].strftime('%Y-%m-%d')
                        
                        rows[int(idx)+1].cells[2].text = f'{Data_frame_fechas.iloc[idx]["Requerimiento"]:,}'
                        Acum_Req += Data_frame_fechas.iloc[idx]["Requerimiento"]
                        rows[int(idx)+1].cells[3].text = f'{Data_frame_fechas.iloc[idx]["Respaldo"]:,}'
                        Acum_Res += Data_frame_fechas.iloc[idx]["Respaldo"]
                        
                        
                    for idx in np.arange(len(index_1)+1,37):
                        
                        remove_row(template_document.tables[1], rows[len(index_1)+1])    
                        
                    rows[-1].cells[1].text = Num_dias(len(Fechas))
                    rows[-1].cells[2].text = f'{Acum_Req:,}'
                    rows[-1].cells[3].text = f'{Acum_Res:,}'
                    
                    version=1
                
                    template_document.save(Ruta_x+usuario+"_OFR"+".docx")
                    zf.write(Ruta_x+usuario+"_OFR"+".docx")
                    if b:
                        
                        docx2pdf.convert(Ruta_x+"_OFR"+".docx", Ruta_pdf+"/"+usuario+"_OFR"+".pdf")
                        zf.write(Ruta_x+usuario+"_OFR"+".pdf")
                    File_names.extend([usuario+"_OFR"+".docx"])
                    
                    steps_done += 1    
                    my_bar.progress(int(steps_done*100/steps))
                        
                        
                Info.update({"File_names":File_names})
                json_info = json.dumps(Info, indent = 4)
                with open(Ruta_x+'/00_data.json', 'w') as f:
                    json.dump(json_info, f)
                zf.write(Ruta_x+'/00_data.json')    
                zf.close()
                
                
                with open("Resultado.zip", "rb") as fp:
                    with columns_3[1]:
                        btn = st.download_button(
                            label="Descargar resultados",
                            data=fp,
                            file_name="Resultado.zip",
                            mime="application/zip"
                        )
                    
        else:
            st.warning("Necesita subir los tres archivos")   
        
        
            
            
        
            
          
   