import pandas as pd
import numpy as np
import os
from os import remove
from os import path
from os import listdir
from os.path import isfile, join
from numpy import pi, cos, sin, sqrt
from numpy import arccos as acos
import math
import pickle
import matplotlib.pyplot as plt
from datetime import datetime
import streamlit as st
import altair as alt
from matplotlib import cm

import requests

import json


from os import path
import numpy as np

from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
from matplotlib import cm
from matplotlib.ticker import LinearLocator
import time

import glob
from PIL import Image
import smtplib


def User_validation():

    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    
    now=datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M")
    time_past=datetime.strptime(past['Acceso']["Hora"], "%d/%m/%Y %H:%M")
    
    timesince = now - time_past
    Time_min= int(timesince.total_seconds() / 60)
    
    
   
    bool_negate = Time_min<30
  
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
            
        Users=["Francisco Amortegui","Murillo","Correa","Merchan","Juan David","Nicolle","Carlos","Camilo Cortes","David Romero Quete"]
        bool_user = Usuario in Users
        bool_clave = Clave in ["PV-Tool-UN"]
    

        bool_user_email = past['Acceso']["User"] == Usuario
    
        bool_time2 = Time_min<60
        
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
                    
                    
                    
                    subject = 'Acceso aplicacion paneles solares' 
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
def posicion_libre():
    
    

    latitud_long = st.sidebar.text_input("Copie y pegue la ubicación del proyecto" )
    pos_lat=latitud_long.find("Latitude")
    pos_lon=latitud_long.find("Longitude")
    
    if (pos_lat==-1) and (pos_lon==-1):
        latitud_f=51.5099
        longitud_f=-0.1187
    else:
        latitud_f=float(latitud_long[pos_lat+10:pos_lat+17])
        longitud_f=float(latitud_long[pos_lon+11:pos_lon+18])
        if longitud_f > 0:
            if (longitud_f/180)%2<1:
                longitud_f=longitud_f%180
            elif (longitud_f/180)%2>1:
                longitud_f=longitud_f%180-180
        elif longitud_f < 0:
            if (longitud_f/-180)%2<1:
                longitud_f=longitud_f%180-180
            elif (longitud_f/-180)%2>1:
                longitud_f=longitud_f%180
    
    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    
    cond=(latitud_f==float(past["Eleccion"]["Latitud"]))and(float(longitud_f==past["Eleccion"]["Longitud"]))
    if cond:

        altura_f=past["Eleccion"]["Altura"]*1000

    else:

        url = "https://maps.googleapis.com/maps/api/elevation/json?locations="+str(latitud_f)+","+str(longitud_f)+"&key=AIzaSyDhpSAZtcnr-OvNY1f58WpMMXFG_V1_lSs"

        payload={}
        headers = {}
        
        response = requests.request("GET", url, headers=headers, data=payload)
        
        results=json.loads(response.text)["results"]
        
        if len(results) == 0:
            st.sidebar.error("No se encontro la altura para la ubicación seleccionada por favor ingresela manualmente")
            altura_f= st.sidebar.slider("Altura ubicación", min_value=0, max_value=8000, value=0, step=1)
        else:
            altura_f= int(json.loads(response.text)['results'][0]['elevation'])
            

    st.sidebar.write("La altura determinada en las bases de datos de Google es: "+str(altura_f)+" m")
    df = pd.DataFrame(np.zeros_like(np.random.randn(1, 2)) / [5000, 5000] + [latitud_f, longitud_f],columns=['lat', 'lon'])

    st.sidebar.map(df)
    
    return altura_f,latitud_f,longitud_f
def reset():
    dn="Data_set_base/Resultados/"
    paths=[dn+"Paneles.csv",dn+"Control_1.csv",
           dn+"Control_2.csv",dn+"Relieve_libre_eleccion.csv",
           dn+"Sombra_libre_eleccion.csv",dn+"Potencia_total.csv","Validation/Potencia.pickle","Validation/Flujo.pickle"]
    for i in paths:
        if path.exists(i):
            remove(i)
def select_box(sitio=0,opcion=0,Tipo=""):
    
    
    if opcion==1:
        eleccion=st.sidebar.selectbox('Seleccione la ubicación del proyecto', ("Elección libre",
                                                                       "Nunchia - Casanare",
                                                                       "Puerto Carreño",
                                                                       "Almeria - España",
                                                                       "Atacama - Chile"))
    if opcion==2:
        eleccion=st.sidebar.selectbox('Seleccione el tipo de proyecto', ("Venta de energía",
                                                                         "Compra y venta de energía"))
    
    elif (sitio=="Puerto Carreño") or (sitio=="Almeria - España") or (sitio=="Atacama - Chile") or (sitio=="Nunchia - Casanare"):
        
        opt=["Visualización del lugar","Modelo teoríco para la Irradiancia global","Comparación datos experimentales",
                                                                         "Dirección de paneles solares",
                                                                         "Seguimiento del sol en dos ejes",
                                                                         "Seguimiento del sol en eje Norte-Sur",
                                                                         "Comparativa irradiancia",
                                                                         "Potencia instalación"]
        
        if Tipo=="Compra y venta de energía":
            opt.append("Flujo de carga")
            opt.append("Analisis financiero compra y venta de energia")
        elif Tipo=="Venta de energía":
            opt.append("Analisis financiero venta de energia") 
        
        eleccion=st.selectbox('Seleccione el analisis a realizar', (opt))  
    elif sitio=="Elección libre":
        
        opt=["Seleccione el lugar","Modelo teoríco para la Irradiancia global",
                                                                             "Dirección de paneles solares",
                                                                         "Seguimiento del sol en dos ejes",
                                                                         "Seguimiento del sol en eje Norte-Sur",
                                                                         "Comparativa irradiancia",
                                                                         "Potencia instalación"]
        if Tipo=="Compra y venta de energía":
            opt.append("Flujo de carga")
            opt.append("Analisis financiero compra y venta de energia")
        elif Tipo=="Venta de energía":
            opt.append("Analisis financiero venta de energia")    
        eleccion=st.selectbox('Seleccione el analisis a realizar', (opt))
    
    
    return eleccion

def str_potencia(Pot,n=2):
    
    if Pot<1e3:
        string=str(round(Pot,n))+" W"
        
    elif Pot<1e6:
        string=str(round(Pot/1e3,n))+" kW"
        
    elif Pot<1e9:
        string=str(round(Pot/1e6,n))+" MW"
        
    elif Pot<1e12:
        string=str(round(Pot/1e9,n))+" GW"
        
    return string
def str_precio(precio):
    
    string=""
    if precio<0:
        precio=-precio
        string+="-"
    cond=False
    if precio>=1e9:
        string+=str(int(precio/1e9))+"'"
        cond=True
        precio-=int(precio/1e9)*1e9
    if cond or precio>=1e6:
        if cond:
            string+="0"*(3-len(str(int(precio/1e6))))+str(int(precio/1e6))+"'"
        else:
            string+=str(int(precio/1e6))+"'"
        cond=True
        precio-=int(precio/1e6)*1e6
    if cond or precio>=1e3:
        if cond:
            string+="0"*(3-len(str(int(precio/1e3))))+str(int(precio/1e3))+"."
        else:
            string+=str(int(precio/1e3))+"."
        cond=True
        precio-=int(precio/1e3)*1e3
    
    if cond:
        string+="0"*(3-len(str(int(precio))))+str(int(precio))
    else:
        string+=str(int(precio))
    
    return string
def label_lim(maximo,tipo=""):
    

    if maximo>1e6:
        string="MW"
        if tipo=="energia":
            string+="h"
        return 1e6,string
    elif maximo>1e3:
        string="kW"
        if tipo=="energia":
            string+="h"
        return 1e3,string 
    elif maximo>1:
        string="W"
        if tipo=="energia":
            string+="h"
        return 1,string           
def recuperacion(number):
    string=""
    if number>=1:
        if int(number)==1:
            string+=str(int(number))+" año "
        else:
            string+=str(int(number))+" años "
        if (number -int(number)):
            string+="y "
    if (number -int(number))>0:
        if round((number -int(number))*12)==1:
            string+=str(round((number -int(number))*12))+ " mes"
        else:
            string+=str(round((number -int(number))*12))+ " meses"
    return string


def make_gif(frame_folder,name):
    onlyfiles = [f for f in listdir(frame_folder) if isfile(join(frame_folder, f))]
    
    frames=[]
    for i in sorted(onlyfiles):
        dir1=frame_folder+i
        frames.append(Image.open(dir1)) 
    frame_one = frames[0]
    frame_one.save(name+".gif", format="GIF", append_images=frames,
               save_all=True, duration=200, loop=0)

def init_imag():
    
    

    
    if True:
        fig = plt.figure(figsize=(7,7))
        ax = fig.add_subplot(projection='3d')
    
        Lar=0.7
        Cab=0.15
        F_NS_x = np.array([Lar,Lar  ,Lar+np.sqrt(3)*Cab/2,Lar   ,Lar,-Lar,-Lar  ,-Lar-np.sqrt(3)*Cab/2,-Lar,-Lar])
        F_NS_y = np.array([0  ,Cab/2,0                   ,-Cab/2,0  ,0   ,-Cab/2,0                    ,Cab/2 ,0 ])
        F_NS_z = np.zeros_like(F_NS_y)
        
        ax.plot(F_NS_x,F_NS_y,F_NS_z,c="grey")
         
        F_OE_x = np.array([0  ,Cab/2,0                   ,-Cab/2,0  ,0   ,-Cab/2,0                    ,Cab/2 ,0 ])
        F_OE_y = np.array([Lar,Lar  ,Lar+np.sqrt(3)*Cab/2,Lar   ,Lar,-Lar,-Lar  ,-Lar-np.sqrt(3)*Cab/2,-Lar,-Lar])
        F_OE_z = np.zeros_like(F_OE_y)
    
        ax.plot(F_OE_x,F_OE_y,F_OE_z,c="grey")
        
        
        ang1=np.linspace(np.pi+1,0,180)
        ang2=np.linspace(np.pi,2*np.pi+1,180)
        S_x = 0.05*np.sin(ang1)
        S_y = 0.05*np.cos(ang1)+0.95
        S_x =np.append(S_x,  0.05*np.sin(ang2))
        S_y =np.append(S_y,  0.05*np.cos(ang2)+1.05)
        S_z = np.zeros_like(S_y)
        ax.plot(S_x,S_y,S_z,c="grey")
        
        
        ang1=np.linspace(0,2*np.pi,360)
        
        O_x = 0.07*np.sin(ang1)+1
        O_y = 0.10*np.cos(ang1)
    
        O_z = np.zeros_like(O_y)
        ax.plot(O_x,O_y,O_z,c="grey")
        
      
        N_x = np.array([0.05  ,0.05,-0.05,-0.05 ])
        N_y = np.array([-0.9,-1.1  ,-0.9,-1.1])
        
        N_z = np.zeros_like(N_y)
        ax.plot(N_x,N_y,N_z,c="grey")
        
        
        E_x = np.array([-1.05  ,-0.95,-0.95,-1.050,-0.95,-0.95,-1.05])
        E_y = np.array([0.1,0.1  ,0,0,0,-0.1,-0.1])
        
        E_z = np.zeros_like(E_y)
        ax.plot(E_x,E_y,E_z,c="grey")
        
        la=1.5/2
        al=0.07
        z = np.linspace(0, la*1.2, 50)
        theta = np.linspace(0, 2*np.pi, 50)
        theta_grid, z_grid=np.meshgrid(theta, z)
        x_grid = al*np.cos(theta_grid)/3 
        y_grid = al*np.sin(theta_grid)/3 
        
        
        ax.plot_surface(x_grid, y_grid, z_grid,alpha=0.3,color="grey")
        
      
        
        dic_fig={"Figura":fig,"Axis":ax}
        f=open("Imagenes/Fig_base_panel.pickle","wb")
        pickle.dump(dic_fig,f)
        f.close()
        
    else:
        f=open("Imagenes/Fig_base_panel.pickle","rb")
        dic_fig= pickle.load(f)
        f.close()
        
        fig=dic_fig["Figura"]
        ax=dic_fig["Axis"]
        
        
        
    return dic_fig

def panel_g(dic_fig,azimut_p,ele_p,ii,colores_paneles,tipo=""):
    
    an=1/2
    la=1.5/2
    al=0.15
    caja_x = np.array([an,an ,-an,-an,an,an,an ,an ,an ,-an,-an,-an,-an,-an,-an,an])
    caja_y = np.array([la,-la,-la,la ,la,la,-la,-la,-la,-la,-la,-la,la ,la ,la ,la])
    caja_z = np.array([0 ,0  ,0  ,0  ,0 ,al,al ,0  ,al ,al ,0  ,al ,al ,0  ,al ,al])
    supe_x = np.array([an,-an,0 ,-an,an ])
    supe_y = np.array([la,-la,0 ,la ,-la])
    supe_z = np.array([al,al ,al,al ,al ])

            
    
    if tipo=="Direccion":
        
        f=open("Validation/Validation.json","r")
        past=json.loads(f.read())
        f.close()
        
        cond1= past['Graficas paneles']['Estaticos']["Imagen_panel_"+str(ii)]["Azimut"]==azimut_p
        cond2= past['Graficas paneles']['Estaticos']["Imagen_panel_"+str(ii)]["Elevacion"]==ele_p
        
        cond =cond1 and cond2
        if cond:
            f=open("Imagenes/Fig_panel_"+str(ii)+".pickle","rb")
            dic_fig2= pickle.load(f)
            f.close()
            fig=dic_fig2["Figura"]
            ax=dic_fig2["Axis"]
    
        else:
        
            vec_dir_x=np.array([0,0])
            vec_dir_y=np.array([0,0])
            vec_dir_z=np.array([al,al*8])
                    
    
            
            
            cz=cos(np.radians(azimut_p))
            sz=sin(np.radians(azimut_p))
            
            cy=cos(np.radians(ele_p))
            sy=sin(np.radians(ele_p))
            
            M1 = np.array([[1,0,0],[0,cy,sy],[0,-sy,cy]])
                        
            M2 = np.array([[cz,-sz,0],[sz,cz,0],[0,0,1]])
            
            MR = np.matmul(M2, M1)
            
            caja_x2 = MR[0,0]*caja_x +MR[0,1]*caja_y +MR[0,2]*caja_z 
            caja_y2 = MR[1,0]*caja_x +MR[1,1]*caja_y +MR[1,2]*caja_z 
            caja_z2 = MR[2,0]*caja_x +MR[2,1]*caja_y +MR[2,2]*caja_z
            
            supe_x2 = MR[0,0]*supe_x +MR[0,1]*supe_y +MR[0,2]*supe_z 
            supe_y2 = MR[1,0]*supe_x +MR[1,1]*supe_y +MR[1,2]*supe_z 
            supe_z2 = MR[2,0]*supe_x +MR[2,1]*supe_y +MR[2,2]*supe_z
            
            vec_dir_x2= MR[0,0]*vec_dir_x +MR[0,1]*vec_dir_y +MR[0,2]*vec_dir_z 
            vec_dir_y2= MR[1,0]*vec_dir_x +MR[1,1]*vec_dir_y +MR[1,2]*vec_dir_z 
            vec_dir_z2= MR[2,0]*vec_dir_x +MR[2,1]*vec_dir_y +MR[2,2]*vec_dir_z
            
            
            fig=dic_fig["Figura"]
            ax=dic_fig["Axis"]
            ax.set_title("Dirección panel#" + str(ii+1))
        
            ax.plot(caja_x2,caja_y2,caja_z2+la*1.2,c="grey")
            ax.plot(supe_x2,supe_y2,supe_z2+la*1.2,c=colores_paneles[ii])
            ax.plot(vec_dir_x2,vec_dir_y2,vec_dir_z2+la*1.2,c=colores_paneles[ii])
    
            ax.view_init(elev=20, azim=45)
            
            ax.set_xlim3d([-1.5,1.5])
            ax.set_ylim3d([-1.5,1.5])
            ax.set_zlim3d([0.05, 3])
                
            
            past['Graficas paneles']['Estaticos'].update({"Imagen_panel_"+str(ii):{"Azimut":azimut_p,"Elevacion":ele_p}})
            
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
            
            dic_fig1={"Figura":fig,"Axis":ax}
            f=open("Imagenes/Fig_panel_"+str(ii)+".pickle","wb")
            pickle.dump(dic_fig1,f)
            f.close()
        
        st.pyplot(fig)
    elif tipo=="Control_1_eje":
        
        f=open("Validation/Validation.json","r")
        past=json.loads(f.read())
        f.close()
        
        cond1=past['Graficas paneles']["Control_1_eje"]["Elevacion"]==ele_p
        cond2=past['Graficas paneles']["Control_1_eje"]["Limites"]==list(azimut_p)
        cond=cond1 and cond2
        if not cond:
            col1,col2,col3=st.columns([2,3,2])
            with col2:
                
                boolean=st.button("Actualizar la animación del panel")
        else:
            boolean=False
            
        
        if boolean:
            ang_lim=np.radians(azimut_p)
            
            ang_sol=np.linspace(-np.pi/2+0.01,np.pi/2-0.01,ii)
            
      
            
            ang_rot_motor=ang_sol*(ang_sol<ang_lim[1])*(ang_sol>ang_lim[0])+ang_lim[1]*(ang_sol>ang_lim[1])+ang_lim[0]*(ang_sol<ang_lim[0])
         
            ang_rot_motor=np.append(ang_rot_motor,np.flip(ang_rot_motor))
         
            x_or=-sin(ang_rot_motor)
            y_or=np.zeros_like(x_or)
            z_or=cos(ang_rot_motor)
    
            cy=cos(np.radians(ele_p))
            sy=sin(np.radians(ele_p))
            
            MR = np.array([[1,0,0],[0,cy,sy],[0,-sy,cy]])
                        
          
            
            x_f = MR[0,0]*x_or +MR[0,1]*y_or +MR[0,2]*z_or 
            y_f = MR[1,0]*x_or +MR[1,1]*y_or +MR[1,2]*z_or 
            z_f = MR[2,0]*x_or +MR[2,1]*y_or +MR[2,2]*z_or
                    
            
             
            r   = sqrt(np.power(x_f,2)+np.power(y_f,2)+np.power(z_f,2))
            th  = acos(z_f/r)
            phi = np.arctan2(x_f,y_f)
    
                 
            fac=10
            vec_dir_x=np.array([0,0])
            vec_dir_y=np.array([0,0])
            vec_dir_z=np.array([al,al*fac])
            col1,col2=st.columns([2,4])
            with col1:
                st.write("Generando la animación-->")
            with col2:
                my_bar=st.progress(0)

            for jj in range(1,2*ii):
                j=2*ii-jj
                cx=cos(-ang_rot_motor[j])
                sx=sin(-ang_rot_motor[j])
                
                cy=cos(np.radians(ele_p))
                sy=sin(np.radians(ele_p))
                
                M1 = np.array([[1,0,0],[0,cy,sy],[0,-sy,cy]])
                            
                M2 = np.array([[cx,0,sx],[0,1,0],[-sx,0,cx]])
                
                MR = np.matmul(M1, M2)
                
                caja_x2 = MR[0,0]*caja_x +MR[0,1]*caja_y +MR[0,2]*caja_z 
                caja_y2 = MR[1,0]*caja_x +MR[1,1]*caja_y +MR[1,2]*caja_z 
                caja_z2 = MR[2,0]*caja_x +MR[2,1]*caja_y +MR[2,2]*caja_z
                
                supe_x2 = MR[0,0]*supe_x +MR[0,1]*supe_y +MR[0,2]*supe_z 
                supe_y2 = MR[1,0]*supe_x +MR[1,1]*supe_y +MR[1,2]*supe_z 
                supe_z2 = MR[2,0]*supe_x +MR[2,1]*supe_y +MR[2,2]*supe_z
                
                vec_dir_x2= MR[0,0]*vec_dir_x +MR[0,1]*vec_dir_y +MR[0,2]*vec_dir_z 
                vec_dir_y2= MR[1,0]*vec_dir_x +MR[1,1]*vec_dir_y +MR[1,2]*vec_dir_z 
                vec_dir_z2= MR[2,0]*vec_dir_x +MR[2,1]*vec_dir_y +MR[2,2]*vec_dir_z
            
                f=open("Imagenes/Fig_base_panel.pickle","rb")
                dic_fig2= pickle.load(f)
                f.close()
                fig=dic_fig2["Figura"]
                ax=dic_fig2["Axis"]
    
                ax.set_title("Control horizontal")
              
                ax.plot(al*fac*x_f[j:ii],al*fac*y_f[j:ii],al*fac*z_f[j:ii]+la*1.2,c="gold")
                
                
            
                ax.plot(caja_x2,caja_y2,caja_z2+la*1.2,c="grey")
                ax.plot(supe_x2,supe_y2,supe_z2+la*1.2,c="black")
                ax.plot(vec_dir_x2,vec_dir_y2,vec_dir_z2+la*1.2,c="black")
                if ele_p>0 and ele_p<30:
                    elev=60-ele_p
                
                else:
                    elev=30
                
                ax.view_init(elev=elev, azim=90-15*cos(j/ii*np.pi))
                # ax.view_init(elev=20, azim=65)
                        
                
                ax.set_xlim3d([-2,2])
                ax.set_ylim3d([-2,2])
                ax.set_zlim3d([0.05, 3])
                if 2*ii-j <10:
                    string="0"+str(2*ii-j)
                else:
                    string=str(2*ii-j)
                dire="Imagenes/Control_1_eje_GIF/"+string+".jpg"
                plt.savefig(dire) 
                my_bar.progress(int(jj*100/(2*ii)))
                del fig, ax
            
            

            # past["Control_1_eje"]={"Elevacion":ele_p,"Limites":azimut_p}
            past['Graficas paneles'].update({"Control_1_eje":{"Elevacion":ele_p,"Limites":azimut_p}})
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
            make_gif("Imagenes/Control_1_eje_GIF/",name="Imagenes/1_eje")
            
        st.image("Imagenes/1_eje.gif")
        
    elif tipo=="Control_2_ejes":
        
        f=open("Validation/Validation.json","r")
        past=json.loads(f.read())
        f.close()
        cond1=past['Graficas paneles']["Control_2_eje"]["Limites_inclinacion"]==list(ele_p)
        cond2=past['Graficas paneles']["Control_2_eje"]["Limites_azimut"]==list(azimut_p)
        cond=cond1 and cond2
        if not cond:
            col1,col2,col3=st.columns([2,3,2])
            with col2:
                
                boolean=st.button("Actualizar la animación del panel")
        else:
            boolean=False
            
        
        if boolean:
            
            azim_lim=np.radians(azimut_p)
            
            
            incl_lim=np.radians(ele_p)

            
            ang_azi=np.linspace(azim_lim[0],azim_lim[1],ii)
            
            ang_g=np.linspace(-np.pi,np.pi,ii)
            
            if (azim_lim[1]-azim_lim[0])<np.pi/2:
                f=1/2
            elif (azim_lim[1]-azim_lim[0])<np.pi:
                f=1.5
            else:
                f=2
            ang_incl=(incl_lim[1]-incl_lim[0])*cos(f*ang_g)/2+(incl_lim[1]+incl_lim[0])/2
            
            ang_azi=np.append(ang_azi,np.flip(ang_azi))
            ang_incl=np.append(ang_incl,np.flip(ang_incl))
            
            r=1
            
            x_f=-r*sin(ang_azi)*sin(ang_incl)
            y_f=r*cos(ang_azi)*sin(ang_incl) 
            z_f=r*cos(ang_incl)
            

            fac=10
            vec_dir_x=np.array([0,0])
            vec_dir_y=np.array([0,0])
            vec_dir_z=np.array([al,al*fac])
            col1,col2=st.columns([2,4])
            
            with col1:
                st.write("Generando la animación-->")
            with col2:
                my_bar=st.progress(0)

            for jj in range(0,2*ii):
                
                
                cz=cos(ang_azi[jj])
                sz=sin(ang_azi[jj])
                
                cy=cos(ang_incl[jj])
                sy=sin(ang_incl[jj])
                
                M1 = np.array([[1,0,0],[0,cy,sy],[0,-sy,cy]])
                            
                M2 = np.array([[cz,-sz,0],[sz,cz,0],[0,0,1]])
            
                MR = np.matmul(M2, M1)
                
                caja_x2 = MR[0,0]*caja_x +MR[0,1]*caja_y +MR[0,2]*caja_z 
                caja_y2 = MR[1,0]*caja_x +MR[1,1]*caja_y +MR[1,2]*caja_z 
                caja_z2 = MR[2,0]*caja_x +MR[2,1]*caja_y +MR[2,2]*caja_z
                
                supe_x2 = MR[0,0]*supe_x +MR[0,1]*supe_y +MR[0,2]*supe_z 
                supe_y2 = MR[1,0]*supe_x +MR[1,1]*supe_y +MR[1,2]*supe_z 
                supe_z2 = MR[2,0]*supe_x +MR[2,1]*supe_y +MR[2,2]*supe_z
                
                vec_dir_x2= MR[0,0]*vec_dir_x +MR[0,1]*vec_dir_y +MR[0,2]*vec_dir_z 
                vec_dir_y2= MR[1,0]*vec_dir_x +MR[1,1]*vec_dir_y +MR[1,2]*vec_dir_z 
                vec_dir_z2= MR[2,0]*vec_dir_x +MR[2,1]*vec_dir_y +MR[2,2]*vec_dir_z
            
                f=open("Imagenes/Fig_base_panel.pickle","rb")
                dic_fig2= pickle.load(f)
                f.close()
                fig=dic_fig2["Figura"]
                ax=dic_fig2["Axis"]
    
                ax.set_title("Control de los dos ejes")
              
                ax.plot(al*fac*x_f,al*fac*y_f,al*fac*z_f+la*1.2,c="gold")
                
                
            
                ax.plot(caja_x2,caja_y2,caja_z2+la*1.2,c="grey")
                ax.plot(supe_x2,supe_y2,supe_z2+la*1.2,c="grey")
                ax.plot(vec_dir_x2,vec_dir_y2,vec_dir_z2+la*1.2,c="black")
                
                
                ax.view_init(elev=30, azim=(azimut_p[1]+azimut_p[0])/2+90-25*cos(jj/ii*np.pi))
                # ax.view_init(elev=20, azim=65)
                        
                
                ax.set_xlim3d([-2,2])
                ax.set_ylim3d([-2,2])
                ax.set_zlim3d([0.05, 3])
                if jj <10:
                    string="0"+str(jj)
                else:
                    string=str(jj)
                
                dire="Imagenes/Control_2_ejes_GIF/"+string+".jpg"
                plt.savefig(dire)
                my_bar.progress(int((jj+1)*100/(2*ii)))
                del fig, ax
            
            
            past['Graficas paneles'].update({"Control_2_eje":{"Limites_inclinacion":ele_p,"Limites_azimut":azimut_p}})
            str_json_0=json.dumps(past, indent=4)
            J_f=open("Validation/Validation.json","w")
            J_f.write(str_json_0)
            J_f.close()
            
            make_gif("Imagenes/Control_2_ejes_GIF/",name="Imagenes/2_ejes")
            
        st.image("Imagenes/2_ejes.gif")

def actual_seccion(opcion):
    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    past["Comprobacion"].update({"Seccion":opcion})
    str_json_0=json.dumps(past, indent=4)
    J_f=open("Validation/Validation.json","w")
    J_f.write(str_json_0)
    J_f.close()
def actual_site(opcion):
    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    past["Comprobacion"].update({"Sitio":opcion})
    str_json_0=json.dumps(past, indent=4)
    J_f=open("Validation/Validation.json","w")
    J_f.write(str_json_0)
    J_f.close()
def actual_lat_long(lat,lon):
    f=open("Validation/Validation.json","r")
    past=json.loads(f.read())
    f.close()
    past["Comprobacion"].update({"Latitud":lat})
    past["Comprobacion"].update({"Longitud":lon})
    str_json_0=json.dumps(past, indent=4)
    J_f=open("Validation/Validation.json","w")
    J_f.write(str_json_0)
    J_f.close()