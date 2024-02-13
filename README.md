from __future__ import print_function
from jinja2 import Environment, FileSystemLoader
import pandas as pd
import numpy as np
import datetime
import math
import os




#Carga archivos globales
metas = pd.read_excel('BaseInforme.xlsx',"Base.Metas")
rutas = pd.read_excel('BaseInforme.xlsx', "Rutas")
objetivos = pd.read_excel('BaseInforme.xlsx',"Base.Objetivos")

##Reportes que se desa generar
generar_rp = False
generar_iv = True
cargar_informacion = False

#Fecha Reporte
fecha = datetime.datetime.now()

#Semáforos
icono_abajo = "icon ion-md-arrow-dropdown-circle"
icono_neutral = "icon ion-md-remove-circle"
icono_arriba = "icon ion-md-arrow-dropup-circle"

#Dotacion activa
semaforo_rojo_dot = 0.6
semaforo_amarillo_dot = 0.8

#objetivo UF
semaforo_rojo_UF = 0.75
semaforo_amarillo_UF = 1.0

#Vs mes anterior
semaforo_rojo_ant = -0.05
semaforo_amarillo_ant = 0

#asegurados nuevos
semaforo_rojo_Aseg_Nuevos = 0.6
semaforo_amarillo_Aseg_Nuevos = 0.7

#Venta no presencial
semaforo_rojo_online = 0.6
semaforo_amarillo_online = 0.4

#Aportes Adicionales
semaforo_rojo_aportes = 0.8
semaforo_amarillo_aportes = 0.9

#Rentas Privadas
#semaforo_rojo_rp = 0.8
#semaforo_amarillo_rp = 0.9

#Tasa Renovación RP Mes
#semaforo_rojo_renov_rp = 0.5
#semaforo_amarillo_renov_rp = 0.8


if generar_iv:
    print("Informe de Ventas")
    #Carga archivos particulares
    base = pd.read_excel('BaseInforme.xlsx',"Base.Principal")       #Datos para el informe diario
    base_ventas = pd.read_excel('BaseInforme.xlsx',"Base.Ventas")   #Base de ventas    
    base_ventas = base_ventas.loc[(base_ventas['Nro. Poliza'] == 0) & (base_ventas['Estado Solicitud'] != 'Anuladas') & (base_ventas['Producto'] != 'Cláusula de Ahorro')  ]
    base_emision = pd.read_excel('BaseInforme.xlsx',"Base.Emision") #Base Emisión
    
    
#     print(base_ventas)    
#     print (base_emision)           
#     base = pd.read_excel('BasePrueba2.xlsx',"Base.Principal Cierre")        ##Datos para el informe de cierre de emisión
    
    
    #Objetivos nacionales    
    
    meta_UF = 1718.3     #Reemplazar con alguna fórmula que traiga los presupuestos y metas por indicador
    presupuesto_CUI = 0.0
    #presupuesto_rp = 0.0    #Meta o Presupuesto?
    presupuesto_aportes = 10901.9
    
    meses_grafico = ['Diciembre', 'Enero','Febrero']
    col_activo3 = 'Activo'
    col_con_venta3 = 'Con Venta'
    

    
    # base.to_excel("Resultado3.xlsx",sheet_name='Nacional')
    
    columnas_totales = ['Vigente','Activo', 'Con Venta','Prima Emitida', 'Prima en Proceso','Prima LT', 'Prima CUI', 'Solicitudes Totales',  'Aportes', 'Prima LT Ant', 'Prima CUI Ant', 'Prima Total Ant', 'Prima Mensual', 'Persistencia', 'Activo -1', 'Activo -2', 'Con Venta -1', 'Con Venta -2', 'Activo 1', 'Con Venta 1', 'Asegurados Nuevos']
    
    
    
    
    columnas_prueba = ['Vigente','Activo', 'Con Venta', 'Prima Mensual', 'Solicitudes Totales', 'Asegurados Nuevos', 'Aportes','Tasa Vta.','% Var Mes Ant','Tasa Aseg. Nuevos']
#     columnas = ['Vigente','Activo', 'Con Venta', 'Con Venta No Presencial', 'Prima LT', 'Prima CUI', 'Prima No Presencial', 'Solicitudes Totales', 'Solicitudes No Presencial', 'RP', 'Aportes']
    columnas_int = ['Vigente','Activo', 'Con Venta', 'Prima Mensual', 'Solicitudes Totales', 'Asegurados Nuevos', 'Aportes', 'Persistencia','% Var Mes Ant','Tasa Aseg. Nuevos']
    mix_producto = ['FAMILIAR', 'SO', 'MCG y VE', 'MP10', 'AP','CUI']
    
    #Formatos
    columnas_enteras = ['Z','S','Vigente', 'Activo', 'Con Venta', 'Solicitudes Totales','Asegurados Nuevos']
    columnas_decimal = ['Prima Mensual','Aportes']
    columnas_porcentaje = ['Tasa Vta.', 'Tasa Aseg. Nuevos']
    
    # base['Prima LT'] = round(base['Prima LT'], 2)
    
    
    
    
    base[columnas_enteras] = base[columnas_enteras].astype(int)
    
    
    #VISTA NACIONAL
    
    base_mix = pd.pivot_table(base,
                              index=['Zona','Sucursal'],
                              values=mix_producto,
                              aggfunc=np.sum)
    base_mix = base_mix.reindex(mix_producto, axis=1)
    base[mix_producto] = base[mix_producto].astype(int)
    
    
    
    # print(tabla_mix_nacional)
    
    tabla_nacional = pd.pivot_table(base,
                                    index=['Z', 'Zona', 'S', 'Sucursal'],
                                    values=columnas_totales)
    
    zonas = tabla_nacional.index.get_level_values(1).unique()
    sucursales = tabla_nacional.index.get_level_values(3).unique()
    
    
    
    
    #Total Nacional
    orden_zonas = ['Norte', 'Sur']
    # orden_zonas_z = ['1','2','3']
    orden_sucursales = ['Arica', 'Iquique', 'Antofagasta', 'La Serena', 'Quilpué', 'Santiago', 'Valparaíso', 'Viña del Mar', 'Rancagua', 'Talca', 'Concepción', 'Talcahuano', 'Temuco', 'Valdivia', 'Puerto Montt', 'Punta Arenas']
    orden_sucursales_n = ['Arica', 'Iquique', 'Antofagasta', 'La Serena', 'Quilpué', 'Santiago']
    salto_linea = 3
    dict_zonas = {'Norte': ['Arica', 'Iquique', 'Antofagasta', 'La Serena', 'Quilpué', 'Santiago', 'Viña del Mar'], 'Sur': ['Rancagua', 'Talca', 'Concepción', 'Temuco', 'Valdivia', 'Puerto Montt', 'Punta Arenas'], 'Suc. Virtual': ['Suc. Virtual'], 'Valparaíso': ['Valparaíso'], 'Talcahuano': ['Talcahuano'], 'Venta Asistida': ['Venta Asistida'], 'Agente Libre': ['Agente Libre'] }
    
    
    def formato(data_frame, columnas, formato):    
    #     data_frame = data_frame.fillna(0)
        if formato == 'Porcentaje':
            formato = "{:.2%}"
            for columna in columnas:
                data_frame[columna] = pd.Series([formato.format(val).replace(',','~').replace('.',',').replace('~','.') for val in data_frame[columna]], index = data_frame.index)
            
        if formato == '1 Decimal':
            for columna in columnas:
                data_frame[columna] = round(data_frame[columna],2)
    #         formato = "{:,.2f}"
        
    #     for columna in columnas:
    #         data_frame[columna] = pd.Series([formato.format(val).replace(',','~').replace('.',',').replace('~','.') for val in data_frame[columna]], index = data_frame.index)
    
    
    
    
    ##Generar reporte en Excel
    with pd.ExcelWriter('Resultado2.xlsx',engine='openpyxl') as ew:    
        nacional_total = pd.pivot_table(base,
                                        index=['Total'],
                                        values = columnas_totales,
                                        aggfunc=np.sum)
    #     print(nacional_total)
        
    
        
        ##Calcular indicadores
        nacional_total['Tasa Vta.'] = nacional_total['Con Venta']/nacional_total['Activo']
        #nacional_total['Tasa Vta. No Presencial'] = nacional_total['Con Venta No Presencial']/nacional_total['Con Venta']
        nacional_total['% Var Mes Ant'] = nacional_total['Prima Mensual']/nacional_total['Prima Total Ant'] - 1
        nacional_total['Tasa Aseg. Nuevos'] = nacional_total['Asegurados Nuevos']/nacional_total['Solicitudes Totales']
        
        ##Ordenar columnas de la tabla
        nacional_total = nacional_total.reindex(columnas_prueba, axis=1)
        
        ##Dar formato
        formato(nacional_total, columnas_porcentaje, 'Porcentaje')
        formato(nacional_total, columnas_decimal, '1 Decimal')
        
        nacional_total.to_excel(ew,sheet_name='Nacional',startrow = 0, startcol=0)
        fila1 = len(nacional_total) + 2
        
        #Hacer Tabla dinámica por canal
        nacional_canal = pd.pivot_table(base,index=['Canal'],
                                       values = columnas_totales,
                                       aggfunc=np.sum)
        nacional_canal['Tasa Vta.'] = nacional_canal['Con Venta']/nacional_canal['Activo']
        #nacional_canal['Tasa Vta. No Presencial'] = nacional_canal['Con Venta No Presencial']/nacional_canal['Con Venta']
        nacional_canal['% Var Mes Ant'] = round(nacional_canal['Prima Mensual']/nacional_canal['Prima Total Ant'] - 1,2)
        nacional_canal['Tasa Aseg. Nuevos'] = nacional_canal['Asegurados Nuevos']/nacional_canal['Solicitudes Totales']
        nacional_canal = nacional_canal.reindex(columnas_prueba, axis=1)
        #Transformar indices en columnas
        nacional_canal.reset_index(level=0, inplace=True)
        
        #Dar formato
        formato(nacional_canal, columnas_porcentaje, 'Porcentaje')
        formato(nacional_canal, columnas_decimal, '1 Decimal')    
    
        nacional_canal.to_excel(ew,sheet_name='Nacional',startrow = fila1, startcol=0)
        fila2 = fila1+len(nacional_canal) + salto_linea
        
        nacional_zona = pd.pivot_table(base,index=['Z','Zona'],
                                       values = columnas_totales,
                                       aggfunc=np.sum)
    
    #     nacional_zona = nacional_zona.reindex(orden_zonas, axis=0)    
        nacional_zona['Tasa Vta.'] = nacional_zona['Con Venta']/nacional_zona['Activo']
        #nacional_zona['Tasa Vta. No Presencial'] = nacional_zona['Con Venta No Presencial']/nacional_zona['Con Venta']
        nacional_zona['% Var Mes Ant'] = round(nacional_zona['Prima Mensual']/nacional_zona['Prima Total Ant'] - 1,2)
        nacional_zona['Tasa Aseg. Nuevos'] = nacional_zona['Asegurados Nuevos']/nacional_zona['Solicitudes Totales']
        nacional_zona = nacional_zona.reindex(columnas_prueba, axis=1)
        
        nacional_zona.reset_index(level=1, inplace=True)
        nacional_zona.reset_index(level=0, inplace=True)
        
        formato(nacional_zona, columnas_porcentaje, 'Porcentaje')
        formato(nacional_zona, columnas_decimal, '1 Decimal')    
    #     nacional_zona['Tasa Vta.'] = pd.Series(["{:.2%}".format(val) for val in nacional_zona['Tasa Vta.']], index = nacional_zona.index)
    #     nacional_zona['Tasa Vta. No Presencial'] = pd.Series(["{:.2%}".format(val) for val in nacional_zona['Tasa Vta. No Presencial']], index = nacional_zona.index)
        nacional_zona.to_excel(ew,sheet_name='Nacional',startrow = fila2, startcol=0)
        fila3 = fila2 + len(nacional_zona) + salto_linea
        
        nacional_sucursal = pd.pivot_table(base,index=['S','Sucursal'],
                                       values = columnas_totales,
                                       aggfunc=np.sum)
    #     nacional_sucursal.to_excel("Resultado.xlsx")
        
        
    #     nacional_sucursal = nacional_sucursal.reindex(orden_sucursales, axis = 0)
        nacional_sucursal['Tasa Vta.'] = nacional_sucursal['Con Venta']/nacional_sucursal['Activo']
        #nacional_sucursal['Tasa Vta. No Presencial'] = nacional_sucursal['Con Venta No Presencial']/nacional_sucursal['Con Venta']
        nacional_sucursal['% Var Mes Ant'] = round(nacional_sucursal['Prima Mensual']/nacional_sucursal['Prima Total Ant'] - 1,2)
        nacional_sucursal['Tasa Aseg. Nuevos'] = nacional_sucursal['Asegurados Nuevos']/nacional_sucursal['Solicitudes Totales']
        nacional_sucursal = nacional_sucursal.reindex(columnas_prueba, axis=1)
        
        nacional_sucursal.reset_index(level=1, inplace=True)
        nacional_sucursal.reset_index(level=0, inplace=True)
        
        formato(nacional_sucursal, columnas_porcentaje, 'Porcentaje')
        formato(nacional_sucursal, columnas_decimal, '1 Decimal')
        nacional_sucursal.to_excel(ew,sheet_name='Nacional',startrow = fila3, startcol=0)
    
    
    
    #REPORTE HTML NACIONAL#
    

    
    
    #Indicadores Generales
    
    #1 Dotacion Nacional
    
    dot_vta = base['Con Venta'].sum()
    dot_act = base['Activo'].sum()
    porcentaje = dot_vta/dot_act
    
    if porcentaje <  semaforo_rojo_dot:
        logro_dot = "text-danger"
        icon_dot = icono_abajo
    elif porcentaje < semaforo_amarillo_dot:
        logro_dot = "text-warning"
        icon_dot = icono_neutral
    else:
        logro_dot = "text-success"
        icon_dot = icono_arriba
        
    porc_act = str(round(porcentaje*100,1)) + "%"
    
    #2 Cumplimiento objetivos
        
    totalUF = round(base['Prima Mensual'].sum(), 2)
    porcentaje = totalUF/meta_UF
    
    if porcentaje <  semaforo_rojo_UF:
        logro_UF = "text-danger"
        icon_UF = icono_abajo
    elif porcentaje < semaforo_amarillo_UF:
        logro_UF = "text-warning"
        icon_UF = icono_neutral
    else:
        logro_UF = "text-success"
        icon_UF = icono_arriba
    
    porc_cump_UF = str(round(porcentaje*100,1)) + "%"
    
    #3 Cumplimiento asegurados nuevos
    
    asegurados_nuevos = base['Asegurados Nuevos'].sum()
    asegurados_totales = base['Solicitudes Totales'].sum()
    porcentaje = asegurados_nuevos/asegurados_totales
        
    if porcentaje <  semaforo_rojo_Aseg_Nuevos:
        logro_Aseg_Nuevos = "text-danger"
        icon_Aseg_Nuevos = icono_abajo
    elif porcentaje < semaforo_amarillo_Aseg_Nuevos:
        logro_Aseg_Nuevos = "text-warning"
        icon_Aseg_Nuevos = icono_neutral
    else:
        logro_Aseg_Nuevos = "text-success"
        icon_Aseg_Nuevos = icono_arriba
        
    porc_cump_Aseg_Nuevos = str(round(porcentaje*100,1)) + "%"
    
    #4 Venta No Presencial (comparada con venta total)
    
    #onlineUF = round(base['Prima No Presencial'].sum(), 2)
    totalUF =  round(base['Prima Mensual'].sum(), 2)
    #porcentaje = onlineUF/totalUF
    
    if porcentaje <  semaforo_rojo_online:
        logro_online = "text-danger"
        icon_online = icono_abajo
    elif porcentaje < semaforo_amarillo_online:
        logro_online = "text-warning"
        icon_online = icono_neutral
    else:
        logro_online = "text-success"
        icon_online = icono_arriba
        
    porc_cump_online = str(round(porcentaje*100,1)) + "%"
    
    #5 Aportes Adicionales
    
    aportes = round(base['Aportes'].sum(), 2)
    porcentaje = aportes/presupuesto_aportes
    
    if porcentaje <  semaforo_rojo_aportes:
        logro_aportes = "text-danger"
        icon_aportes = icono_abajo
    elif porcentaje < semaforo_amarillo_aportes:
        logro_aportes = "text-warning"
        icon_aportes = icono_neutral
    else:
        logro_aportes = "text-success"
        icon_aportes = icono_arriba
        
    porc_cump_aportes = str(round(porcentaje*100,1)) + "%"
    
    #6 Rentas Privadas
   # rp = round(base['RP'].sum(), 2)
    #porcentaje = rp/presupuesto_rp
    
    #if porcentaje <  semaforo_rojo_rp:
        #logro_rp = "text-danger"
        #icon_rp = icono_abajo
   # elif porcentaje < semaforo_amarillo_rp:
        #logro_rp = "text-warning"
        #icon_rp = icono_neutral
    #else:
        #logro_rp = "text-success"
        #icon_rp = icono_arriba
        
    #porc_cump_rp = str(round(porcentaje*100,1)) + "%"
    
    
    
    
    
    #Gráfico mix de producto
    # ['FAMILIAR', 'MCG y VE', 'AP', 'SO', 'MP10', 'CUI']
    
    familiares = base_mix['FAMILIAR'].sum()
    mcg_ve = base_mix['MCG y VE'].sum()
    ap = base_mix['AP'].sum()
    so = base_mix['SO'].sum()
    mp10 = base_mix['MP10'].sum()
    cui = base_mix['CUI'].sum()
    
    #Gráfico Dotación
    tabla_a = base[base['Canal'] == "Asesor"]
    a_activo1 = tabla_a['Activo -2'].sum()
    a_activo2 = tabla_a['Activo -1'].sum()
    a_activo3 = tabla_a[col_activo3].sum()
    a_con_venta1 = tabla_a['Con Venta -2'].sum()
    a_con_venta2 = tabla_a['Con Venta -1'].sum()
    a_con_venta3 = tabla_a[col_con_venta3].sum()
    a_activo = [a_activo1, a_activo2, a_activo3]
    a_con_venta = [a_con_venta1, a_con_venta2, a_con_venta3]
    
    
    
    #tabla_af = base[base['Canal'] == "Asesor Financiero"]
    #af_activo1 = tabla_af['Activo -2'].sum()
    #af_activo2 = tabla_af['Activo -1'].sum()
    #af_activo3 = tabla_af[col_activo3].sum()
    #af_con_venta1 = tabla_af['Con Venta -2'].sum()
    #af_con_venta2 = tabla_af['Con Venta -1'].sum()
    #af_con_venta3 = tabla_af[col_con_venta3].sum()
    #af_activo = [af_activo1, af_activo2, af_activo3]
    #af_con_venta = [af_con_venta1, af_con_venta2, af_con_venta3]
    
    
    
    
    
    
    
    env = Environment(loader = FileSystemLoader(''))
    template = env.get_template("template2.html")
    
    #index = False para que no se muestren el 0 al 3
    html = template.render(vista = "Nacional",
                           fecha = fecha.strftime("%d-%m-%y"),
                           dotacion_vta = dot_vta,
                           dotacion_act = dot_act,
                           porcentaje_act = porc_act,
                           indicador_dot = logro_dot,
                           icon_dotacion = icon_dot,
                           total_UF = totalUF,
                           meta_UF = meta_UF,
                           porcentaje_UF = porc_cump_UF,
                           indicador_UF = logro_UF,
                           icon_UF = icon_UF,
                           asegurados_nuevos = asegurados_nuevos,
                           asegurados_totales = asegurados_totales,
                           porcentaje_Aseg_Nuevos = porc_cump_Aseg_Nuevos,
                           indicador_Aseg_Nuevos = logro_Aseg_Nuevos, 
                           icon_Aseg_Nuevos = icon_Aseg_Nuevos,                 
                           #online_UF = onlineUF,
                           porcentaje_online = porc_cump_online,
                           indicador_online = logro_online,
                           icon_online = icon_online,
                           aportes_UF= aportes,
                           presupuesto_aportes = presupuesto_aportes,
                           porcentaje_aportes = porc_cump_aportes,
                           indicador_aportes = logro_aportes,
                           icon_aportes = icon_aportes,
                          # rp_UF = rp,
                          # presupuesto_rp = presupuesto_rp,
                          # porcentaje_rp= porc_cump_rp,
                          # indicador_rp= logro_rp,
                          # icon_rp = icon_rp,
                           grafico_canal=[familiares, mcg_ve, ap, so, mp10, cui],
                           meses_grafico = meses_grafico,
                           a_activo = a_activo,
                           #af_activo = af_activo,
                           a_con_venta = a_con_venta,
                           #af_con_venta = af_con_venta,                       
                           tabla_canal = nacional_canal.to_html(classes="table table-striped table-bordered compact hover formato_dt", index = False),
                           tabla_zona = nacional_zona.to_html(classes="table table-striped table-bordered compact hover formato_dt_ord", index = False),
                           tabla_sucursal = nacional_sucursal.to_html(classes="table table-striped table-bordered compact hover formato_dt_ord", index = False))
    
    
    
    # print(html)
    with open('sitio/index.html', 'w', encoding="utf-8") as f:
        f.write(html)
        
   

    
    
    #REPORTE HTML ZONAS#
    
    for zona in zonas:    
        tabla_zona_sucursal = pd.pivot_table(base,index=['Zona', 'Sucursal'],values=columnas_totales,aggfunc=np.sum)
        tabla_zona_canal =  pd.pivot_table(base,index=['Zona', 'Canal'],values=columnas_totales,aggfunc=np.sum)
    
        tabla_zona_sucursal = tabla_zona_sucursal.xs(zona, level = 0)   #Filtro zona (solo dejo los datos de la zona y elimino el índice)
        tabla_zona_sucursal = tabla_zona_sucursal.reindex(dict_zonas[zona], axis = 0)
    
        tabla_zona_sucursal['Tasa Vta.'] = tabla_zona_sucursal['Con Venta']/tabla_zona_sucursal['Activo']
        #tabla_zona_sucursal['Tasa Vta. No Presencial'] = tabla_zona_sucursal['Con Venta No Presencial']/tabla_zona_sucursal['Con Venta']
        tabla_zona_sucursal['% Var Mes Ant'] = round(tabla_zona_sucursal['Prima Mensual']/tabla_zona_sucursal['Prima Total Ant'] - 1,2)
        tabla_zona_sucursal['Tasa Aseg. Nuevos'] = tabla_zona_sucursal['Asegurados Nuevos']/tabla_zona_sucursal['Solicitudes Totales']
        
        tabla_zona_canal = tabla_zona_canal.xs(zona, level = 0)
        tabla_zona_canal['Tasa Vta.'] = tabla_zona_canal['Con Venta']/tabla_zona_canal['Activo']
        #tabla_zona_canal['Tasa Vta. No Presencial'] = tabla_zona_canal['Con Venta No Presencial']/tabla_zona_canal['Con Venta']
        tabla_zona_canal['% Var Mes Ant'] = round(tabla_zona_canal['Prima Mensual']/tabla_zona_canal['Prima Total Ant'] - 1,2)
        tabla_zona_canal['Tasa Aseg. Nuevos'] = tabla_zona_canal['Asegurados Nuevos']/tabla_zona_canal['Solicitudes Totales']
    
        
        #1 Dotacion Zona
        dot_vta = tabla_zona_sucursal['Con Venta'].sum()
        dot_act = tabla_zona_sucursal['Activo'].sum()
        porcentaje = dot_vta/dot_act
        
        if porcentaje <  semaforo_rojo_dot:
            logro_dot = "text-danger"
            icon_dot = icono_abajo
        elif porcentaje < semaforo_amarillo_dot:
            logro_dot = "text-warning"
            icon_dot = icono_neutral
        else:
            logro_dot = "text-success"
            icon_dot = icono_arriba
            
        porc_act = str(round(porcentaje*100,1)) + "%"
        
        #2 Cumplimiento objetivos
        metas_uf = objetivos[objetivos['Zona'] == zona]  ##filtrar metas por sucursal
        meta_UF = round(metas_uf['Meta Mes Ant'].sum(), 2)
        
        totalUF = round(tabla_zona_sucursal['Prima Mensual'].sum(), 2)
        
        porcentaje = totalUF/meta_UF        
        
        if porcentaje <  semaforo_rojo_UF:
            logro_UF = "text-danger"
            icon_UF = icono_abajo
        elif porcentaje < semaforo_amarillo_UF:
            logro_UF = "text-warning"
            icon_UF = icono_neutral
        else:
            logro_UF = "text-success"
            icon_UF = icono_arriba
        
        porc_cump_UF = str(round(porcentaje*100,1)) + "%"
        
        
        #3 Cumplimiento asegurados nuevos  
        
        asegurados_nuevos = tabla_zona_sucursal['Asegurados Nuevos'].sum()    
        asegurados_totales = tabla_zona_sucursal['Solicitudes Totales'].sum()    
        porcentaje = asegurados_nuevos/asegurados_totales    
        
        if porcentaje <  semaforo_rojo_Aseg_Nuevos:
            logro_Aseg_Nuevos = "text-danger"
            icon_Aseg_Nuevos = icono_abajo
        elif porcentaje < semaforo_amarillo_Aseg_Nuevos:
            logro_Aseg_Nuevos = "text-warning"
            icon_Aseg_Nuevos = icono_neutral
        else:
            logro_Aseg_Nuevos = "text-success"
            icon_Aseg_Nuevos = icono_arriba        
        porc_cump_Aseg_Nuevos = str(round(porcentaje*100,1)) + "%"   
        
        
        
    
        
        #Venta Online (comparada con venta total)
    
        #onlineUF = round(tabla_zona_sucursal['Prima No Presencial'].sum(), 2)
        totalUF = round(tabla_zona_sucursal['Prima Mensual'].sum(), 2)
        #porcentaje = onlineUF/totalUF
        
        if porcentaje <  semaforo_rojo_online:
            logro_online = "text-danger"
            icon_online = icono_abajo
        elif porcentaje < semaforo_amarillo_online:
            logro_online = "text-warning"
            icon_online = icono_neutral
        else:
            logro_online = "text-success"
            icon_online = icono_arriba
            
        porc_cump_online = str(round(porcentaje*100,1)) + "%"
        
        
        #Aportes Adicionales
        aportes = round(tabla_zona_sucursal['Aportes'].sum(), 2)    
        
        
        #Rentas Privadas
        
        #rp = round(tabla_zona_sucursal['RP'].sum(), 2)
        
        #metas_rp = metas[metas['Zona'] == zona]  ##filtrar metas por sucursal
        #presupuesto_rp = round(metas_rp['Meta RP'].sum(), 2)
        
        #porcentaje = rp/presupuesto_rp
        
        #if porcentaje <  semaforo_rojo_rp:
            #logro_rp = "text-danger"
            #icon_rp = icono_abajo
        #elif porcentaje < semaforo_amarillo_rp:
            #logro_rp = "text-warning"
            #icon_rp = icono_neutral
        #else:
            #logro_rp = "text-success"
            #icon_rp = icono_arriba
            
        #porc_cump_rp = str(round(porcentaje*100,1)) + "%"  
        
        #Gráficos
        
        #Mix de producto de la zona
        tabla_zona_mix = base_mix.xs(zona,level = 0)
        familiares = tabla_zona_mix['FAMILIAR'].sum()
        mcg_ve = tabla_zona_mix['MCG y VE'].sum()
        ap = tabla_zona_mix['AP'].sum()
        so = tabla_zona_mix['SO'].sum()
        mp10 = tabla_zona_mix['MP10'].sum()
        cui = tabla_zona_mix['CUI'].sum()
        
        #Gráfico Dotación
    
        tabla_a = tabla_zona_canal[np.in1d(tabla_zona_canal.index.get_level_values('Canal'), ['Asesor'])]    
        a_activo1 = tabla_a['Activo -2'].sum()
        a_activo2 = tabla_a['Activo -1'].sum()
        a_activo3 = tabla_a[col_activo3].sum()
        a_con_venta1 = tabla_a['Con Venta -2'].sum()
        a_con_venta2 = tabla_a['Con Venta -1'].sum()
        a_con_venta3 = tabla_a[col_con_venta3].sum()
        a_activo = [a_activo1, a_activo2, a_activo3]
        a_con_venta = [a_con_venta1, a_con_venta2, a_con_venta3]    
        
        #tabla_af = tabla_zona_canal[np.in1d(tabla_zona_canal.index.get_level_values('Canal'), ['Asesor Financiero'])]  
        #af_activo1 = tabla_af['Activo -2'].sum()
        #af_activo2 = tabla_af['Activo -1'].sum()
        #af_activo3 = tabla_af[col_activo3].sum()
        #af_con_venta1 = tabla_af['Con Venta -2'].sum()
        #af_con_venta2 = tabla_af['Con Venta -1'].sum()
        #af_con_venta3 = tabla_af[col_con_venta3].sum()
        #af_activo = [af_activo1, af_activo2, af_activo3]
        #af_con_venta = [af_con_venta1, af_con_venta2, af_con_venta3]
        
        #Ajuste de tablas y formatos
        
        tabla_zona_sucursal = tabla_zona_sucursal.reindex(columnas_prueba, axis = 1)
        tabla_zona_sucursal.reset_index(level=0, inplace=True)    
        
        formato(tabla_zona_sucursal, columnas_porcentaje, 'Porcentaje')
        formato(tabla_zona_sucursal, columnas_decimal, '1 Decimal')
        
        
        tabla_zona_canal = tabla_zona_canal.reindex(columnas_prueba, axis = 1)
        tabla_zona_canal.reset_index(level=0, inplace=True)  
        
        formato(tabla_zona_canal, columnas_porcentaje, 'Porcentaje')
        formato(tabla_zona_canal, columnas_decimal, '1 Decimal')    
        
        #Carga HTML ZONA
       
        env = Environment(loader = FileSystemLoader(''))
        template = env.get_template("template_zona.html")
        html = template.render(vista = zona,
                               fecha = fecha.strftime("%d-%m-%y"),
                               dotacion_vta = dot_vta,
                               dotacion_act = dot_act,
                               porcentaje_act = porc_act,
                               indicador_dot = logro_dot,
                               icon_dotacion = icon_dot,
                               total_UF = totalUF,
                               meta_UF = meta_UF,
                               porcentaje_UF = porc_cump_UF,
                               indicador_UF = logro_UF,
                               icon_UF = icon_UF,
                               asegurados_nuevos = asegurados_nuevos,
                               asegurados_totales = asegurados_totales,
                               porcentaje_Aseg_Nuevos = porc_cump_Aseg_Nuevos,
                               indicador_Aseg_Nuevos = logro_Aseg_Nuevos, 
                               icon_Aseg_Nuevos = icon_Aseg_Nuevos,                  
                               #online_UF = onlineUF,
                               porcentaje_online = porc_cump_online,
                               indicador_online = logro_online,
                               icon_online = icon_online,                        
                               aportes_UF= aportes,
                               #rp_UF = rp,
                               #presupuesto_rp = presupuesto_rp,
                               #porcentaje_rp= porc_cump_rp,
                               #indicador_rp= logro_rp,
                               #icon_rp = icon_rp,
                               grafico_canal=[familiares, mcg_ve, ap, so, mp10, cui],
                               meses_grafico = meses_grafico,
                               a_activo = a_activo,
                               #af_activo = af_activo,
                               a_con_venta = a_con_venta,
                               #af_con_venta = af_con_venta,    
                               tabla_canal = tabla_zona_canal.to_html(classes="table table-striped table-bordered compact hover formato_dt", index = False),
                               tabla_sucursal = tabla_zona_sucursal.to_html(classes="table-striped table-bordered compact hover formato_dt", index = False))
        
        ruta = "sitio/zonas/"+str(zona)+".html"
    
        with open(ruta, 'w', encoding="utf-8") as f:
            f.write(html)
            
            
#         print(rutas[metas['Zona'] == zona].reset_index())
        
        if cargar_informacion:            
            try:                    
                ruta_nas_zona = str(rutas[rutas['Zona'] == zona].reset_index().loc[0, 'Carga Zona']) + "\\" + str(fecha.strftime("%d-%m-%Y")) +"\\" + "Zonas"         
                if not os.path.exists(ruta_nas_zona):
                    os.makedirs(ruta_nas_zona)
                ruta_nas_zona = ruta_nas_zona + "\\" +str(zona) + ".html"
                with open(ruta_nas_zona, 'w', encoding="utf-8") as f:
                    f.write(html)    
            except:
                print("Zona: " + zona +" no encontrada")
        
    
    
    for sucursal in sucursales:
        #ocultar_af = "false"
        ocultar_a = "false"
        tabla_sucursal_canal =  pd.pivot_table(base,index=['Sucursal', 'Canal'],values=columnas_totales,aggfunc=np.sum)
        tabla_sucursal_equipo = pd.pivot_table(base,index=['Sucursal', 'Canal', 'Supervisor'],values=columnas_totales,aggfunc=np.sum)
        tabla_sucursal_equipo_int = pd.pivot_table(base, index=['Sucursal', 'I', 'Nombre', 'Canal', 'Supervisor'], values = columnas_totales, aggfunc=np.sum)
        tabla_sucursal_emitidas = pd.pivot_table(base_emision, index = ['Sucursal', 'Nro. Solicitud', 'Producto Estándar', 'Estado Solicitud', 'Nro. Poliza', 'Nombre Intermediario', 'Canal', 'Asegurado Nuevo', 'Supervisor', 'Observación', 'Pendiente de Emisión'], values = ['Prima Mensual'], aggfunc=np.sum)
        tabla_sucursal_ventas = pd.pivot_table(base_ventas, index = ['Agencia Aux', 'Nro. Solicitud', 'Producto Estándar', 'Estado Solicitud', 'Nro. Poliza', 'Nombre Intermediario', 'Canal', 'Asegurado Nuevo', 'Supervisor', 'Observación', 'Pendiente de Emisión'], values = ['Prima Mensual'], aggfunc=np.sum)
        
        
        
        tabla_sucursal_canal = tabla_sucursal_canal.xs(sucursal, level = 0)
     
    
        tabla_sucursal_canal['Tasa Vta.'] = tabla_sucursal_canal['Con Venta']/tabla_sucursal_canal['Activo']
        #tabla_sucursal_canal['Tasa Vta. No Presencial'] = tabla_sucursal_canal['Con Venta No Presencial']/tabla_sucursal_canal['Con Venta']
        tabla_sucursal_canal['% Var Mes Ant'] = round(tabla_sucursal_canal['Prima Mensual']/tabla_sucursal_canal['Prima Total Ant'] - 1,2)
        tabla_sucursal_canal['Tasa Aseg. Nuevos'] = tabla_sucursal_canal['Asegurados Nuevos']/tabla_sucursal_canal['Solicitudes Totales']
        
    #     tabla_sucursal_canal.replace([np.inf, -np.inf], np.nan).dropna(axis=1)
        
        #Indicadores
        #Dotacion
        dot_vta = tabla_sucursal_canal['Con Venta'].sum()
        dot_act = tabla_sucursal_canal['Activo'].sum()
        porcentaje = dot_vta/dot_act
        
        if porcentaje <  semaforo_rojo_dot:
            logro_dot = "text-danger"
            icon_dot = icono_abajo
        elif porcentaje < semaforo_amarillo_dot:
            logro_dot = "text-warning"
            icon_dot = icono_neutral
        else:
            logro_dot = "text-success"
            icon_dot = icono_arriba
            
        porc_act = str(round(porcentaje*100,1)) + "%"
    
        
        #Venta    
        metas_uf = objetivos[objetivos['Sucursal'] == sucursal]  ##filtrar metas por sucursal
        meta_UF = round(metas_uf['Meta Mes Ant'].sum(), 2)
                
        totalUF = round(tabla_sucursal_canal['Prima Mensual'].sum(), 2)
        
        porcentaje = totalUF/meta_UF
        
        if porcentaje <  semaforo_rojo_UF:
            logro_UF = "text-danger"
            icon_UF = icono_abajo
        elif porcentaje < semaforo_amarillo_UF:
            logro_UF = "text-warning"
            icon_UF = icono_neutral
        else:
            logro_UF = "text-success"
            icon_UF = icono_arriba
        
        porc_cump_UF = str(round(porcentaje*100,1)) + "%"
        
        #3 Cumplimiento asegurados nuevos
        asegurados_nuevos = tabla_sucursal_canal['Asegurados Nuevos'].sum()
        
        asegurados_totales = tabla_sucursal_canal['Solicitudes Totales'].sum()
        
        porcentaje = asegurados_nuevos/asegurados_totales
        
        
        if porcentaje <  semaforo_rojo_Aseg_Nuevos:
            logro_Aseg_Nuevos = "text-danger"
            icon_Aseg_Nuevos = icono_abajo
        elif porcentaje < semaforo_amarillo_Aseg_Nuevos:
            logro_Aseg_Nuevos = "text-warning"
            icon_Aseg_Nuevos = icono_neutral
        else:
            logro_Aseg_Nuevos = "text-success"
            icon_Aseg_Nuevos = icono_arriba        
        porc_cump_Aseg_Nuevos = str(round(porcentaje*100,1)) + "%"
        
        
    
        #Venta Online (comparada con venta total)
    
        #onlineUF = round(tabla_sucursal_canal['Prima No Presencial'].sum(), 2)
        totalUF = round(tabla_sucursal_canal['Prima Mensual'].sum(), 2)
        #porcentaje = onlineUF/totalUF
        
        if porcentaje <  semaforo_rojo_online:
            logro_online = "text-danger"
            icon_online = icono_abajo
        elif porcentaje < semaforo_amarillo_online:
            logro_online = "text-warning"
            icon_online = icono_neutral
        else:
            logro_online = "text-success"
            icon_online = icono_arriba
            
        porc_cump_online = str(round(porcentaje*100,1)) + "%"
    
        aportes = round(tabla_sucursal_canal['Aportes'].sum(), 2)
        
        #Meta RP sucursal
        #metas_rp = metas[metas['Sucursal'] == sucursal]  ##filtrar metas por sucursal
        #presupuesto_rp = round(metas_rp['Meta RP'].sum(), 2)
        
        #rp = round(tabla_sucursal_canal['RP'].sum(), 2)  
    
        #porcentaje = rp/presupuesto_rp
        
        #if porcentaje <  semaforo_rojo_rp:
            #logro_rp = "text-danger"
            #icon_rp = icono_abajo
        #elif porcentaje < semaforo_amarillo_rp:
            #logro_rp = "text-warning"
            #icon_rp = icono_neutral
        #else:
            #logro_rp = "text-success"
            #icon_rp = icono_arriba
            
        #porc_cump_rp = str(round(porcentaje*100,1)) + "%"
        
        #Gráfico Dotación
    
        tabla_a = tabla_sucursal_canal[np.in1d(tabla_sucursal_canal.index.get_level_values('Canal'), ['Asesor'])]    
        a_activo1 = tabla_a['Activo -2'].sum()
        a_activo2 = tabla_a['Activo -1'].sum()
        a_activo3 = tabla_a[col_activo3].sum()
        a_con_venta1 = tabla_a['Con Venta -2'].sum()
        a_con_venta2 = tabla_a['Con Venta -1'].sum()
        a_con_venta3 = tabla_a[col_con_venta3].sum()
        a_activo = [a_activo1, a_activo2, a_activo3]
        a_con_venta = [a_con_venta1, a_con_venta2, a_con_venta3]
        if tabla_a['Vigente'].sum() == 0:
            ocultar_a = "true";    
        
        #tabla_af = tabla_sucursal_canal[np.in1d(tabla_sucursal_canal.index.get_level_values('Canal'), ['Asesor Financiero'])]  
        #af_activo1 = tabla_af['Activo -2'].sum()
        #af_activo2 = tabla_af['Activo -1'].sum()
        #af_activo3 = tabla_af[col_activo3].sum()
        #af_con_venta1 = tabla_af['Con Venta -2'].sum()
        #af_con_venta2 = tabla_af['Con Venta -1'].sum()
        #af_con_venta3 = tabla_af[col_con_venta3].sum()
        #vaf_activo = [af_activo1, af_activo2, af_activo3]
        #af_con_venta = [af_con_venta1, af_con_venta2, af_con_venta3]
        #if tabla_af['Vigente'].sum() == 0:
            #ocultar_af = "true";
        
            
        
        #Mix de producto sucursal
        tabla_sucursal_mix = base_mix.xs(sucursal,level = 1)
        familiares = tabla_sucursal_mix['FAMILIAR'].sum()
        mcg_ve = tabla_sucursal_mix['MCG y VE'].sum()
        ap = tabla_sucursal_mix['AP'].sum()
        so = tabla_sucursal_mix['SO'].sum()
        mp10 = tabla_sucursal_mix['MP10'].sum()
        cui = tabla_sucursal_mix['CUI'].sum()      
        
        tabla_sucursal_canal = tabla_sucursal_canal.reindex(columnas_prueba, axis = 1)  
        tabla_sucursal_canal.reset_index(level=0, inplace=True) #Se ajusta tabla despues de calcular indicadores    
        formato(tabla_sucursal_canal, columnas_porcentaje, 'Porcentaje')
        formato(tabla_sucursal_canal, columnas_decimal, '1 Decimal')
    
    #     print(tabla_sucursal_equipo)
        
        tabla_sucursal_equipo = tabla_sucursal_equipo.xs(sucursal, level = 0)    
        tabla_sucursal_equipo['Tasa Vta.'] = tabla_sucursal_equipo['Con Venta']/tabla_sucursal_equipo['Activo']
       # tabla_sucursal_equipo['Tasa Vta. No Presencial'] = tabla_sucursal_equipo['Con Venta No Presencial']/tabla_sucursal_equipo['Con Venta']
        tabla_sucursal_equipo['% Var Mes Ant'] = round(tabla_sucursal_equipo['Prima Mensual']/tabla_sucursal_equipo['Prima Total Ant'] - 1,2)
        tabla_sucursal_equipo['Tasa Aseg. Nuevos'] = tabla_sucursal_equipo['Asegurados Nuevos']/tabla_sucursal_equipo['Solicitudes Totales']
        tabla_sucursal_equipo = tabla_sucursal_equipo.reindex(columnas_prueba, axis = 1)
        tabla_sucursal_equipo.reset_index(level=0, inplace=True)
        tabla_sucursal_equipo.reset_index(level=0, inplace=True) 
        
    #     tabla_sucursal_equipo.fillna(0)    
        
        formato(tabla_sucursal_equipo, columnas_porcentaje, 'Porcentaje')
        formato(tabla_sucursal_equipo, columnas_decimal, '1 Decimal')   
       
    #     tabla_sucursal_equipo_int.to_excel("Resultado4.xlsx")   
        tabla_sucursal_equipo_int = tabla_sucursal_equipo_int.xs(sucursal, level=0)
        tabla_sucursal_equipo_int['% Var Mes Ant'] = round(tabla_sucursal_equipo_int['Prima Mensual']/tabla_sucursal_equipo_int['Prima Total Ant'] - 1,2)
        tabla_sucursal_equipo_int['Tasa Aseg. Nuevos'] = round(tabla_sucursal_equipo_int['Asegurados Nuevos']/tabla_sucursal_equipo_int['Solicitudes Totales'], 2)
        tabla_sucursal_equipo_int = tabla_sucursal_equipo_int.reindex(columnas_int, axis = 1)
    
        tabla_sucursal_equipo_int.reset_index(inplace=True)
#         tabla_sucursal_equipo_int.reset_index(level=2, inplace=True)
#         tabla_sucursal_equipo_int.reset_index(level=1, inplace=True)
#         tabla_sucursal_equipo_int.reset_index(level=0, inplace=True)    
#     
    #     formato(tabla_sucursal_equipo_int, columnas_porcentaje, 'Porcentaje')
        formato(tabla_sucursal_equipo_int, columnas_decimal, '1 Decimal')
    
        try:                        
#             print("Sucursal: " + sucursal)
            tabla_sucursal_emitidas = tabla_sucursal_emitidas.xs(sucursal, level = 0)
            tabla_sucursal_emitidas.reset_index(inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=5, inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=4, inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=3, inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=2, inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=2, inplace=True)
#             tabla_sucursal_emitidas.reset_index(level=0, inplace=True)
#             print(tabla_sucursal_emitidas)
        except:
            tabla_sucursal_emitidas = pd.DataFrame(columns = ['Nro. Solicitud', 'Producto Estándar', 'Estado Solicitud', 'Nro. Poliza', 'Nombre Intermediario', 'Canal', 'Prima Mensual','Asegurado Nuevo', 'Supervisor', 'Observación', 'Pendiente de Emisión'])
            print("Sucursal " + sucursal +" no tiene pólizas emitidas")
#             print(tabla_sucursal_emitidas)
            
        try:                        
#             print("Sucursal: " + sucursal)
            tabla_sucursal_ventas = tabla_sucursal_ventas.xs(sucursal, level = 0)
            tabla_sucursal_ventas.reset_index(inplace=True)

        except:            
            print("Sucursal " + sucursal +" no tiene pólizas en proceso")
            tabla_sucursal_ventas = pd.DataFrame(columns = ['Nro. Solicitud', 'Producto Estándar', 'Estado Solicitud', 'Nro. Poliza', 'Nombre Intermediario', 'Canal', 'Prima Mensual','Asegurado Nuevo', 'Supervisor', 'Observación', 'Pendiente de Emisión'])
#             print(tabla_sucursal_ventas)

        #Unir base de ventas con base de emisión
        tabla_solicitudes = pd.concat([tabla_sucursal_emitidas, tabla_sucursal_ventas], ignore_index=True)
        
        
        env = Environment(loader = FileSystemLoader(''))
        template = env.get_template("template_sucursal.html")
        html = template.render(vista = sucursal,
                               fecha = fecha.strftime("%d-%m-%y"),
                               dotacion_vta = dot_vta,
                               dotacion_act = dot_act,
                               porcentaje_act = porc_act,
                               indicador_dot = logro_dot,
                               icon_dotacion = icon_dot,
                               total_UF = totalUF,
                               meta_UF = meta_UF,
                               porcentaje_UF = porc_cump_UF,
                               indicador_UF = logro_UF,
                               icon_UF = icon_UF,
                               asegurados_nuevos = asegurados_nuevos,
                               asegurados_totales = asegurados_totales,
                               porcentaje_Aseg_Nuevos = porc_cump_Aseg_Nuevos,
                               indicador_Aseg_Nuevos = logro_Aseg_Nuevos, 
                               icon_Aseg_Nuevos = icon_Aseg_Nuevos,                  
                               #online_UF = onlineUF,
                               porcentaje_online = porc_cump_online,
                               indicador_online = logro_online,
                               icon_online = icon_online,                        
                               aportes_UF= aportes,
                               #rp_UF = rp,
                               #presupuesto_rp = presupuesto_rp,
                               #porcentaje_rp= porc_cump_rp,
                               #indicador_rp= logro_rp,
                               #icon_rp = icon_rp,
                               grafico_canal=[familiares, mcg_ve, ap, so, mp10, cui],
                               meses_grafico = meses_grafico,
                               a_activo = a_activo,
                               #af_activo = af_activo,
                               a_con_venta = a_con_venta,
                               #af_con_venta = af_con_venta,
                               #ocultar_af = ocultar_af,
                               ocultar_a = ocultar_a,
                               tabla_canal = tabla_sucursal_canal.to_html(classes="table table-striped table-bordered compact hover", table_id='formato_suc_canal', index = False),
                               tabla_equipo = tabla_sucursal_equipo.to_html(classes="table table-striped table-bordered compact hover", table_id='formato_suc_equipo', index = False),
                               tabla_equipo_int = tabla_sucursal_equipo_int.to_html(classes="table table-striped table-bordered compact hover nowrap", table_id='formato_suc_int', index = False),
                               tabla_solicitudes = tabla_solicitudes.to_html(classes="table table-striped table-bordered compact hover nowrap", table_id='formato_solicitudes', index = False))
        
        ruta = "sitio/sucursales/"+str(sucursal)+".html"
        with open(ruta, 'w', encoding="utf-8") as f:
            f.write(html)  
        
#         print(rutas[metas['Sucursal'] == sucursal].reset_index())
        
        if cargar_informacion: 
            
            try: 
                print("Entre a cargar informacion")                
                ruta_nas_sucursal = str(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, 'Carga Sucursal'])
                if not os.path.exists(ruta_nas_sucursal):
                    os.makedirs(ruta_nas_sucursal)  
                ruta_nas_sucursal = str(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, 'Carga Sucursal']) + "\\" + str(fecha.strftime("%d-%m-%Y")) + "_" + str(sucursal) + ".html"                  
                    
                with open(ruta_nas_sucursal, 'w', encoding="utf-8") as f:
                    f.write(html)                
                #Pedir acceso a las zonas y luego quitar comentarios
                # ruta_nas_zona = str(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, 'Carga Zona']) + "\\" + str(fecha.strftime("%d-%m-%Y")) +"\\" + "Sucursales"         
                # if not os.path.exists(ruta_nas_zona):
                #     print("Entre al if")
                #     os.makedirs(ruta_nas_zona)
                #     print("Entre zona")
                # ruta_nas_zona = ruta_nas_zona + "\\" +str(sucursal) + ".html"
                # print(ruta_nas_zona)
                # with open(ruta_nas_zona, 'w', encoding="utf-8") as f:
                #     f.write(html)  
                #     print("Entre zona")
                equipos = int(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, 'Equipos'])
                print(sucursal)
                print(equipos)                
                if equipos > 0:
                    for i in range(equipos):
                        print("i =" + str(i))
                        equipo = "Equipo "+str(i+1)
                        print(equipo)
                        ruta_nas_equipo = str(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, equipo])
                        print(ruta_nas_equipo)
                        if not os.path.exists(ruta_nas_equipo):
                            os.makedirs(ruta_nas_equipo)
                        ruta_nas_equipo = str(rutas[rutas['Sucursal'] == sucursal].reset_index().loc[0, equipo]) + "\\" + str(fecha.strftime("%d-%m-%Y")) + "_" + str(sucursal) + ".html"
                        with open(ruta_nas_equipo, 'w', encoding="utf-8") as f:
                            f.write(html)                 
                                
            except:
                print("Sucursal: " + sucursal +" no encontrada")
        

        
        
            
            
            
        
    # 
    # print("PIVOT TABLE")
    # 
    # tabla_nacional = pd.pivot_table(base,
    #                                 index=['Z', 'Zona', 'S', 'Sucursal', 'Canal', 'Nombre'],
    #                                 values=columnas)
    # zonas = tabla_nacional.index.get_level_values(1).unique()
    # sucursales = tabla_nacional.index.get_level_values(3).unique()
    # fila_nacional = 0
    #  
    # with pd.ExcelWriter('Resultado2.xlsx',engine='xlsxwriter') as ew:    
    #     for zona in zonas:            
    #          
    #         fila_zona = 0
    #         tabla_zona = pd.pivot_table(base,index=['Zona', 'S', 'Sucursal'],values=columnas_totales,aggfunc=np.sum)
    #         tabla_zona['Tasa Vta.'] = tabla_zona['Con Venta']/tabla_zona['Activo']
    #         tabla_zona['Tasa Vta. Online'] = tabla_zona['Con Venta Online']/tabla_zona['Con Venta']
    #         tabla_zona['% Var Mes Ant'] = round(tabla_zona['Prima Mensual']/tabla_zona['Prima Total Ant'] - 1,2)        
    #          
    #         filtro_zona = 'Zona == [\"'+str(zona)+ '\"]'   
    #          
    #         tabla_zona = tabla_zona.query(filtro_zona) 
    #         tabla_zona = tabla_zona.reindex(columnas_prueba, axis = 1)
    #         tabla_zona.to_excel(ew,sheet_name='Nacional',startrow = fila_nacional, startcol=0)
    #         tabla_zona.to_excel(ew, sheet_name = zona, startrow = fila_zona)        
    #         sucursales_zona = tabla_zona.index.get_level_values(2).unique()
    #         fila_nacional = fila_nacional + len(tabla_zona) + 3
    #         fila_zona = fila_zona + len(tabla_zona) + 3
    #         for sucursal in sucursales_zona:
    #             tabla_sucursal = pd.pivot_table(base,index=['S', 'Sucursal', 'Canal', 'Nombre'],values=columnas_totales,aggfunc=np.sum)
    #             tabla_sucursal['Tasa Vta.'] = tabla_sucursal['Con Venta']/tabla_sucursal['Activo']
    #             tabla_sucursal['Tasa Vta. Online'] = tabla_sucursal['Con Venta Online']/tabla_sucursal['Con Venta']
    #             tabla_sucursal['% Var Mes Ant'] = round(tabla_sucursal['Prima Mensual']/tabla_sucursal['Prima Total Ant'] - 1,2)                
    #             print(tabla_sucursal)
    #             filtro_sucursal = 'Sucursal == [\"'+str(sucursal)+ '\"]' 
    #             print(filtro_sucursal) 
    #             tabla_sucursal = tabla_sucursal.query(filtro_sucursal)
    #             tabla_sucursal = tabla_sucursal.reindex(columnas_prueba, axis = 1)
    #              
    #             tabla_sucursal.to_excel(ew,sheet_name='Nacional',startrow = fila_nacional)
    #             tabla_sucursal.to_excel(ew, sheet_name = zona, startrow = fila_zona)
    #             tabla_sucursal.to_excel(ew, sheet_name = sucursal, startrow = 0)
    #              
    #             fila_zona = fila_zona + len(tabla_sucursal) + 3
    #             fila_nacional = fila_nacional + len(tabla_sucursal)+3            
    #         print(sucursales_zona)        
    #     print(tabla_zona.count)  
    #     print(tabla_zona)     
    # for sucursal in sucursales:
    #     print(tabla_nacional.xs(sucursal, level=3))
    # tabla_nacional.to_excel("Resultado2.xlsx", index = False, sheet_name="Dos")
        
    #RENTAS PRIVADAS
    
#if generar_rp:
    #Definiciones
    #columnas_rp: considera las columnas y el orden de las mismas que tendrá finalmente el reporte 
    #base_rp_ingresos_mes: Considera tabla dinamica con todos los ingresos del mes. (se debe ajustar la fecha)
    #base_rp_renov_vctos_mes: Considera solo las renovaciones de los vencimientos del mes (todas las columnas)
    #rp_renov_vctos_mes: Considera las renovaciones de los vencimientos del mes ordenados por zona y considerando solo la prima.
    #base_rp_vencimientos_mes: Filtra desde la hoja Base.Vencimientos los vencimientos del mes (se debe ajustar la fecha)
    #meta_rp_zona: Consideras las metas ordenadas por zona.
    #rp_zona: Es la base que se cargará en el reporte finalmente.  
    
    
    #print("Generando Informe de Rentas Privadas")
    
    #Carga de archivos
    #base_rp_ingresos = pd.read_excel('BaseInformeRP.xlsx',"Base.Ingresos 2020")
    #base_rp_vencimientos = pd.read_excel('BaseInformeRP.xlsx',"Base.Vencimientos 2020")  

   # columnas_rp_todas = ['Nuevas UF', 'Meta Nuevas', '% Cump. Nuevas', 'Renovación UF', 'Meta Renovación', '% Cump. Renov.', 'Vencimientos mes', 'Tasa Renov.', 'Nro. Rentas', 'Total UF', '% Cump. Meta' ]
   # columnas_rp = ['Nro. Rentas', 'Nuevas UF', 'Meta Nuevas', 'Renovación UF', 'Meta Renovación','Total UF', 'Meta','% Cump. Nuevas','% Cump. Renov.','% Cump. Meta','Tasa Renov.']
    
    #Tablas auxiliares
    #base_rp_ingresos_mes = base_rp_ingresos[(base_rp_ingresos['Fecha Ingreso Solicitud Renta']>= datetime.datetime(2020,7,1))]    
    #base_rp_renov_vctos_mes = base_rp_ingresos[(base_rp_ingresos['Fecha Ingreso Solicitud Renta']>= datetime.datetime(2020,7,1))
                                         # & (base_rp_ingresos['Fecha Vencimiento']>=datetime.datetime(2020,7,1))] #Renovaciones de los vencimientos del mes
    
    #ajustar para que sean los vencimientos a la fecha - 2 días
    #base_rp_vencimientos_mes = base_rp_vencimientos[(base_rp_vencimientos['Fecha Vencimiento']>= datetime.datetime(2020,7,1)) & ((base_rp_vencimientos['Fecha Vencimiento']<=datetime.datetime(2020,7,31))| base_rp_vencimientos['Nro RP']>0)]
    
    #ZONAS   
    #print("Generando tablas por zona")     
    
    #Cálculo columna de prima renovada asociada a vencimientos del mes
    #rp_renov_vctos_mes = pd.pivot_table(base_rp_renov_vctos_mes,
                                  #index=['Z', 'Zona'],
                                  #values=['Prima UF'],
                                  #aggfunc='sum',
                                  #fill_value=0)
    #rp_renov_vctos_mes.columns = ['Prima Renovada mes']    
    
    #Tabla zonas
    #rp_zona = pd.pivot_table(base_rp_ingresos_mes,
                             #index=['Z', 'Zona'],
                             #columns=['Origen Renta'],
                             #values =['Prima UF'],
                             #aggfunc=('count', 'sum'))    
    #rp_zona.columns = ['Nuevas', 'Renovación', 'Nuevas UF', 'Renovación UF']
    
#     print(rp_zona)
    #meta_rp_zona = pd.pivot_table(metas,
                             #index=['Z','Zona'],
                             #values =['Meta Renovación', 'Meta Nuevas'],
                             #aggfunc='sum',
                             #fill_value=0)
    
    #Cruzar con metas: En la tabla de metas estan todas las zonas, en la tabla de ingresos podrian faltar
    #rp_zona = pd.merge(rp_zona, meta_rp_zona, on=['Z', 'Zona'], how='right')
    
    #Vencimientos por zona
    #vencimientos_rp_zona = pd.pivot_table(base_rp_vencimientos_mes,
                                          #index=['Z', 'Zona'],
                                          #values=['Renta UF'],
                                          #aggfunc = 'sum',
                                          #fill_value = 0)
    #vencimientos_rp_zona.columns = ['Vencimientos mes']
    
    
    #Cruzar con vencimientos (y renovaciones de los vencimientos)
    #rp_zona = pd.merge(rp_zona, vencimientos_rp_zona, on=['Z', 'Zona'], how='left') 
    #rp_zona = pd.merge(rp_zona, rp_renov_vctos_mes, on=['Z', 'Zona'], how='left')    
#     print(rp_zona)
    
    #Cálculo de nuevas columnas
    #rp_zona['Nro. Rentas'] = rp_zona['Nuevas']+rp_zona['Renovación']
    #rp_zona['Total UF'] = rp_zona['Nuevas UF'] + rp_zona['Renovación UF']
    #rp_zona['Meta'] = rp_zona['Meta Renovación'] + rp_zona['Meta Nuevas']
    #rp_zona['% Cump. Nuevas'] = round(rp_zona['Nuevas UF']/rp_zona['Meta Nuevas'],2)
    #rp_zona['% Cump. Renov.'] = round(rp_zona['Renovación UF']/rp_zona['Meta Renovación'],2)
    #rp_zona['Tasa Renov.'] = round(rp_zona['Prima Renovada mes']/rp_zona['Vencimientos mes'],2).fillna(0)
    #rp_zona['% Cump. Meta'] = round(rp_zona['Total UF']/rp_zona['Meta'],2)
  

    #SUCURSALES
    #print("Generando tablas por sucursal")
    
    #rp_renov_vctos_mes = pd.pivot_table(base_rp_renov_vctos_mes,
                                  #index=['S', 'Sucursal'],
                                  #values=['Prima UF'],
                                  #aggfunc='sum',
                                  #fill_value=0)
    #rp_renov_vctos_mes.columns = ['Prima Renovada mes']

    #rp_sucursal = pd.pivot_table(base_rp_ingresos_mes,
                             #index=['S', 'Sucursal'],
                             #columns=['Origen Renta'],
                             #values =['Prima UF'],
                             #aggfunc=('count', 'sum'))    
    #rp_sucursal.columns = ['Nuevas', 'Renovación', 'Nuevas UF', 'Renovación UF']

    #meta_rp_sucursal = pd.pivot_table(metas,
                         #index=['S', 'Sucursal'],
                         #values =['Meta Renovación', 'Meta Nuevas'],
                         #aggfunc='sum',
                         #fill_value=0)
    
    #rp_sucursal = pd.merge(rp_sucursal, meta_rp_sucursal, on=['S', 'Sucursal'], how='right')
    
    #vencimientos_rp_sucursal = pd.pivot_table(base_rp_vencimientos_mes,
                                          #index=['S', 'Sucursal'],
                                          #values=['Renta UF'],
                                          #aggfunc = 'sum',
                                          #fill_value = 0)
    #vencimientos_rp_sucursal.columns = ['Vencimientos mes']
    
    #rp_sucursal = pd.merge(rp_sucursal, vencimientos_rp_sucursal, on=['S', 'Sucursal'], how='left') 
    #rp_sucursal = pd.merge(rp_sucursal, rp_renov_vctos_mes, on=['S', 'Sucursal'], how='left')  
    
    #rp_sucursal['Nro. Rentas'] = rp_sucursal['Nuevas']+rp_sucursal['Renovación']
    #rp_sucursal['Total UF'] = rp_sucursal['Nuevas UF'] + rp_sucursal['Renovación UF']
    #rp_sucursal['Meta'] = rp_sucursal['Meta Renovación'] + rp_sucursal['Meta Nuevas']
    #rp_sucursal['% Cump. Nuevas'] = round(rp_sucursal['Nuevas UF']/rp_sucursal['Meta Nuevas'],2)
    #rp_sucursal['% Cump. Renov.'] = round(rp_sucursal['Renovación UF']/rp_sucursal['Meta Renovación'],2)
    #rp_sucursal['Tasa Renov.'] = round(rp_sucursal['Prima Renovada mes']/rp_sucursal['Vencimientos mes'],2).fillna(0)
    #rp_sucursal['% Cump. Meta'] = round(rp_sucursal['Total UF']/rp_sucursal['Meta'],2)
    

    
    #Cálculo de indicadores
    
    #1 Cumplimiento de meta
    #uf_total = rp_zona['Total UF'].sum()
    #meta_total = rp_zona['Meta'].sum()    
    #porcentaje = uf_total/meta_total
    
    #if porcentaje <  semaforo_rojo_rp:
        #logro_rp = "text-danger"
        #icon_rp = icono_abajo
    #elif porcentaje < semaforo_amarillo_rp:
        #logro_rp = "text-warning"
        #icon_rp = icono_neutral
    #else:
        #logro_rp = "text-success"
        #icon_rp = icono_arriba
    
    #porcentaje_cump = str(round(porcentaje*100,1)) + "%"
    
    #2 Meta de Rentas Nuevas
    #uf_nuevas = rp_zona['Nuevas UF'].sum()
    #meta_nuevas = rp_zona['Meta Nuevas'].sum()    
    #porcentaje = uf_nuevas/meta_nuevas
    
    #if porcentaje <  semaforo_rojo_rp:
        #logro_rp_nuevas = "text-danger"
        #icon_rp_nuevas = icono_abajo
    #elif porcentaje < semaforo_amarillo_rp:
        #logro_rp_nuevas = "text-warning"
        #icon_rp_nuevas = icono_neutral
    #else:
        #logro_rp_nuevas = "text-success"
        #icon_rp_nuevas = icono_arriba
    
    #porcentaje_cump_nuevas = str(round(porcentaje*100,1)) + "%"
    
    #3 Meta de Renovación
    #uf_renov = rp_zona['Renovación UF'].sum()
    #meta_renov = rp_zona['Meta Renovación'].sum()    
    #porcentaje = uf_renov/meta_renov
    
    #if porcentaje <  semaforo_rojo_rp:
        #logro_rp_renov = "text-danger"
        #icon_rp_renov = icono_abajo
    #elif porcentaje < semaforo_amarillo_rp:
        #logro_rp_renov = "text-warning"
        #icon_rp_renov = icono_neutral
    #else:
        #logro_rp_renov = "text-success"
        #icon_rp_renov = icono_arriba
    
    #porcentaje_cump_renov = str(round(porcentaje*100,1)) + "%"
    
    #4 Tasa de Renovación
    #venc_mes = rp_zona['Vencimientos mes'].sum()
    #prima_renovada_mes = rp_zona['Prima Renovada mes'].sum()
    #porcentaje = prima_renovada_mes/venc_mes
    
    #if porcentaje <  semaforo_rojo_renov_rp:
        #logro_rp = "text-danger"
        #icon_rp = icono_abajo
    #elif porcentaje < semaforo_amarillo_renov_rp:
        #logro_rp = "text-warning"
        #icon_rp = icono_neutral
    #else:
        #logro_rp = "text-success"
        #icon_rp = icono_arriba
        
    #tasa_renov = str(round(porcentaje*100,1)) + "%"

    #Eliminar columnas no utilizadas
    #rp_zona = rp_zona[columnas_rp]
    #rp_zona.reset_index(level=1, inplace=True)
    #rp_zona.reset_index(level=0, inplace=True)
    #rp_zona.fillna(0, inplace=True)
    
    #rp_sucursal = rp_sucursal[columnas_rp]
    #rp_sucursal.reset_index(level=1, inplace=True)
    #rp_sucursal.reset_index(level=0, inplace=True)
    #rp_sucursal.fillna(0, inplace=True) 
