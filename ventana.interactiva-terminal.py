import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import psycopg2
import csv
import datetime
from tkinter import messagebox, simpledialog
import pymysql
from openpyxl import Workbook
from datetime import datetime
## python -m PyInstaller ventana.interactiva.py    --- CODIGO PARA CREAR EJECUTABLE

usuari = None
pg_params = {
        "host": "192.168.137.1",
        "database": "ROBOT_VB",
        "user": "postgres",
        "password": "1030571984"
    }


def ventana_inicial():

    global usuari
    global inicial
    

    inicial = tk.Tk()
    inicial.title("Inicio de secion")
    inicial.geometry("300x150")

    tit_user = tk.Label(inicial, text="Usuario: ")
    tit_user.pack()

    usuari = tk.Entry(inicial)
    usuari.pack(pady=5)

    acceso = tk.Button(inicial,text="Ingresar", command=ventana_principal)
    acceso.pack()
   
    inicial.mainloop()


def ventana_principal():
    global inicial
    global latitud
    global longitud
    global ciudad
    global direccion
    global tabla_respuesta
    

    user = usuari.get()

    if user == '':
        messagebox.showinfo("Advertencia","Usuario no definido")
        return

    if user != 'Flancherospi' and  user != 'Daniela' :
        messagebox.showinfo("Advertencia","Usuario no autorizado")
        return

    # 👉 1. Ocultar la ventana inicial, NO destruirla
    inicial.withdraw()

    # 👉 2. Crear la ventana principal correctamente
    principal = tk.Toplevel()
    principal.title("Ventana de acceso a proceso Robot VB")
    principal.geometry("1300x500")

    # 👉 3. Ahora sí es seguro crear widgets
    notebook = ttk.Notebook(principal)
    notebook.pack(expand=True, fill="both")

    pestaña1 = tk.Frame(notebook, bg="#032733")
    pestaña2 = tk.Frame(notebook, bg="#084308")

    if user == 'Flancherospi':
        pestaña3 = tk.Frame(notebook, bg="#0e007b")
        notebook.add(pestaña3, text="Proceso VB con carga")
        tk.Label(pestaña3, text="CONSULTA PUNTO A PUNTO", bg="#032733", fg="white",font=("Arial Narrow", 20)).pack()

    notebook.add(pestaña1, text="Consulta Punto a Punto")
    notebook.add(pestaña2, text="Proceso VB")
   

    tk.Label(pestaña1, text="CONSULTA PUNTO A PUNTO", bg="#032733", fg="white", font=("Arial Narrow", 20)).pack(pady=10)

    tk.Label(pestaña2, text="ROBOT VIABILIDADES V.2", bg="#084308", fg="white", font=("Arial Narrow", 20) ).pack(pady=20)
    btn_inicio = tk.Button(pestaña2, text="Robot VB", command=proceso_robot_consulta, width=20, height=2, border=5, borderwidth=8 )
    btn_inicio.pack()

    

    tk.Label(pestaña1, text="LATITUD", bg="#032733",fg="white", font=("Arial",13)).pack()
    vcmd = pestaña1.register(validar_latitud)
    latitud = tk.Entry(pestaña1, width=20, font=("Arial", 18), validate="key", validatecommand=(vcmd, "%P"))
    latitud.pack()

    tk.Label(pestaña1,text="LONGITUD", bg="#032733",fg="white", font=("Arial",13)).pack()
    vcmd = pestaña1.register(validar_longitud)
    longitud = tk.Entry(pestaña1, width=20, font=("Arial", 18), validate="key", validatecommand=(vcmd, "%P"))
    longitud.pack()

    tk.Label(pestaña1,text="CIUDAD", bg="#032733",fg="white", font=("Arial",13)).pack()
    ciudad = tk.Entry(pestaña1, width=40, font=("Arial", 18))
    ciudad.pack()

    tk.Label(pestaña1,text="DIRECCION", bg="#032733",fg="white", font=("Arial",13)).pack()
    direccion = tk.Entry(pestaña1, width=40, font=("Arial", 18))
    direccion.pack()
    consulta = tk.Button(pestaña1, text="Consulta", command=consulta_por_punto)
    consulta.pack(pady=10)

    latitud.bind("<KeyRelease>", formatear_latitud)
    longitud.bind("<KeyRelease>", formatear_longitud)
    # consulta de creacion de tablas para consulta punto a punto.


    tabla_respuesta = ttk.Treeview(
        pestaña1,
        columns=("LAT_CONSULTADA","LONG_CONSULTADA","DIRECCION","LAT_CTO","LONG_CTO","ID_CTO","OPC_CTO","DISTANCIA_CTO","OP_FEEDER","PUERTOS_DISPONIBLES","OLT","CONSULTA"),
        show="headings",
        height=3

    )

    tabla_respuesta.heading("LAT_CONSULTADA", text="LAT_CONSULTADA")
    tabla_respuesta.column("LAT_CONSULTADA", width=90)
    tabla_respuesta.heading("LONG_CONSULTADA", text="LONG_CONSULTADA")
    tabla_respuesta.column("LONG_CONSULTADA", width=90)
    tabla_respuesta.heading("DIRECCION", text="DIRECCION")
    tabla_respuesta.column("DIRECCION", width=90)
    tabla_respuesta.heading("LAT_CTO", text="LAT_CTO")
    tabla_respuesta.column("LAT_CTO", width=90)
    tabla_respuesta.heading("LONG_CTO", text="LONG_CTO")
    tabla_respuesta.column("LONG_CTO", width=90)
    tabla_respuesta.heading("ID_CTO", text="ID_CTO")
    tabla_respuesta.column("ID_CTO", width=90)
    tabla_respuesta.heading("OPC_CTO", text="OPC_CTO")
    tabla_respuesta.column("OPC_CTO", width=90)
    tabla_respuesta.heading("DISTANCIA_CTO", text="DISTANCIA_CTO")
    tabla_respuesta.column("DISTANCIA_CTO", width=90)
    tabla_respuesta.heading("OP_FEEDER", text="OP_FEEDER")
    tabla_respuesta.column("OP_FEEDER", width=90)
    tabla_respuesta.heading("PUERTOS_DISPONIBLES", text="PUERTOS_DISPONIBLES")
    tabla_respuesta.column("PUERTOS_DISPONIBLES", width=90)
    tabla_respuesta.heading("OLT", text="OLT")
    tabla_respuesta.column("OLT", width=90)
    tabla_respuesta.heading("CONSULTA", text="CONSULTA")
    tabla_respuesta.column("CONSULTA", width=90)

    tabla_respuesta.pack()

def limpiar_tabla():
    global tabla_respuesta
    tabla_respuesta.delete(*tabla_respuesta.get_children())
    
def consulta_coordenada_zc():
    global pg_params

    # Configuración PostgreSQL
    pg_params = {
        "host": "192.168.137.1",
        "database": "ROBOT_VB",
        "user": "postgres",
        "password": "1030571984"
    }

    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
        print("Conexion exitosa a base PosgreSQL")
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL zc")

    
    consulta_zc = """ INSERT INTO respuesta (
                        id_pre,
                        anio,
                        mes,
                        estadovb,
                        tipo,
                        nit,
                        ds,
                        cliente,
                        proyecto,
                        departamento,
                        ciudad,
                        direccion,
                        complemento,
                        nombre_edifi,
                        zona,
                        latitud,
                        longitud,
                        tipo_servicio,
                        pb,
                        ancho_banda,
                        op1_coinversor,
                        op1_central,
                        op1olt,
                        op1feeder,
                        id_cto1,
                        op1cto,
                        op1distanciacto,
                        op1dipscto,
                        op2_coinversor,
                        op2_central,
                        op2feeder,
                        id_cto2,
                        op2cto,
                        op2distanciacto,
                        oc2dipscto,
                        fecha_solicitud,
                        lat_cto,
                        long_cto,
                        grupo_cto
                    )
                    SELECT 
                        p.id_pre AS id_proyecto,
                        EXTRACT(YEAR FROM CURRENT_DATE),
                        EXTRACT(MONTH FROM CURRENT_DATE),
                        NULL AS estadovb,
                        NULL AS tipo,
                        NULL AS nit,
                        p.ds,
                        p.nombre_cli,
                        P.proyecto,
                        P.departamento,
                        p.ciudad,
                        p.direccion,
                        p.complemento,
                        p.nombre_edifi,
                        null as zona,
                        p.latitud AS lat_proyecto,
                        p.longitud AS lon_proyecto,
                        p.producto as tipo_servicio,
                        p.tipo_enlace as pb,
                        p.ancho_banda,
                        o.proveedor AS op1_coinversor,
                        NULL AS op1_central,
                        o.olt_no AS op1olt,
                        CONCAT('CA', o.cable) AS op1feeder,
                        o.idplaca AS id_cto1,
                        o.cto AS cto_osp,
                        ST_Distance(p.geom, o.geom)::numeric(10,2) AS distancia_metros,
                        o.puertos_libres_cto as puertos_libres,
                        '' AS op2_coinversor,
                        '' AS op2_central,
                        '' AS op2feeder,
                        '' AS id_cto2,
                        '' AS op2cto,
                        '' AS op2distanciacto,
                        '' AS oc2dipscto,
                        CURRENT_DATE AS fecha_solicitud,
                        o.lat_equipo AS lat_osp,
                        o.long_equipo AS lon_osp,
                        'ZC'
                    FROM proyecto p
                    JOIN LATERAL (
                        SELECT *
                        FROM (
                            SELECT 
                                o.*,
                                ROW_NUMBER() OVER (
                                    PARTITION BY o.cto 
                                    ORDER BY p.geom <-> o.geom
                                ) as rn
                            FROM osp_3gys o
                            WHERE 
                                o.grupo_cto  <> 'Zona Abierta'
                                AND ST_DWithin(p.geom::geography, o.geom::geography, 50)
                                and o.puertos_libres_cto > 1
                                and p.respuesta is null
                        ) t
                        WHERE rn = 1   -- solo la más cercana por cada CTO
                        ORDER BY p.geom <-> t.geom
                        LIMIT 2        -- ahora sí: 2 CTO diferentes
                    ) o ON TRUE
                   """
    
    try:
        cur_pg.execute(consulta_zc)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Punto ZC con distancia menos a 400 metros",hora2)
    except Exception as e:
        print("Error en consulta de zona cerrada",e)
def consulta_coordenada_za():
    global pg_params
    # Configuración PostgreSQL
    #pg_params = {
    #    "host": "192.168.137.1",
    #    "database": "ROBOT_VB",
    #    "user": "postgres",
    #    "password": "1030571984"
    #}

    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
        print("Conexion exitosa a base PosgreSQL")
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL az")

    
    consulta_za = """ INSERT INTO respuesta (
                    id_pre,
                    anio,
                    mes,
                    estadovb,
                    tipo,
                    nit,
                    ds,
                    cliente,
                    proyecto,
                    departamento,
                    ciudad,
                    direccion,
                    complemento,
                    nombre_edifi,
                    zona,
                    latitud,
                    longitud,
                    tipo_servicio,
                    pb,
                    ancho_banda,
                    op1_coinversor,
                    op1_central,
                    op1olt,
                    op1feeder,
                    id_cto1,
                    op1cto,
                    op1distanciacto,
                    op1dipscto,
                    op2_coinversor,
                    op2_central,
                    op2feeder,
                    id_cto2,
                    op2cto,
                    op2distanciacto,
                    oc2dipscto,
                    fecha_solicitud,
                    lat_cto,
                    long_cto,
                    grupo_cto
                )
                SELECT 
                    p.id_pre AS id_proyecto,
                    EXTRACT(YEAR FROM CURRENT_DATE),
                    EXTRACT(MONTH FROM CURRENT_DATE),
                    NULL AS estadovb,
                    NULL AS tipo,
                    NULL AS nit,
                    p.ds,
                    p.nombre_cli,
                    P.proyecto,
                    P.departamento,
                    p.ciudad,
                    p.direccion,
                    p.complemento,
                    p.nombre_edifi,
                    null as zona,
                    p.latitud AS lat_proyecto,
                    p.longitud AS lon_proyecto,
                    p.producto as tipo_servicio,
                    p.tipo_enlace as pb,
                    p.ancho_banda,
                    o.proveedor AS op1_coinversor,
                    NULL AS op1_central,
                    o.olt_no AS op1olt,
                    CONCAT('CA', o.cable) AS op1feeder,
                    o.idplaca AS id_cto1,
                    o.cto AS cto_osp,
                    ST_Distance(p.geom, o.geom) * 1.5::numeric(10,2) AS distancia_metros,
                    o.puertos_libres_cto as puertos_libres,
                    '' AS op2_coinversor,
                    '' AS op2_central,
                    '' AS op2feeder,
                    '' AS id_cto2,
                    '' AS op2cto,
                    '' AS op2distanciacto,
                    '' AS oc2dipscto,
                    CURRENT_DATE AS fecha_solicitud,
                    o.lat_equipo AS lat_osp,
                    o.long_equipo AS lon_osp,
                    'ZA'
                FROM proyecto p
                JOIN LATERAL (
                        SELECT *
                        FROM (
                            SELECT 
                                o.*,
                                ROW_NUMBER() OVER (
                                    PARTITION BY o.cto 
                                    ORDER BY p.geom <-> o.geom
                                ) as rn
                            FROM osp_3gys o
                            WHERE 
                                ST_DWithin(p.geom::geography, o.geom::geography, 600)
                                AND o.grupo_cto LIKE '%Abierta%'
                                AND o.puertos_libres_cto > 1
                                AND p.respuesta IS NULL
                        ) t
                        WHERE rn = 1   -- solo la más cercana por cada CTO
                        ORDER BY p.geom <-> t.geom
                        LIMIT 2        -- ahora sí: 2 CTO diferentes
                    ) o ON TRUE
                   """
    
    try:
        cur_pg.execute(consulta_za)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Proceso ZA con distancia a 500 metros terminada.")

        
    except Exception as e:
        print("Error en consulta de zona abierta",e)

    limpieza_duplicados = """ DELETE FROM respuesta a
                                USING (
                                    SELECT 
                                        MIN(ctid) AS ctid_keep,
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    FROM respuesta
                                    GROUP BY 
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    HAVING COUNT(*) > 1
                                ) b
                                WHERE
                                    a.id_pre = b.id_pre
                                    AND a.lat_cto = b.lat_cto
                                    AND a.long_cto = b.long_cto
                                    AND a.op1distanciacto = b.op1distanciacto
                                    AND a.ctid <> b.ctid_keep; """
    
    try:
        cur_pg.execute(limpieza_duplicados)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        messagebox.showinfo("Advertencia", f"Proceso Terminado puntos Viabilizados.\nHora: {hora2}")

        
    except Exception as e:
        print("Error en funcion limpieza de duplicados",e)
def consulta_coordenada_emp():
    global pg_params
    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
       
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL emp")

    
    consulta_emp = """ INSERT INTO respuesta (
                    id_pre,
                    anio,
                    mes,
                    estadovb,
                    tipo,
                    nit,
                    ds,
                    cliente,
                    proyecto,
                    departamento,
                    ciudad,
                    direccion,
                    complemento,
                    nombre_edifi,
                    zona,
                    latitud,
                    longitud,
                    tipo_servicio,
                    pb,
                    ancho_banda,
                    op1_coinversor,
                    op1_central,
                    op1olt,
                    op1feeder,
                    id_cto1,
                    op1cto,
                    op1distanciacto,
                    op1dipscto,
                    op2_coinversor,
                    op2_central,
                    op2feeder,
                    id_cto2,
                    op2cto,
                    op2distanciacto,
                    oc2dipscto,
                    fecha_solicitud,
                    lat_cto,
                    long_cto,
                    grupo_cto
                )
                SELECT 
                    p.id_pre AS id_proyecto,
                    EXTRACT(YEAR FROM CURRENT_DATE),
                    EXTRACT(MONTH FROM CURRENT_DATE),
                    NULL AS estadovb,
                    NULL AS tipo,
                    NULL AS nit,
                    p.ds,
                    p.nombre_cli,
                    P.proyecto,
                    P.departamento,
                    p.ciudad,
                    p.direccion,
                    p.complemento,
                    p.nombre_edifi,
                    null as zona,
                    p.latitud AS lat_proyecto,
                    p.longitud AS lon_proyecto,
                    p.producto as tipo_servicio,
                    p.tipo_enlace as pb,
                    p.ancho_banda,
                    o.propietario AS op1_coinversor,
                    '-' AS op1_central,
                    '-' AS op1olt,
                    '-' AS op1feeder,
                    o.terminal_fibra_optica_id AS id_cto1,
                    o.terminal_fibra_optica_codigo AS cto_osp,
                    ST_Distance(p.geom, o.geom) AS distancia_metros,
                    '-' as puertos_libres,
                    '-' AS op2_coinversor,
                    '' AS op2_central,
                    '' AS op2feeder,
                    '' AS id_cto2,
                    '' AS op2cto,
                    '' AS op2distanciacto,
                    '' AS oc2dipscto,
                    CURRENT_DATE AS fecha_solicitud,
                    o.coordenada_x AS lat_osp,
                    o.coordenada_y AS lon_osp,
                    'EMP'
                FROM proyecto p
                JOIN LATERAL (
                    SELECT *
                    FROM (
                        SELECT 
                            o.*,
                            ROW_NUMBER() OVER (
                                PARTITION BY o.terminal_fibra_optica_codigo 
                                ORDER BY p.geom <-> o.geom
                            ) as rn
                        FROM emp_v2 o
                        WHERE 
                            ST_DWithin(p.geom::geography, o.geom::geography, 1000)
                            AND o.punto_acceso_tipo_punto_acceso like '%Cámara%'
                    ) t
                    WHERE rn = 1   -- solo la más cercana por cada CTO
                    ORDER BY p.geom <-> t.geom
                    LIMIT 2        -- ahora sí: 2 CTO diferentes
                ) o ON TRUE
                """
    
    try:
        cur_pg.execute(consulta_emp)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Proceso EMP con distancia a 1.000 metros terminada.")

        
    except Exception as e:
        print("Error en consulta de emp",e)

    limpieza_duplicados = """ DELETE FROM respuesta a
                                USING (
                                    SELECT 
                                        MIN(ctid) AS ctid_keep,
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    FROM respuesta
                                    GROUP BY 
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    HAVING COUNT(*) > 1
                                ) b
                                WHERE
                                    a.id_pre = b.id_pre
                                    AND a.lat_cto = b.lat_cto
                                    AND a.long_cto = b.long_cto
                                    AND a.op1distanciacto = b.op1distanciacto
                                    AND a.ctid <> b.ctid_keep; """
    
    try:
        cur_pg.execute(limpieza_duplicados)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        messagebox.showinfo("Advertencia", f"Proceso Terminado puntos Viabilizados.\nHora: {hora2}")

        
    except Exception as e:
        print("Error en consulta de zona abierta",e)
def consulta_coordenada_nodos():
    global pg_params
    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
       
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL nodos")

    
    consulta_nod = """ INSERT INTO respuesta (
                    id_pre,
                    anio,
                    mes,
                    estadovb,
                    tipo,
                    nit,
                    ds,
                    cliente,
                    proyecto,
                    departamento,
                    ciudad,
                    direccion,
                    complemento,
                    nombre_edifi,
                    zona,
                    latitud,
                    longitud,
                    tipo_servicio,
                    pb,
                    ancho_banda,
                    op1_coinversor,
                    op1_central,
                    op1olt,
                    op1feeder,
                    id_cto1,
                    op1cto,
                    op1distanciacto,
                    op1dipscto,
                    op2_coinversor,
                    op2_central,
                    op2feeder,
                    id_cto2,
                    op2cto,
                    op2distanciacto,
                    oc2dipscto,
                    fecha_solicitud,
                    lat_cto,
                    long_cto,
                    grupo_cto
                )
                SELECT 
                    p.id_pre AS id_proyecto,
                    EXTRACT(YEAR FROM CURRENT_DATE),
                    EXTRACT(MONTH FROM CURRENT_DATE),
                    NULL AS estadovb,
                    NULL AS tipo,
                    NULL AS nit,
                    p.ds,
                    p.nombre_cli,
                    P.proyecto,
                    P.departamento,
                    p.ciudad,
                    p.direccion,
                    p.complemento,
                    p.nombre_edifi,
                    null as zona,
                    p.latitud AS lat_proyecto,
                    p.longitud AS lon_proyecto,
                    p.producto as tipo_servicio,
                    p.tipo_enlace as pb,
                    p.ancho_banda,
                    'TEF' AS op1_coinversor,
                    '-' AS op1_central,
                    '-' AS op1olt,
                    '-' AS op1feeder,
                    o.cod_mov AS id_cto1,
                    o.hl AS cto_osp,
                    ST_Distance(p.geom, o.geom) AS distancia_metros,
                    '-' as puertos_libres,
                    '-' AS op2_coinversor,
                    '' AS op2_central,
                    '' AS op2feeder,
                    '' AS id_cto2,
                    '' AS op2cto,
                    '' AS op2distanciacto,
                    '' AS oc2dipscto,
                    CURRENT_DATE AS fecha_solicitud,
                    o.lat AS lat_osp,
                    o.lon AS lon_osp,
                    'nodo_TEF'
                FROM proyecto p
                JOIN LATERAL (
                    SELECT *
                    FROM (
                        SELECT 
                            o.*,
                            ROW_NUMBER() OVER (
                                PARTITION BY o.nombre_equipo 
                                ORDER BY p.geom <-> o.geom
                            ) as rn
                        FROM nodo_central o
                        WHERE 
                            ST_DWithin(p.geom::geography, o.geom::geography, 50000)
                            
                    ) t
                    WHERE rn = 1   -- solo la más cercana por cada CTO
                    and tx <> 'RADIO'
                    ORDER BY p.geom <-> t.geom
                    LIMIT 2       -- ahora sí: 2 NODOS diferentes
                ) o ON TRUE
                """
    
    
    
    try:
        cur_pg.execute(consulta_nod)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Proceso de nodos ")

        
    except Exception as e:
        print("Error en consulta de emp",e)

    limpieza_duplicados = """ DELETE FROM respuesta a
                                USING (
                                    SELECT 
                                        MIN(ctid) AS ctid_keep,
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    FROM respuesta
                                    GROUP BY 
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    HAVING COUNT(*) > 1
                                ) b
                                WHERE
                                    a.id_pre = b.id_pre
                                    AND a.lat_cto = b.lat_cto
                                    AND a.long_cto = b.long_cto
                                    AND a.op1distanciacto = b.op1distanciacto
                                    AND a.ctid <> b.ctid_keep; """
    
    try:
        cur_pg.execute(limpieza_duplicados)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        messagebox.showinfo("Advertencia", f"Proceso Terminado puntos Viabilizados.\nHora: {hora2}")

        
    except Exception as e:
        print("Error en consulta de zona abierta",e)
def consulta_coordenada_nodos_tigo():
    global pg_params
    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
       
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL nodos")

    
    consulta_nod_tigo = """ INSERT INTO respuesta (
                    id_pre,
                    anio,
                    mes,
                    estadovb,
                    tipo,
                    nit,
                    ds,
                    cliente,
                    proyecto,
                    departamento,
                    ciudad,
                    direccion,
                    complemento,
                    nombre_edifi,
                    zona,
                    latitud,
                    longitud,
                    tipo_servicio,
                    pb,
                    ancho_banda,
                    op1_coinversor,
                    op1_central,
                    op1olt,
                    op1feeder,
                    id_cto1,
                    op1cto,
                    op1distanciacto,
                    op1dipscto,
                    op2_coinversor,
                    op2_central,
                    op2feeder,
                    id_cto2,
                    op2cto,
                    op2distanciacto,
                    oc2dipscto,
                    fecha_solicitud,
                    lat_cto,
                    long_cto,
                    grupo_cto
                )
                SELECT 
                    p.id_pre AS id_proyecto,
                    EXTRACT(YEAR FROM CURRENT_DATE),
                    EXTRACT(MONTH FROM CURRENT_DATE),
                    NULL AS estadovb,
                    NULL AS tipo,
                    NULL AS nit,
                    p.ds,
                    p.nombre_cli,
                    P.proyecto,
                    P.departamento,
                    p.ciudad,
                    p.direccion,
                    p.complemento,
                    p.nombre_edifi,
                    null as zona,
                    p.latitud AS lat_proyecto,
                    p.longitud AS lon_proyecto,
                    p.producto as tipo_servicio,
                    p.tipo_enlace as pb,
                    p.ancho_banda,
                    'TEF' AS op1_coinversor,
                    '-' AS op1_central,
                    '-' AS op1olt,
                    '-' AS op1feeder,
                    o.cod_mov AS id_cto1,
                    o.hl AS cto_osp,
                    ST_Distance(p.geom, o.geom) AS distancia_metros,
                    '-' as puertos_libres,
                    '-' AS op2_coinversor,
                    '' AS op2_central,
                    '' AS op2feeder,
                    '' AS id_cto2,
                    '' AS op2cto,
                    '' AS op2distanciacto,
                    '' AS oc2dipscto,
                    CURRENT_DATE AS fecha_solicitud,
                    o.lat AS lat_osp,
                    o.lon AS lon_osp,
                    'NODO_TIGO'
                FROM proyecto p
                JOIN LATERAL (
                    SELECT *
                    FROM (
                        SELECT 
                            o.*,
                            ROW_NUMBER() OVER (
                                PARTITION BY o.nombre_equipo 
                                ORDER BY p.geom <-> o.geom
                            ) as rn
                        FROM nodo_tigo o
                        WHERE 
                            ST_DWithin(p.geom::geography, o.geom::geography, 50000)
                            
                    ) t
                    WHERE rn = 1   -- solo la más cercana por cada CTO
                    ORDER BY p.geom <-> t.geom
                    LIMIT 2       -- ahora sí: 2 NODOS diferentes
                ) o ON TRUE
                """
    
    
    
    try:
        cur_pg.execute(consulta_nod_tigo)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Proceso de nodos TIGO ejecutado ")

        
    except Exception as e:
        print("Error en consulta de nodos TIGO",e)

    limpieza_duplicados = """ DELETE FROM respuesta a
                                USING (
                                    SELECT 
                                        MIN(ctid) AS ctid_keep,
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    FROM respuesta
                                    GROUP BY 
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    HAVING COUNT(*) > 1
                                ) b
                                WHERE
                                    a.id_pre = b.id_pre
                                    AND a.lat_cto = b.lat_cto
                                    AND a.long_cto = b.long_cto
                                    AND a.op1distanciacto = b.op1distanciacto
                                    AND a.ctid <> b.ctid_keep; """
    
    try:
        cur_pg.execute(limpieza_duplicados)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        messagebox.showinfo("Advertencia", f"Proceso Terminado puntos Viabilizados.\nHora: {hora2}")

        
    except Exception as e:
        print("Error en consulta de zona abierta",e)
def consulta_coordenada_nodos_azteca():
    global pg_params
    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
       
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL nodos")

    
    consulta_nod_azteca = """ INSERT INTO respuesta (
                    id_pre,
                    anio,
                    mes,
                    estadovb,
                    tipo,
                    nit,
                    ds,
                    cliente,
                    proyecto,
                    departamento,
                    ciudad,
                    direccion,
                    complemento,
                    nombre_edifi,
                    zona,
                    latitud,
                    longitud,
                    tipo_servicio,
                    pb,
                    ancho_banda,
                    op1_coinversor,
                    op1_central,
                    op1olt,
                    op1feeder,
                    id_cto1,
                    op1cto,
                    op1distanciacto,
                    op1dipscto,
                    op2_coinversor,
                    op2_central,
                    op2feeder,
                    id_cto2,
                    op2cto,
                    op2distanciacto,
                    oc2dipscto,
                    fecha_solicitud,
                    lat_cto,
                    long_cto,
                    grupo_cto
                )
                SELECT 
                    p.id_pre AS id_proyecto,
                    EXTRACT(YEAR FROM CURRENT_DATE),
                    EXTRACT(MONTH FROM CURRENT_DATE),
                    NULL AS estadovb,
                    NULL AS tipo,
                    NULL AS nit,
                    p.ds,
                    p.nombre_cli,
                    P.proyecto,
                    P.departamento,
                    p.ciudad,
                    p.direccion,
                    p.complemento,
                    p.nombre_edifi,
                    null as zona,
                    p.latitud AS lat_proyecto,
                    p.longitud AS lon_proyecto,
                    p.producto as tipo_servicio,
                    p.tipo_enlace as pb,
                    p.ancho_banda,
                    'TEF' AS op1_coinversor,
                    '-' AS op1_central,
                    '-' AS op1olt,
                    '-' AS op1feeder,
                    o.cod_mov AS id_cto1,
                    o.hl AS cto_osp,
                    ST_Distance(p.geom, o.geom) AS distancia_metros,
                    '-' as puertos_libres,
                    '-' AS op2_coinversor,
                    '' AS op2_central,
                    '' AS op2feeder,
                    '' AS id_cto2,
                    '' AS op2cto,
                    '' AS op2distanciacto,
                    '' AS oc2dipscto,
                    CURRENT_DATE AS fecha_solicitud,
                    o.lat AS lat_osp,
                    o.lon AS lon_osp,
                    'NODO_AZTECA'
                FROM proyecto p
                JOIN LATERAL (
                    SELECT *
                    FROM (
                        SELECT 
                            o.*,
                            ROW_NUMBER() OVER (
                                PARTITION BY o.nombre_equipo 
                                ORDER BY p.geom <-> o.geom
                            ) as rn
                        FROM nodo_azteca o
                        WHERE 
                            ST_DWithin(p.geom::geography, o.geom::geography, 50000)
                            
                    ) t
                    WHERE rn = 1   -- solo la más cercana por cada CTO
                    ORDER BY p.geom <-> t.geom
                    LIMIT 2       -- ahora sí: 2 NODOS diferentes
                ) o ON TRUE
                """
    
    
    
    try:
        cur_pg.execute(consulta_nod_azteca)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        print("Proceso de nodos azteca ejecutado ")

        
    except Exception as e:
        print("Error en nodos azteca de emp",e)

    limpieza_duplicados = """ DELETE FROM respuesta a
                                USING (
                                    SELECT 
                                        MIN(ctid) AS ctid_keep,
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    FROM respuesta
                                    GROUP BY 
                                        id_pre,
                                        lat_cto,
                                        long_cto,
                                        op1distanciacto
                                    HAVING COUNT(*) > 1
                                ) b
                                WHERE
                                    a.id_pre = b.id_pre
                                    AND a.lat_cto = b.lat_cto
                                    AND a.long_cto = b.long_cto
                                    AND a.op1distanciacto = b.op1distanciacto
                                    AND a.ctid <> b.ctid_keep; """
    
    try:
        cur_pg.execute(limpieza_duplicados)
        con_pg.commit()
        hora2 = datetime.now().strftime("%H:%M")
        messagebox.showinfo("Advertencia", f"Proceso Terminado puntos Viabilizados.\nHora: {hora2}")

        
    except Exception as e:
        print("Error en consulta de zona abierta",e)




def proceso_robot_consulta():
    global pg_params
    hora = datetime.now().strftime("%H:%M")

    # Abrir cuadro de diálogo para seleccionar archivo
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo",
        filetypes=(("Archivos de texto", "viabilidad.csv"), ("Todos los archivos", "*.csv"))
    )

    print(hora,"Archivo seleccionado:", archivo)


    # Configuración PostgreSQL
    pg_params = {
        "host": "192.168.137.1",
        "database": "ROBOT_VB",
       "user": "postgres",
        "password": "1030571984"
    }

   

    # Configuración PostgreSQL
    
    hora = datetime.now().strftime("%H:%M")
    print("Conectando a postgress",hora)
    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
        print("Conexion exitosa a base PosgreSQL")
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL robot consulta")


    limpiezatablaproyecto = "truncate table proyecto"
    try:
        
        cur_pg.execute(limpiezatablaproyecto)
        con_pg.commit()
        hora = datetime.now().strftime("%H:%M")
        print("Limpieza de tabla proyecto Exitosa",hora)
    except Exception as e:
        print("Error en limpieza de tabla proyecto Base PosgreSQL",e)
    
    limpieza_tb_respuesta = """TRUNCATE TABLE respuesta """
    try:
        
        cur_pg.execute(limpieza_tb_respuesta)
        con_pg.commit()
        hora = datetime.now().strftime("%H:%M")
        print("Limpieza de tabla proyecto Exitosa",hora)
    except Exception as e:
        print("Error en limpieza de tabla proyecto Base PosgreSQL",e)


    datos_proyecto = []

    with open(archivo, newline='', encoding='utf-8')  as archivoproyec:
                lector = csv.reader(archivoproyec, delimiter=";")
                next(lector)
                for fila in lector:
                    if len(fila)< 18:
                        print("Fila con datos incompletos, se saltara la fila:", fila)
                        continue

                    idpre = fila[0]
                    ds    = fila[1]
                    Cliente = fila[2]
                    proyecto = fila[3]
                    departamento = fila[4]
                    ciudad = fila[5]
                    direccion = fila[6]
                    complemento1 = fila[7]
                    nom_edificio = fila[8]
                    barrio = fila[9]
                    latitud = fila[10].replace(",", ".")
                    longitud = fila[11].replace(",", ".")
                    coordenadas = fila[12]
                    producto = fila[13]
                    ancho_banda = fila[14]
                    enlace = fila[15]
                    fech_registro = fila[16].strip()
                    usuario = fila[17]


                    datos_proyecto.append((
                        idpre,
                        ds,
                        Cliente,
                        proyecto,
                        departamento,
                        ciudad,
                        direccion,
                        complemento1,
                        nom_edificio,
                        barrio,
                        latitud,
                        longitud,
                        coordenadas,
                        producto,
                        ancho_banda,
                        enlace,
                        usuario
                    ))

            # insertar a mysql

    try:

        insert_datos_proyecto = """ INSERT into proyecto(id_pre,
                                                        ds,
                                                        nombre_cli,
                                                        proyecto,
                                                        departamento,
                                                        ciudad,
                                                        direccion,
                                                        complemento,
                                                        nombre_edifi,
                                                        barrio,
                                                        latitud,
                                                        longitud,
                                                        coordenadas,
                                                        producto,
                                                        ancho_banda,
                                                        tipo_enlace,
                                                        fech_regi,
                                                        usuario_reg,
                                                        respuesta,
                                                        geom) 
                                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,CURRENT_DATE,%s,null,null); """

        cur_pg.executemany(insert_datos_proyecto, datos_proyecto)
        con_pg.commit()
        hora = datetime.now().strftime("%H:%M")
        print("Datos tabla ordenes de trabajo insertados correctamente", hora)
    
    except Exception as e:
        print("Error en ejecucion ordenes de trabajo ")


    datogeom_proyec = """ UPDATE proyecto
                        SET geom = ST_SetSRID(ST_MakePoint(longitud, latitud), 4326); """


    try:
        
        cur_pg.execute(datogeom_proyec)
        con_pg.commit()
        hora = datetime.now().strftime("%H:%M")
        print("Dato para cruce de coordenadas creado",hora)
    except Exception as e:
        print("Error en creacion de geom para coordenadas",hora)


    consulta_direccion_con_complemento = """ INSERT INTO respuesta (
                                                    id_pre,
                                                    anio,
                                                    mes,
                                                    estadovb,
                                                    tipo,
                                                    nit,
                                                    ds,
                                                    cliente,
                                                    proyecto,
                                                    departamento,
                                                    ciudad,
                                                    direccion,
                                                    complemento,
                                                    nombre_edifi,
                                                    zona,
                                                    latitud,
                                                    longitud,
                                                    tipo_servicio,
                                                    pb,
                                                    ancho_banda,
                                                    op1_coinversor,
                                                    op1_central,
                                                    op1olt,
                                                    op1feeder,
                                                    id_cto1,
                                                    op1cto,
                                                    op1distanciacto,
                                                    op1dipscto,
                                                    op2_coinversor,
                                                    op2_central,
                                                    op2feeder,
                                                    id_cto2,
                                                    op2cto,
                                                    op2distanciacto,
                                                    oc2dipscto,
                                                    fecha_solicitud,
                                                    lat_cto,
                                                    long_cto,
                                                    grupo_cto
                                                )
                                                SELECT DISTINCT ON (c.direccion)
                                                    c.id_pre,
                                                    EXTRACT(YEAR FROM CURRENT_DATE),
                                                    EXTRACT(MONTH FROM CURRENT_DATE),
                                                    NULL AS estadovb,
                                                    NULL AS tipo,
                                                    NULL AS nit,
                                                    c.ds,
                                                    c.nombre_cli AS cliente,
                                                    c.proyecto,
                                                    NULL AS departamento,
                                                    c.ciudad,
                                                    c.direccion,
                                                    c.complemento,
                                                    c.nombre_edifi,
                                                    NULL AS zona,
                                                    c.latitud,
                                                    c.longitud,
                                                    c.producto AS tipo_servicio,
                                                    c.tipo_enlace AS pb,
                                                    c.ancho_banda,
                                                    b.proveedor AS op1_coinversor,
                                                    NULL AS op1_central,
                                                    b.olt_no AS op1olt,
                                                    CONCAT('CA', b.cable) AS op1feeder,
                                                    b.idplaca AS id_cto1,
                                                    b.cto AS op1cto,
                                                    '0' AS op1distanciacto,
                                                    b.puertos_libres_cto AS op1dipscto,
                                                    '' AS op2_coinversor,
                                                    '' AS op2_central,
                                                    '' AS op2feeder,
                                                    '' AS id_cto2,
                                                    '' AS op2cto,
                                                    '' AS op2distanciacto,
                                                    '' AS oc2dipscto,
                                                    CURRENT_DATE AS fecha_solicitud,
                                                    0 as lat_cto,
                                                    0 as long_cto,
                                                    '' as grupo_cto
                                                FROM osp_3gys AS b
                                                JOIN proyecto AS c
                                                    ON b.localidad = c.ciudad
                                                WHERE REPLACE(b.direccion_cli,' ','') = REPLACE(concat(c.direccion,c.complemento),' ','')
                                                ORDER BY c.direccion;"""
    
    try:
        cur_pg.execute(consulta_direccion_con_complemento)
        con_pg.commit()
        print("Consulta por direccion con complemento ejecutada")
    except Exception as e:
        print("Error en consulta por direccion con complemento",e)

    marca_consulta_por_direccion = """ UPDATE proyecto AS B
                                        SET respuesta = 'DIRECCION EXACTA'
                                    FROM respuesta AS A
                                        WHERE A.id_pre = B.id_pre; """
    try:
        cur_pg.execute(marca_consulta_por_direccion)
        con_pg.commit()
        print("Punto por direccion marcado")
    except Exception as e:
        print("Error en consulta por direccion exacta.",e)
    consulta_direccion_SIN_complemento = """ INSERT INTO respuesta (
                                                    id_pre,
                                                    anio,
                                                    mes,
                                                    estadovb,
                                                    tipo,
                                                    nit,
                                                    ds,
                                                    cliente,
                                                    proyecto,
                                                    departamento,
                                                    ciudad,
                                                    direccion,
                                                    complemento,
                                                    nombre_edifi,
                                                    zona,
                                                    latitud,
                                                    longitud,
                                                    tipo_servicio,
                                                    pb,
                                                    ancho_banda,
                                                    op1_coinversor,
                                                    op1_central,
                                                    op1olt,
                                                    op1feeder,
                                                    id_cto1,
                                                    op1cto,
                                                    op1distanciacto,
                                                    op1dipscto,
                                                    op2_coinversor,
                                                    op2_central,
                                                    op2feeder,
                                                    id_cto2,
                                                    op2cto,
                                                    op2distanciacto,
                                                    oc2dipscto,
                                                    fecha_solicitud,
                                                    lat_cto,
                                                    long_cto,
                                                    grupo_cto
                                                )
                                                SELECT DISTINCT ON (c.direccion)
                                                    c.id_pre,
                                                    EXTRACT(YEAR FROM CURRENT_DATE),
                                                    EXTRACT(MONTH FROM CURRENT_DATE),
                                                    NULL AS estadovb,
                                                    NULL AS tipo,
                                                    NULL AS nit,
                                                    c.ds,
                                                    c.nombre_cli AS cliente,
                                                    c.proyecto,
                                                    NULL AS departamento,
                                                    c.ciudad,
                                                    c.direccion,
                                                    c.complemento,
                                                    c.nombre_edifi,
                                                    NULL AS zona,
                                                    c.latitud,
                                                    c.longitud,
                                                    c.producto AS tipo_servicio,
                                                    c.tipo_enlace AS pb,
                                                    c.ancho_banda,
                                                    b.proveedor AS op1_coinversor,
                                                    NULL AS op1_central,
                                                    b.olt_no AS op1olt,
                                                    CONCAT('CA', b.cable) AS op1feeder,
                                                    b.idplaca AS id_cto1,
                                                    b.cto AS op1cto,
                                                    '0' AS op1distanciacto,
                                                    b.puertos_libres_cto AS op1dipscto,
                                                    '' AS op2_coinversor,
                                                    '' AS op2_central,
                                                    '' AS op2feeder,
                                                    '' AS id_cto2,
                                                    '' AS op2cto,
                                                    '' AS op2distanciacto,
                                                    '' AS oc2dipscto,
                                                    CURRENT_DATE AS fecha_solicitud,
                                                    0 as lat_cto,
                                                    0 as long_cto,
                                                    '' as grupo_cto
                                                FROM osp_3gys AS b
                                                JOIN proyecto AS c
                                                    ON b.localidad = c.ciudad
                                                WHERE REPLACE(b.direccion_sin_complemento,' ','') = REPLACE(c.direccion,' ','')
                                                AND and C.respuesta is null
                                                ORDER BY c.direccion;"""
    
    try:
        cur_pg.execute(consulta_direccion_SIN_complemento)
        con_pg.commit()
        print("Consulta por direccion SIN complemento ejecutada")
    except Exception as e:
        print("Error en consulta por direccion SIN complemento",e)

    marca_consulta_por_direccion = """ UPDATE proyecto AS B
                                        SET respuesta = 'DIRECCION EXACTA'
                                    FROM respuesta AS A
                                        WHERE A.id_pre = B.id_pre 
                                        AND B.respuesta is null ; """
    try:
        cur_pg.execute(marca_consulta_por_direccion)
        con_pg.commit()
        print("Punto por direccion marcado")
    except Exception as e:
        print("Error en consulta por direccion exacta.",e)

    consulta_coordenada_zc()
    consulta_coordenada_za()
    consulta_coordenada_emp()
    #consulta_coordenada_nodos()
    #consulta_coordenada_nodos_tigo()
    #consulta_coordenada_nodos_azteca()
    archiv()





def consulta_por_punto():
    global latitud
    global longitud
    global ciudad
    global tabla_respuesta
    global pg_params

    #pg_params = {
    #    "host": "192.168.137.1",
    #   "database": "ROBOT_VB",
    #    "user": "postgres",
    #    "password": "1030571984"
    #}

    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
        print("Conexion exitosa a base PosgreSQL")
    except Exception as e:
        print("Error en conexion a base de datos PosgreSQL consulta por punto")
        return


    
    limpieza_punto_sencillo = """ TRUNCATE table punto_sencillo; """

    try:
        cur_pg.execute(limpieza_punto_sencillo)
        con_pg.commit()
        print("Limpieza de tabla punto sencillo exitosa")
    except Exception as e:
        print("Error en la limpieza de punto sencillo",e)

    lat = latitud.get() or 0
    long = longitud.get() or 0
    ciud = ciudad.get()
    direc = direccion.get()

    dato_punto_sencillo = """ INSERT into punto_sencillo(id_pre,
                                                    ds,
                                                    nombre_cli,
                                                    proyecto,
                                                    departamento,
                                                    ciudad,
                                                    direccion,
                                                    complemento,
                                                    nombre_edifi,
                                                    barrio,
                                                    latitud,
                                                    longitud,
                                                    coordenadas,
                                                    producto,
                                                    ancho_banda,
                                                    tipo_enlace,
                                                    fech_regi,
                                                    usuario_reg,
                                                    respuesta,
                                                    geom) 
                                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,CURRENT_DATE,%s,null,null); """
    valores = (
        1, '', '', '', '', ciud, direc, '', '', '',
        lat,         # <-- LATITUD AQUÍ
        long,         # <-- LONGITUD AQUÍ
        '', '', '', '', ''
    )

    try:
        cur_pg.execute(dato_punto_sencillo, valores)
        con_pg.commit()
        print("Inserción exitosa en la tabla proyecto")
    except Exception as e:
        print("Error insertando datos", e)

    ajuste_geom = """ UPDATE punto_sencillo
                SET geom = ST_SetSRID(ST_MakePoint(longitud, latitud), 4326); """
    try:
        cur_pg.execute(ajuste_geom)
        con_pg.commit()
        print("ajuste geom en punto sencillo para consulta")
    except Exception as e:
        print("Error en el proceso de construccion de geom para punto sencillo")


    limpiar_tabla()

    consulta_direccion_punto_sencillo = """ SELECT 
                                            c.latitud,
                                            c.longitud,
                                            b.direccion_sin_complemento,
                                            b.proveedor AS op1_coinversor,
                                            b.olt_no AS op1olt,
                                            CONCAT('CA', b.cable) AS op1feeder,
                                            b.idplaca AS id_cto1,
                                            b.cto AS op1cto,
                                            '0' AS op1distanciacto,
                                            b.puertos_libres_cto AS op1dipscto,
                                            0 as lat_cto,
                                            0 as long_cto,
                                            '' as grupo_cto
                                        FROM osp_3gys AS b
                                        JOIN punto_sencillo AS c
                                            ON b.localidad = c.ciudad
                                        WHERE REPLACE(b.direccion_sin_complemento,' ','') like  CONCAT('%', REPLACE(c.direccion, ' ', ''), '%') ORDER BY b.direccion_sin_complemento
                                        """
    try:
        cur_pg.execute(consulta_direccion_punto_sencillo)
        resultados = cur_pg.fetchall()
    except Exception as e:
        print("Error ejecutando consulta:", e)
        return

    for item in tabla_respuesta.get_children():
        tabla_respuesta.delete(item)
    # Validación
    if not resultados:
        print("No se encontraron resultados.")
       
    else:
        # Insertar en tabla
         for resultado in resultados:
            tabla_respuesta.insert(
                "",
                "end",
                values=(
                    resultado[0],   # LAT_CONSULTADA
                    resultado[1],   # LONG_CONSULTADA
                    resultado[2],   # DIRECCION
                    resultado[10],  # LAT_CTO
                    resultado[11],  # LONG_CTO
                    resultado[6],   # ID_CTO
                    resultado[7],   # OPC_CTO
                    resultado[8],   # DISTANCIA_CTO
                    resultado[5],   # OP_FEEDER
                    resultado[9],   # PUERTOS_DISPONIBLES
                    resultado[4],   # OLT
                    resultado[12]   # CONSULTA
                )
            )

    consulta_zc = """ SELECT 
                        p.latitud,
						p.longitud,
						p.direccion,
						o.proveedor AS op1_coinversor,
						o.olt_no AS op1olt,
						CONCAT('CA', o.cable) AS op1feeder,
						o.idplaca AS id_cto1,
						o.cto AS op1cto,
						ST_Distance(p.geom, o.geom)::numeric(10,2) AS op1distanciacto,
						o.puertos_libres_cto AS op1dipscto,
						o.lat_equipo as lat_cto,
						o.long_equipo as long_cto,
						o.grupo_cto
                    FROM punto_sencillo p
                    JOIN LATERAL (
                        SELECT DISTINCT ON (geom) *
                        FROM (
                            SELECT o.*
                            FROM osp_3gys o
                            WHERE o.grupo_cto  <> 'Zona Abierta'
                            and o.cto like '%PC%'
                            and o.puertos_libres_cto > 1
                            and ST_DWithin(p.geom, o.geom, 50)
                            and p.respuesta is null
                            ORDER BY p.geom <-> o.geom       -- orden por distancia entre coordenadas
                            LIMIT 50                         -- limite de datos analizar
                        ) x
                        ORDER BY geom, p.geom <-> x.geom     -- ahora eliminamos duplicados manteniendo los más cercanos
                        
                    ) o ON TRUE; """

    try:
        cur_pg.execute(consulta_zc)
        resultado = cur_pg.fetchone()
    except Exception as e:
        print("Error ejecutando consulta:", e)
        return

    # Validación
    if not resultado:
        print("No se encontraron resultados ZC.")
    else:       
        # Insertar en tabla
        tabla_respuesta.insert(
            "",
            "end",
            values=(
                resultado[0],   # LAT_CONSULTADA
                resultado[1],   # LONG_CONSULTADA
                resultado[2],   # DIRECCION
                resultado[10],  # LAT_CTO
                resultado[11],  # LONG_CTO
                resultado[6],   # ID_CTO
                resultado[7],   # OPC_CTO
                resultado[8],   # DISTANCIA_CTO
                resultado[5],   # OP_FEEDER
                resultado[9],   # PUERTOS_DISPONIBLES
                resultado[4],   # OLT
                resultado[12]    # CONSULTA
            )
        )

    consulta_za = """ SELECT 
                        p.latitud,
						p.longitud,
						p.direccion,
						o.proveedor AS op1_coinversor,
						o.olt_no AS op1olt,
						CONCAT('CA', o.cable) AS op1feeder,
						o.idplaca AS id_cto1,
						o.cto AS op1cto,
						ST_Distance(p.geom, o.geom)::numeric(10,2) AS op1distanciacto,
						o.puertos_libres_cto AS op1dipscto,
						o.lat_equipo as lat_cto,
						o.long_equipo as long_cto,
						o.grupo_cto
                    FROM punto_sencillo p
                    JOIN LATERAL (
                        SELECT DISTINCT ON (geom) *
                        FROM (
                            SELECT o.*
                            FROM osp_3gys o
                            WHERE o.grupo_cto  like '%Abierta%'
                            and O.puertos_libres_cto > 1
                            and ST_DWithin(p.geom, o.geom, 500)
                            and p.respuesta is null
                            ORDER BY p.geom <-> o.geom       -- orden por distancia entre coordenadas
                            LIMIT 50                         -- limite de datos analizar
                        ) x
                        ORDER BY geom, p.geom <-> x.geom     -- ahora eliminamos duplicados manteniendo los más cercanos
                        LIMIT 2
                    ) o ON TRUE; """

    try:
        cur_pg.execute(consulta_za)
        resultado = cur_pg.fetchone()
    except Exception as e:
        print("Error ejecutando consulta:", e)
        return

    # Validación
    if not resultado:
        print("No se encontraron resultados ZA.")
      
    else:
        # Insertar en tabla
        tabla_respuesta.insert(
            "",
            "end",
            values=(
                resultado[0],   # LAT_CONSULTADA
                resultado[1],   # LONG_CONSULTADA
                resultado[2],   # DIRECCION
                resultado[10],  # LAT_CTO
                resultado[11],  # LONG_CTO
                resultado[6],   # ID_CTO
                resultado[7],   # OPC_CTO
                resultado[8],   # DISTANCIA_CTO
                resultado[5],   # OP_FEEDER
                resultado[9],   # PUERTOS_DISPONIBLES
                resultado[4],   # OLT
                resultado[12]    # CONSULTA
            )
        )

def validar_latitud(texto):
    global latitud
    # Permitir vacío (mientras escriben)
    if texto == "":
        return True

    # Solo permitir números y un punto
    if not all(c in "0123456789." for c in texto):
        return False

    # Evitar más de un punto
    if texto.count(".") > 1:
        return False

    # Separar antes y después del punto
    partes = texto.split(".")
    antes = partes[0]

    # No permitir más de 2 dígitos antes del punto
    if len(antes) > 2:
        return False

    # Si tiene decimales, no permitir más de 8
    if len(partes) == 2 and len(partes[1]) > 8:
        return False

    # Todo correcto
    return True
def validar_longitud(texto):
    # Permitir vacío mientras escriben
    if texto == "":
        return True

    # Debe empezar siempre con "-"
    if texto[0] != "-":
        return False

    # Quitar el signo para validar solo números y punto
    contenido = texto[1:]

    # Solo permitir números y un punto
    if not all(c in "0123456789." for c in contenido):
        return False

    # Máximo un punto
    if contenido.count(".") > 1:
        return False

    partes = contenido.split(".")
    antes = partes[0]

    # Máximo 3 dígitos antes del punto
    if len(antes) > 3:
        return False

    # Máximo 8 dígitos después del punto
    if len(partes) == 2 and len(partes[1]) > 8:
        return False

    return True


def formatear_latitud(event):
    global latitud
    valor = latitud.get()

    # Quitar punto para volver a contar dígitos reales
    solo_digitos = valor.replace(".", "")

    # Si hay más de 2 dígitos y NO tiene punto → insertar punto automático
    if len(solo_digitos) > 2 and "." not in valor:
        nuevo = solo_digitos[:2] + "." + solo_digitos[2:]
        latitud.delete(0, tk.END)
        latitud.insert(0, nuevo)
def formatear_longitud(event):
    widget = event.widget
    valor = widget.get()

    # Si está vacío no hacer nada
    if valor == "":
        return

    # Si no empieza con "-", agregarlo
    if valor[0] != "-":
        valor = "-" + valor
        widget.delete(0, tk.END)
        widget.insert(0, valor)

    # Quitar el signo para trabajar
    numero = valor[1:]

    # Eliminar puntos temporales
    solo_digitos = numero.replace(".", "")

    # Insertar punto automáticamente después del tercer dígito
    if len(solo_digitos) > 2 and "." not in numero:
        nuevo = "-" + solo_digitos[:2] + "." + solo_digitos[2:]
        widget.delete(0, tk.END)
        widget.insert(0, nuevo)

def archiv():
    global pg_params

    try:
        con_pg = psycopg2.connect(**pg_params)
        cur_pg = con_pg.cursor()
        print("Conexión exitosa a PostgreSQL")
    except Exception as e:
        print("Error en conexión a la base de datos PosgreSQL:", e)
        messagebox.showinfo("Advertencia", "Error en la conexión a la base de datos.")
        return

    consulta_creada = """
            SELECT id_pre AS ID_PRE,
                ds AS DS,
                cliente AS NOMBRE_CLIENTE,
                ciudad AS CIUDAD,
                direccion AS DIRECCION,
                latitud AS LATITUD,
                longitud AS LONGITUD,
                ROW_NUMBER() OVER (PARTITION BY id_pre ORDER BY id_pre, grupo_cto DESC, op1distanciacto ) AS ORDEN,
                id_cto1 AS ID_CTO,
                op1_coinversor AS PROPIETARIO,
                '' as CENTRAL,
                op1olt as OLT_NOMBRE,
                op1feeder as CABLE,
                op1cto as CTO,
                lat_cto AS LAT_CTO,
                long_cto AS LONG_CTO,
                ROUND((op1distanciacto::numeric) +((op1distanciacto::numeric) * 50 / 100)) AS DISTANCIA_CTO,
                op1dipscto as CTO_DISPO,
                '' as OT,
                'OSP_3GYS' as TABLA,
                CASE 
                    WHEN op1distanciacto  = '0' 
                        THEN 'DIR EXACTA' 
                        ELSE 'CTO CERCANA' 
                    END AS RESPUESTA,
                CASE WHEN op1cto  LIKE '%EMP%'
                THEN 'EMP'
                ELSE
                CASE 
                    WHEN grupo_cto = 'ZA' THEN 'ZA'
                    WHEN grupo_cto = 'nodo_TEF' THEN 'NODO_TEF'
                    WHEN grupo_cto = 'NODO_AZTECA' THEN 'NODO_AZTECA'
                    WHEN grupo_cto = 'NODO_TIGO' THEN 'NODO_TIGO'
                         
                        ELSE 'ZC' 
                    END END AS TIPO_ZONA
            FROM respuesta 
            WHERE 
			(grupo_cto <> 'EMP' AND ROUND((op1distanciacto::numeric) +((op1distanciacto::numeric) * 50 / 100)) < 501)
			OR 
			(grupo_cto = 'EMP' AND ROUND((op1distanciacto::numeric) +((op1distanciacto::numeric) * 50 / 100)) < 1001)
            ORDER BY id_pre, orden, tipo_zona, OP1DISTANCIACTO DESC;
    """

    try:
        cur_pg.execute(consulta_creada)
        rows = cur_pg.fetchall()
        cols = [desc[0] for desc in cur_pg.description]

        
        # Crear archivo Excel (.xlsx)
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"

        # Escribir encabezados
        ws.append(cols)

        # Escribir filas
        for row in rows:
            ws.append(row)

        # Crear nombre del archivo dinámico
        dia = datetime.now().day
        meses = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"]
        mes = meses[datetime.now().month - 1]
        cantidad = len(rows)

        nombre_archivo = f"ENTREGAS {dia}-{mes}-{cantidad} VB_FTTH_E.xlsx"
        ruta_archivo = fr"C:\RESPUESTAS_ROBOT\{nombre_archivo}"

        wb.save(ruta_archivo)

        con_pg.commit()
        cur_pg.close()
        con_pg.close()

        messagebox.showinfo("Éxito", f"Archivo generado correctamente:\n{ruta_archivo}")

    except Exception as e:
        print("Error:", e)
        messagebox.showinfo("Advertencia", "Error al ejecutar la consulta o crear el archivo.")

    





    

ventana_inicial()