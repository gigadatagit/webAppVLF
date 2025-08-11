import streamlit as st
import json
import io
import os
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from staticmap import StaticMap, CircleMarker

def obtener_template_path(tipo_tramo: str, cantidad_tramos: int) -> str:
    """
    Retorna el path del template a usar basado en el tipo de tramo y la cantidad.
    Ejemplo: 'Trifásicos' y 3 → 'templateVLF3FS3TR.docx'
             'Monofásicos' y 10 → 'templateVLF1FS10TR.docx'
    """
    fases = "3FS" if tipo_tramo == "Trifásicos" else "1FS"
    nombre_template = f"templateVLF{fases}{cantidad_tramos}TR.docx"
    return os.path.join('templates', nombre_template)

def convertir_a_mayusculas(data):
    if isinstance(data, str):
        return data.upper()
    elif isinstance(data, dict):
        return {k: convertir_a_mayusculas(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [convertir_a_mayusculas(v) for v in data]
    elif isinstance(data, tuple):
        return tuple(convertir_a_mayusculas(v) for v in data)
    else:
        return data  # cualquier otro tipo se deja igual

# Configuración de plantilla de Word (una sola instancia)
#template_path = os.path.join('templates', 'templateVLF3FS3TR.docx')

# Diccionario de labels para preguntas de verificación
preguntas_verificacion = {
    'frmVerfCabPreg1': 'El rótulo del cable en su chaqueta es legible y congruente con lo instalado en sitio',
    'frmVerfCabPreg2': 'Limpieza de cada una de las terminales',
    'frmVerfCabPreg3': 'Marcación correcta de los cables en ambos extremos',
    'frmVerfCabPreg4': 'Verificación de continuidad del cable de acuerdo a las marcaciones',
    'frmVerfCabPreg5': 'Verificación del tendido y conexionado del cable XLPE',
    'frmVerfCabPreg6': 'Distancias de seguridad entre cables apropiadas para hacer la prueba VLF'
}

#if 'doc' not in st.session_state:
#    st.session_state.doc = DocxTemplate(template_path)

# Inicialización de estado
if 'step' not in st.session_state:
    st.session_state.step = 1
    st.session_state.data = {}

st.title("Formulario VLF - Word Automatizado")

# Funciones de navegación
def next_step():
    missing = [k for k, v in st.session_state.data.items() if v is None or v == ""]
    if missing:
        st.error("Por favor completa todos los campos antes de continuar.")
    else:
        st.session_state.step += 1
        st.rerun()

def prev_step():
    if st.session_state.step > 1:
        st.session_state.step -= 1
        st.rerun()

# Paso 1: Información General
if st.session_state.step == 1:
    st.header("Paso 1: Información General")
    st.session_state.data['nombreProyecto'] = st.text_input("Nombre del Proyecto", key='nombreProyecto')
    st.session_state.data['nombreCiudadoMunicipio'] = st.text_input("Ciudad o Municipio", key='ciudad')
    st.session_state.data['nombreDepartamento'] = st.text_input("Departamento", key='departamento')
    st.session_state.data['tipoCoordenada'] = st.selectbox(f"Tipo de Imagen para las Coordenadas", ["Urbano", "Rural"], key=f'tipo_coordenada')
    st.session_state.data['nombreCompleto'] = st.text_input("Nombre Completo", key='nombre')
    st.session_state.data['nroConteoTarjeta'] = st.text_input("Número de CONTE o Tarjeta Profesional", key='conte_tarjeta')
    st.session_state.data['nombreCargo'] = st.text_input("Nombre del Cargo", key='cargo')
    st.session_state.data['fechaCreacionSinFormato'] = st.date_input("Fecha de Creación", key='fecha_creacion', value=datetime.now())
    st.session_state.data['fechaCreacion'] = st.session_state.data['fechaCreacionSinFormato'].strftime("%Y-%m-%d")
    st.session_state.data['direccion'] = st.text_input("Dirección", key='direccion')
    #Agregar el campo de selección de Rural o Urbano para la generación de la imagen 

    cols = st.columns([1,1])
    if cols[1].button("Siguiente"):
        next_step()

# Paso 2: Datos Técnicos
elif st.session_state.step == 2:
    st.header("Paso 2: Datos Técnicos")
    st.session_state.data['tensionPrueba'] = st.selectbox("Tensión de Prueba", ["Aceptación", "Mantenimiento"], key='tension')
    st.session_state.data['valTensionPrueba'] = 21 if st.session_state.data['tensionPrueba'] == "Aceptación" else 16
    tipo = st.selectbox("Tipo de Tramos", ["Trifásicos", "Monofásicos"], key='tipo_tramos')
    st.session_state.data['tipoTramos'] = tipo
    max_tramos = 10 if tipo == "Trifásicos" else 20
    st.session_state.data['cantidadTramos'] = st.number_input("Cantidad de Tramos", min_value=1, max_value=max_tramos, step=1, key='cantidad_tramos')
    st.session_state.data['latitud'] = st.number_input("Latitud", key='latitud', format="%.6f")
    st.session_state.data['longitud'] = st.number_input("Longitud", key='longitud', format="%.6f")
    st.session_state.data['caracteristicasCable'] = st.text_input("Características del Cable", key='caracteristicas')
    st.session_state.data['fechaCalibracionSinFormato'] = st.date_input("Fecha de Calibración", key='fecha_calibracion')
    st.session_state.data['fechaCalibracion'] = st.session_state.data['fechaCalibracionSinFormato'].strftime("%Y-%m-%d")

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        #next_step()
        # Crear ruta del template dinámicamente
        tipo_tramos = st.session_state.data.get('tipoTramos')
        cantidad_tramos = st.session_state.data.get('cantidadTramos')
        template_path = obtener_template_path(tipo_tramos, cantidad_tramos)
        
        # Cargar la plantilla en el estado de sesión
        try:
            st.session_state.doc = DocxTemplate(template_path)
            next_step()
        except FileNotFoundError:
            st.error(f"No se encontró la plantilla: {template_path}")

# Paso 3: Formulario de Verificación
elif st.session_state.step == 3:
    st.header("Paso 3: Formulario de Verificación del Cable")
    opciones = ["Sí ", "No "]
    for key, label in preguntas_verificacion.items():
        st.session_state.data[key] = st.selectbox(label, opciones, key=key)
    st.session_state.data['comVerificacion'] = st.text_area("Comentarios de Verificación", key='comentarios_verificacion')

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        next_step()

# Paso 4: Detalles por Tramo
elif st.session_state.step == 4:
    st.header("Paso 4: Detalles por Tramo")
    tipo = st.session_state.data['tipoTramos']
    cantidad = int(st.session_state.data['cantidadTramos'])
    fases = ['A', 'B', 'C'] if tipo == 'Trifásicos' else ['']
    for i in range(1, cantidad + 1):
        for f in fases:
            suf = f"Trm{i}{f or ''}"
            st.subheader(f"Tramo {i} Fase {f or 'Única'}")
            st.session_state.data[f'descripcionTramo_{suf}'] = st.text_input(f"Descripción {suf}", key=f'desc_{suf}')
            st.session_state.data[f'nombreCircuito{suf}'] = st.text_input(f"Nombre del Circuito {suf}", key=f'circuito_{suf}')
            st.session_state.data[f'corrienteTramo{suf}'] = st.number_input(f"Corriente del Tramo {suf} (μArms)", key=f'corr_{suf}', min_value=0.0, format="%.2f")
            st.session_state.data[f'distanciaCable{suf}'] = st.number_input(f"Distancia del Cable {suf} (m)", key=f'dist_{suf}', min_value=0.0, format="%.2f")
            st.session_state.data[f'evaluacionFinal{suf}'] = st.selectbox(f"Evaluación Final {suf}", ["CUMPLE", "NO CUMPLE"], key=f'eval_{suf}')

    cols = st.columns([1,1,1])
    if cols[0].button("Anterior"):
        prev_step()
    if cols[1].button("Siguiente"):
        next_step()

# Paso 5: Subida de Imágenes y Generación de Word
elif st.session_state.step == 5:
    
    st.header("Paso 5: Subida de Imágenes de Pruebas y Mapa")
    datos_Sin_Mayuscula = st.session_state.data.copy()
    datos = convertir_a_mayusculas(datos_Sin_Mayuscula)
    cantidad = int(datos.get('cantidadTramos', 0))
    tipo = datos.get('tipoTramos')
    fases = ['A', 'B', 'C'] if tipo == 'Trifásicos' else ['']

    # Mapa
    if datos.get('latitud') and datos.get('longitud'):
        try:
            lat = float(datos['latitud'])
            lon = float(datos['longitud'])
            mapa = StaticMap(600, 400)
            mapa.add_marker(CircleMarker((lon, lat), 'red', 12))
            img_map = mapa.render()
            buf_map = io.BytesIO()
            img_map.save(buf_map, format='PNG')
            buf_map.seek(0)
            datos['imgMapsProyecto'] = InlineImage(st.session_state.doc, buf_map, Cm(18))
        except:
            st.error("Coordenadas inválidas para el mapa.")
    else:
        st.error("Faltan coordenadas para el mapa.")

    # Imagen de tensión
    tension = datos.get('tensionPrueba')
    img_path = None
    if tension == 'Aceptación':
        img_path = 'images/imgAceptacion.png'
    elif tension == 'Mantenimiento':
        img_path = 'images/imgMantenimiento.png'
    if img_path and os.path.exists(img_path):
        buf_t = io.BytesIO(open(img_path, 'rb').read())
        buf_t.seek(0)
        datos['imgTablaTensionPrueba'] = InlineImage(st.session_state.doc, buf_t, Cm(18))

    # Subida de imágenes por tramo
    st.subheader("Imágenes de Pruebas de Tramos")
    for i in range(1, cantidad + 1):
        for f in fases:
            key = f"imgPruebaTramoTrm{i}{f or ''}"
            uploaded = st.file_uploader(f"Imagen para Tramo {i} Fase {f or 'Única'}", type=['png','jpg','jpeg'], key=key)
            if uploaded:
                buf = io.BytesIO(uploaded.read())
                buf.seek(0)
                datos[key] = InlineImage(st.session_state.doc, buf, Cm(14))
            else:
                datos[key] = None

    if st.button("Generar Word"):
        doc = st.session_state.doc
        # Añadir fecha al contexto
        ahora = datetime.now()
        meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
        datos['dia'] = ahora.day
        datos['mes'] = meses[ahora.month-1]
        datos['anio'] = ahora.year

        doc.render(datos)
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            "Descargar Reporte Word",
            data=output,
            file_name="reporteProtocoloVLF.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
