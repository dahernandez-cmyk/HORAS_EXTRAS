import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Consolidador de Nómina", layout="wide")

st.title("📂 Validador y Consolidador de Archivos de Horas Extras")
st.write("Sube los archivos Excel para validarlos y unirlos en uno solo.")

# Estructura requerida
COLUMNAS_REQUERIDAS = ["CEDULA", "NOMBRE", "TIPO HE", "VALOR HORA"]

archivos_subidos = st.file_uploader("Selecciona archivos Excel", type=["xlsx", "xls"], accept_multiple_files=True)

if archivos_subidos:
    lista_df = []
    archivos_con_error = False

    for archivo in archivos_subidos:
        try:
            # Leer el archivo
            df = pd.read_excel(archivo)
            
            # Limpiar nombres de columnas (Mayúsculas y sin espacios extra)
            df.columns = [str(c).upper().strip() for c in df.columns]
            
            # 1. Validar columnas faltantes
            faltantes = [col for col in COLUMNAS_REQUERIDAS if col not in df.columns]
            
            if faltantes:
                st.error(f"❌ **{archivo.name}**: Faltan las columnas: {', '.join(faltantes)}")
                archivos_con_error = True
            else:
                # 2. Validar/Convertir tipos de datos
                # Coerce convierte errores a NaN para detectarlos fácilmente
                df['CEDULA'] = pd.to_numeric(df['CEDULA'], errors='coerce')
                df['VALOR HORA'] = pd.to_numeric(df['VALOR HORA'], errors='coerce')
                
                if df['CEDULA'].isnull().any() or df['VALOR HORA'].isnull().any():
                    st.warning(f"⚠️ **{archivo.name}**: Tiene datos no numéricos en Cédula o Valor Hora. Se procede a consolidar el resto de registros.")
                
                # Agregar a la lista para consolidar (solo las columnas que nos interesan)
                lista_df.append(df[COLUMNAS_REQUERIDAS])
                st.success(f"✅ **{archivo.name}** listo para consolidar.")
                
        except Exception as e:
            st.error(f"Error procesando {archivo.name}: {e}")

    # Sección de consolidación
    if lista_df:
        st.divider()
        if archivos_con_error:
            st.warning("Nota: Hay archivos con errores de columnas que no serán incluidos en la descarga final.")

        # Crear el DataFrame único
        df_final = pd.concat(lista_df, ignore_index=True)
        
        st.subheader("Vista previa del consolidado final")
        st.dataframe(df_final)

        # Crear archivo Excel en memoria para la descarga
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Consolidado')
        
        datos_excel = output.getvalue()

        st.download_button(
            label="📥 Descargar Todo en un solo Excel",
            data=datos_excel,
            file_name="consolidado_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )