import streamlit as st
import pandas as pd
import re
import io
import zipfile

def procesar_logica_pedidos(df):
    pedidos_procesados = []
    for index, row in df.iterrows():
        # --- LÓGICA SKU ---
        sku = str(row['REFERENCIA']).strip()
        if not sku.upper().startswith("A"):
            sku = re.sub(r'^[sS]', '', sku, flags=re.IGNORECASE)
            sku = sku.lstrip('0')

        # --- EXTRACCIÓN DEL PAÍS ---
        notas_raw = str(row['notas'])
        partes_notas = [n.strip() for n in notas_raw.split(';') if n.strip()]

        if len(partes_notas) >= 2:
            dato_pais = partes_notas[-1].upper()
            if "FR" in dato_pais:
                extension_email, pais_code = ".fr", "FR"
            elif "IT" in dato_pais:
                extension_email, pais_code = ".it", "IT"
            elif "PT" in dato_pais:
                extension_email, pais_code = ".pt", "PT"
            else:
                extension_email, pais_code = ".es", "ES"

            mkt = str(row['MARKETPLACE'])
            mkt_limpio = re.sub(r'[^a-zA-Z0-9]', '', mkt).lower()
            email_final = f"{row['Nº PEDIDO MARKETPLACE']}@{mkt_limpio}{extension_email}"

            pedidos_procesados.append({
                "MARKETPLACE": mkt,
                "article": sku,
                "quantity": row['CANTIDAD'],
                "customer_name": partes_notas[0],
                "nif": "",
                "attention_of": partes_notas[0],
                "address": partes_notas[2] if len(partes_notas) > 2 else "",
                "postal_code": partes_notas[3] if len(partes_notas) > 3 else "",
                "phone": partes_notas[1] if len(partes_notas) > 1 else "",
                "city": partes_notas[4] if len(partes_notas) > 4 else "",
                "country_code": pais_code,
                "customer_email": email_final,
                "comment": row['Nº PEDIDO PRESTASHOP'] if pd.notna(row['Nº PEDIDO PRESTASHOP']) else "",
                "addressee_order_number": row['Nº PEDIDO MARKETPLACE']
            })
    return pd.DataFrame(pedidos_procesados)

# --- INTERFAZ STREAMLIT ---
st.title("📦 Procesador de Pedidos Marketplace")
st.write("Sube tu archivo 'UnidadNueva' y descarga los pedidos unificados.")

archivo_subido = st.file_uploader("Elige un archivo Excel", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        df_origen = pd.read_excel(archivo_subido)
        df_resultado = procesar_logica_pedidos(df_origen)

        if not df_resultado.empty:
            st.success(f"Se han procesado {len(df_resultado)} filas correctamente.")
            
            # Crear un archivo ZIP en memoria para la descarga
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "x", zipfile.ZIP_DEFLATED) as zf:
                for mkt_name, grupo in df_resultado.groupby("MARKETPLACE"):
                    nombre_archivo = f"{re.sub(r'[^a-zA-Z0-9]', '_', str(mkt_name))}.xlsx"
                    
                    # Guardar cada grupo en un Excel dentro del ZIP
                    output_excel = io.BytesIO()
                    columnas = ["article", "quantity", "customer_name", "nif", "attention_of", 
                                "address", "postal_code", "phone", "city", "country_code", 
                                "customer_email", "comment", "addressee_order_number"]
                    grupo[columnas].to_excel(output_excel, index=False)
                    zf.writestr(nombre_archivo, output_excel.getvalue())

            st.download_button(
                label="📥 Descargar Pedidos Unificados (ZIP)",
                data=buf.getvalue(),
                file_name="pedidos_procesados.zip",
                mime="application/zip"
            )
        else:
            st.warning("No se encontraron datos válidos para procesar.")
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")