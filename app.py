import streamlit as st
import pandas as pd
import re
import io
import zipfile

def procesar_logica_pedidos(df):
    pedidos_procesados = []
    
    # Diccionario robusto con 12 países y soporte para nombres completos (Alias)
    mapeo_paises = {
        "FR": { "ext": ".fr", "code": "FR", "alias": ["FRANCE", "FRANCIA"] },
        "IT": { "ext": ".it", "code": "IT", "alias": ["ITALY", "ITALIA"] },
        "PT": { "ext": ".pt", "code": "PT", "alias": ["PORTUGAL"] },
        "DE": { "ext": ".de", "code": "DE", "alias": ["GERMANY", "ALEMANIA"] },
        "PL": { "ext": ".pl", "code": "PL", "alias": ["POLAND", "POLONIA"] },
        "SE": { "ext": ".se", "code": "SE", "alias": ["SWEDEN", "SUECIA"] },
        "CZ": { "ext": ".cz", "code": "CZ", "alias": ["CZECH", "CHECO", "REPUBLICA CHECA"] },
        "NL": { "ext": ".nl", "code": "NL", "alias": ["NETHERLANDS", "HOLANDA", "PAISES BAJOS"] },
        "BE": { "ext": ".be", "code": "BE", "alias": ["BELGIUM", "BELGICA"] },
        "AT": { "ext": ".at", "code": "AT", "alias": ["AUSTRIA"] },
        "IE": { "ext": ".ie", "code": "IE", "alias": ["IRELAND", "IRLANDA"] },
        "ES": { "ext": ".es", "code": "ES", "alias": ["SPAIN", "ESPAÑA"] }
    }

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
            # El país suele estar en la última parte de la nota
            dato_pais = partes_notas[-1].upper()
            
            # Valores por defecto (España)
            extension_email = ".es"
            pais_code = "ES"
            encontrado = False

            # Búsqueda inteligente en el mapeo
            for sigla, info in mapeo_paises.items():
                # Si coincide con la sigla o con alguno de los nombres (alias)
                if sigla == dato_pais or any(alias in dato_pais for alias in info.get("alias", [])):
                    extension_email = info["ext"]
                    pais_code = info["code"]
                    encontrado = True
                    break
            
            # Si no está en nuestra lista pero es un código de 2 letras (ej. DK, FI...)
            if not encontrado and len(dato_pais) == 2:
                pais_code = dato_pais
                extension_email = f".{dato_pais.lower()}"

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
st.set_page_config(page_title="Procesador de Pedidos", page_icon="📦")
st.title("📦 Procesador de Pedidos Marketplace")
st.write("Sube tu archivo Excel y el sistema detectará automáticamente los países (incluyendo nombres como 'Germany' o 'Poland').")

archivo_subido = st.file_uploader("Elige un archivo Excel", type=["xlsx", "xls"])

if archivo_subido is not None:
    try:
        df_origen = pd.read_excel(archivo_subido)
        df_resultado = procesar_logica_pedidos(df_origen)

        if not df_resultado.empty:
            st.success(f"✅ Se han procesado {len(df_resultado)} filas correctamente.")
            
            # Mostrar vista previa de los códigos de país detectados para verificar
            with st.expander("Ver desglose por país detectado"):
                st.write(df_resultado['country_code'].value_counts())

            # Crear archivo ZIP en memoria
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "x", zipfile.ZIP_DEFLATED) as zf:
                for mkt_name, grupo in df_resultado.groupby("MARKETPLACE"):
                    nombre_archivo = f"{re.sub(r'[^a-zA-Z0-9]', '_', str(mkt_name))}.xlsx"
                    
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
            st.warning("No se encontraron datos válidos. Revisa el formato de la columna 'notas'.")
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")