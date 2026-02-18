import streamlit as st
import pandas as pd
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="SAP Converter Web - Julian Drones", layout="centered")

st.title("🚀 Conversor de Carga SAP")
st.markdown("### Processamento Automático e Ordenação por VBELN")
st.write("Suba seu arquivo Excel (A=Documento, B=Valor) para gerar o TXT formatado.")

# Upload do arquivo
arquivo_upload = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls"])

if arquivo_upload:
    try:
        # Lendo o Excel
        df_excel = pd.read_excel(arquivo_upload, header=None, dtype=str, engine='openpyxl')

        linhas_formatadas = []

        for index, row in df_excel.iterrows():
            vbeln = str(row[0]).strip().zfill(8)
            valor_celula = str(row[1]).strip()

            is_date = False
            # Lógica de detecção inteligente
            if any(char in valor_celula for char in ['/', '-']) or len(valor_celula) >= 6:
                try:
                    data_dt = pd.to_datetime(valor_celula, dayfirst=True, errors='raise')
                    atnam = "YAE_ALLOWED_INCR_DATE"
                    valor_final = data_dt.strftime('%d%m%Y')
                    is_date = True
                except:
                    is_date = False

            if not is_date:
                atnam = "YAE_INDEX_RULE"
                valor_final = valor_celula

            linhas_formatadas.append({
                'VBELN': vbeln, 'POSNR': '', 'STEP': '', 'ATNAM': atnam, 'ATWRT': f";;;{valor_final}"
            })

        df_final = pd.DataFrame(linhas_formatadas)

        # Limpeza de duplicatas e Ordenação
        total_antes = len(df_final)
        df_final = df_final.drop_duplicates(subset=['VBELN', 'ATNAM'], keep='first')
        df_final = df_final.sort_values(by=['VBELN', 'ATNAM'], ascending=[True, False])

        total_depois = len(df_final)
        removidos = total_antes - total_depois

        # Exibição de resultados na tela
        st.success(f"✅ Processamento concluído! {total_depois} registros gerados.")
        if removidos > 0:
            st.warning(f"⚠️ {removidos} duplicatas de características no mesmo documento foram removidas.")

        # Preparar o arquivo para download
        txt_output = df_final.to_csv(sep='\t', index=False)

        st.download_button(
            label="📥 Baixar Arquivo TXT para SAP",
            data=txt_output,
            file_name="carga_sap_web.txt",
            mime="text/plain"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")

st.markdown("---")
st.caption("v10.0 Web | Desenvolvido Julian Oliveira")