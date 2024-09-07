import pandas as pd
import streamlit as st
from tempfile import NamedTemporaryFile
import os
import shutil

icone="/workspaces/Dks_Solucoes/logoCortado_free-file.ico"

st.set_page_config(
    page_title="Dks Soluções", 
    page_icon=icone,  
    layout="wide"
)

# Função para processar o arquivo Excel
def process_excel(file_path, file_name):
    df = pd.read_excel(file_path, header=None)

    date_str = str(df.iloc[2, 6])[-10:]
    Ano_str = str(df.iloc[2, 6])[-4:]

    try:
        period_month = pd.to_datetime(date_str, dayfirst=True).strftime('%B')
    except ValueError:
        period_month = "Unknown"

    df['Período'] = period_month
    df['Ano'] = Ano_str

    new_header = df.iloc[7]
    df = df[8:]
    df.columns = new_header

    df = df.dropna(axis=1, how='all')
    df.columns = df.columns.str.strip()
    df_cleaned = df.dropna(axis=1, how='all')
    df_cleaned.columns = [f"Coluna_{i+1}" for i in range(df_cleaned.shape[1])]
    df_cleaned = df_cleaned.dropna(subset=['Coluna_4'])
    df_cleaned = df_cleaned.drop(columns=['Coluna_2', 'Coluna_3', 'Coluna_10', 'Coluna_12', 'Coluna_14', 'Coluna_16', 'Coluna_18'])

    num_columns = df_cleaned.shape[1]
    new_column_names = [
        "Código", "Classificação", "Descrição da conta ClassificacaoN1", 
        "DescriçãoN2", "DescriçãoN3", 
        "DescriçãoN4", "Descrição Detalhada", "Saldo Anterior", 
        "Débito", "Crédito", "Saldo Atual", "Período", "Ano"
    ]

    if num_columns > len(new_column_names):
        extra_columns = [f"Extra_{i}" for i in range(num_columns - len(new_column_names))]
        new_column_names.extend(extra_columns)

    df_cleaned.columns = new_column_names

    def preencher_colunas(df, coluna_classificacao, coluna_descricao, num_caracteres):
        df[coluna_classificacao] = df[coluna_classificacao].astype(str)
        agrupado = df.groupby(df[coluna_classificacao].str[:num_caracteres])[coluna_descricao].first().reset_index()
        agrupado.columns = [f"{coluna_classificacao}_agrupado", f"{coluna_descricao}_agrupado"]
        df = df.merge(agrupado, left_on=df[coluna_classificacao].str[:num_caracteres], right_on=f"{coluna_classificacao}_agrupado", how='left')
        df[coluna_descricao] = df[coluna_descricao].combine_first(df[f"{coluna_descricao}_agrupado"])
        df.drop(columns=[f"{coluna_classificacao}_agrupado", f"{coluna_descricao}_agrupado"], inplace=True)
        return df

    df_cleaned = preencher_colunas(df_cleaned, 'Classificação', 'Descrição da conta ClassificacaoN1', 1)
    df_cleaned = preencher_colunas(df_cleaned, 'Classificação', 'DescriçãoN2', 3)
    df_cleaned = preencher_colunas(df_cleaned, 'Classificação', 'DescriçãoN3', 5)
    df_cleaned = preencher_colunas(df_cleaned, 'Classificação', 'DescriçãoN4', 8)
    df_cleaned = preencher_colunas(df_cleaned, 'Classificação', 'Descrição Detalhada', 12)

    df_cleaned = df_cleaned.dropna(subset=['Classificação'])

    base_name = os.path.splitext(file_name)[0]
    processed_file_name = f"{base_name}_Processado.xlsx"

    with NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df_cleaned.to_excel(tmp.name, index=False)
        return tmp.name, processed_file_name

def download_to_user_folder(output_file_path):
    user_download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    if not os.path.exists(user_download_folder):
        os.makedirs(user_download_folder)
    destination_path = os.path.join(user_download_folder, os.path.basename(output_file_path))
    shutil.move(output_file_path, destination_path)
    return destination_path

st.title("Processador de Arquivos Excel")

uploaded_file = st.file_uploader("Escolha um arquivo Excel", type=["xlsx"])

if uploaded_file is not None:
    with NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    processed_file_path, processed_file_name = process_excel(tmp_path, uploaded_file.name)

    with open(processed_file_path, "rb") as file:
        btn = st.download_button(
            label="Baixar arquivo processado",
            data=file,
            file_name=processed_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
