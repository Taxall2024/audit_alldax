import streamlit as st
import pandas as pd
from io import BytesIO
from bs4 import BeautifulSoup

def parse_balancete_html(html_content: str) -> pd.DataFrame:
    """
    Faz o parse do arquivo HTML do balancete, criando um DataFrame com as colunas relevantes.

    Retorna um DataFrame com:
      - Código
      - Classificação
      - Descrição
      - Saldo Atual (Valor numérico)
      - Saldo Atual (D/C)
      - Empresa
      - CNPJ
      - Período
    """
    soup = BeautifulSoup(html_content, "html.parser")
    rows = soup.find_all("tr")
    
    # Extração dos dados do cabeçalho: Empresa, CNPJ e Período
    empresa = None
    cnpj = None
    periodo = None

    # Procura em todas as linhas do cabeçalho (com TH)
    for row in rows:
        ths = row.find_all("th")
        for idx, th in enumerate(ths):
            text = th.get_text(strip=True)
            if text.startswith("Empresa"):
                if len(ths) > idx + 1:
                    empresa = ths[idx+1].get_text(strip=True)
            elif text.startswith("C.N.P.J."):
                if len(ths) > idx + 1:
                    cnpj = ths[idx+1].get_text(strip=True)
            elif text.startswith("Período"):
                if len(ths) > idx + 1:
                    periodo = ths[idx+1].get_text(strip=True)
    
    data_rows = []
    # Processa somente as linhas que contêm células TD (supostamente os dados das contas)
    for row in rows:
        cols = row.find_all("td")
        cols_text = [c.get_text(strip=True) for c in cols]
        # Checa se há colunas suficientes para ser uma linha de conta
        if len(cols_text) >= 10:
            codigo = cols_text[0].strip()
            classificacao = cols_text[2].strip()

            # Localiza a descrição nas colunas 4 a 11
            descricao = None
            for i in range(4, min(len(cols_text), 12)):
                if cols_text[i]:
                    descricao = cols_text[i]
                    break

            # Detecta saldo atual e seu indicador (D/C)
            saldo_atual_valor = ""
            saldo_atual_indicador = ""
            tail = cols_text[-5:]
            for item in reversed(tail):
                item = item.strip()
                if item.endswith("D") or item.endswith("C"):
                    saldo_atual_valor = item[:-1].strip()
                    saldo_atual_indicador = item[-1]
                    break

            # Converter saldo em valor numérico
            numeric_value = 0.0
            if saldo_atual_valor:
                clean_str = saldo_atual_valor.replace(".", "").replace(",", ".")
                try:
                    numeric_value = float(clean_str)
                except:
                    numeric_value = 0.0

            if codigo or descricao:
                data_rows.append({
                    "Código": codigo,
                    "Classificação": classificacao,
                    "Descrição": descricao,
                    "Saldo Atual (Valor)": numeric_value,
                    "Saldo Atual (D/C)": saldo_atual_indicador,
                    "Empresa": empresa,
                    "CNPJ": cnpj,
                    "Período": periodo
                })

    df = pd.DataFrame(data_rows)
    return df

def marcar_contas_viradas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ajusta as colunas 'Virada', 'Motivo' e 'Avaliar' com base na regra:
      - Se a Classificação inicia com '1' e o Saldo Atual (D/C) é 'C' => conta virada
      - Se a Classificação inicia com '2' e o Saldo Atual (D/C) é 'D' => conta virada
      - Caso contrário, marca "Avaliar no detalhe"

    Além disso, se a Descrição iniciar com "(-)", a conta (redutora) não deve ser considerada virada.
    """
    df = df.copy()

    # Cria as colunas padrão
    df['ViradaBool'] = False
    df['Virada'] = "Não"
    df['Motivo'] = ""
    df['Avaliar'] = "Avaliar no detalhe"

    # Condições para marcar como conta virada
    cond_ativo_c = df['Classificação'].str.startswith('1') & (df['Saldo Atual (D/C)'] == 'C')
    cond_passivo_d = df['Classificação'].str.startswith('2') & (df['Saldo Atual (D/C)'] == 'D')

    df.loc[cond_ativo_c, 'ViradaBool'] = True
    df.loc[cond_ativo_c, 'Virada'] = "Sim"
    df.loc[cond_ativo_c, 'Motivo'] = "Ativo (1) com saldo Credor (C)"
    df.loc[cond_ativo_c, 'Avaliar'] = ""

    df.loc[cond_passivo_d, 'ViradaBool'] = True
    df.loc[cond_passivo_d, 'Virada'] = "Sim"
    df.loc[cond_passivo_d, 'Motivo'] = "Passivo (2) com saldo Devedor (D)"
    df.loc[cond_passivo_d, 'Avaliar'] = ""

    # ----- REGRA DE CONSIDERAÇÃO DE CONTAS REDUTORAS -----
    # Se a Descrição iniciar com "(-)", a conta não deve ser considerada virada
    cond_redutora = df['Descrição'].str.startswith("(-)") & (df['ViradaBool'] == True)
    df.loc[cond_redutora, 'ViradaBool'] = False
    df.loc[cond_redutora, 'Virada'] = "Não"
    df.loc[cond_redutora, 'Motivo'] = ""
    df.loc[cond_redutora, 'Avaliar'] = "Avaliar no detalhe"

    return df

def gerar_download_excel(df: pd.DataFrame, nome_arquivo: str) -> None:
    """
    Gera um botão de download para o DataFrame em formato XLSX dentro do Streamlit.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planilha')
    data = output.getvalue()

    st.download_button(
        label=f"Download: {nome_arquivo}",
        data=data,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def main():
    st.title("Verificador de Contas Viradas em Balancete")

    uploaded_file = st.file_uploader("Arraste o arquivo HTML do balancete aqui", type=["htm", "html"])
    
    if uploaded_file is not None:
        html_content = uploaded_file.read().decode('utf-8', errors='ignore')
        
        with st.spinner("Processando balancete..."):
            df = parse_balancete_html(html_content)
            df = marcar_contas_viradas(df)

        if df.empty:
            st.warning("Não foi possível encontrar dados de contas no arquivo enviado.")
            return

        # Contagem de contas viradas
        df_viradas = df[df['ViradaBool'] == True]
        total_viradas = df_viradas.shape[0]

        if total_viradas > 0:
            st.error(f"Foram encontradas {total_viradas} contas viradas no balancete.")
        else:
            st.success("Nenhuma conta virada identificada!")

        st.write("---")

        # Exibe tabela completa
        st.subheader("Tabela Completa de Contas")
        styled_df = df.style.apply(
            lambda row: ['background-color: #ffcccc' if row['ViradaBool'] else '' for _ in row],
            axis=1
        )
        st.dataframe(styled_df)

        # Exibe tabela somente das contas viradas
        st.subheader("Tabela de Contas Viradas")
        if not df_viradas.empty:
            st.dataframe(df_viradas)
        else:
            st.info("Não há contas viradas para mostrar.")

        st.write("---")
        # Botões para exportar os dados em Excel
        gerar_download_excel(df, "todas_contas.xlsx")
        if not df_viradas.empty:
            gerar_download_excel(df_viradas, "contas_viradas.xlsx")

if __name__ == "__main__":
    main()