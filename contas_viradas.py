import streamlit as st
import pandas as pd
from io import BytesIO
from bs4 import BeautifulSoup

def parse_balancete_html(html_content: str) -> pd.DataFrame:
    """
    Faz o parse do arquivo HTML do balancete,
    criando um DataFrame com as colunas relevantes.

    Retorna um DataFrame com:
        - Código da conta
        - Descrição da conta
        - Saldo Atual (valor numérico)
        - Indicador (D ou C) do saldo atual
        - Tipo (Ativo, Passivo, etc.) deduzido a partir do Código
    """

    soup = BeautifulSoup(html_content, "html.parser")
    rows = soup.find_all("tr")
    data_rows = []

    for row in rows:
        cols = row.find_all("td")
        cols_text = [c.get_text(strip=True) for c in cols]

        # Fazemos uma checagem simples para evitar as linhas que não sejam de conta
        if len(cols_text) >= 10:
            codigo = cols_text[0].strip()
            classificacao = cols_text[2].strip()  # pode ou não estar sempre presente

            # Pegar a primeira "descrição" que aparecer nas colunas 4..11
            descricao = None
            for i in range(4, min(len(cols_text), 12)):
                if cols_text[i]:
                    descricao = cols_text[i]
                    break

            # Detectar saldo atual e indicador (D/C)
            saldo_atual_valor = ""
            saldo_atual_indicador = ""
            tail = cols_text[-5:]
            for item in reversed(tail):
                item = item.strip()
                if item.endswith("D") or item.endswith("C"):
                    saldo_atual_valor = item[:-1].strip()
                    saldo_atual_indicador = item[-1]
                    break

            numeric_value = 0.0
            if saldo_atual_valor:
                clean_str = saldo_atual_valor.replace(".", "").replace(",", ".")
                try:
                    numeric_value = float(clean_str)
                except:
                    numeric_value = 0.0

            # Se tiver código ou descrição, adiciona no DataFrame
            if codigo or descricao:
                data_rows.append({
                    "Código": codigo,
                    "Classificação": classificacao,
                    "Descrição": descricao,
                    "Saldo Atual (Valor)": numeric_value,
                    "Saldo Atual (D/C)": saldo_atual_indicador
                })

    df = pd.DataFrame(data_rows)

    # Identificar tipo da conta
    df["Tipo"] = df["Código"].apply(
        lambda x: "Ativo" if x.startswith("1") else
                  "Passivo" if x.startswith("2") else "Outros"
    )

    return df

def marcar_contas_viradas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Marca as contas viradas e cria uma coluna 'Motivo'.
    Regras:
      - Ativo com saldo 'C'
      - Passivo com saldo 'D'
    """
    df = df.copy()

    df['ViradaBool'] = False
    df['Motivo'] = ""

    cond_ativo_c = (df['Tipo'] == 'Ativo') & (df['Saldo Atual (D/C)'] == 'C')
    cond_passivo_d = (df['Tipo'] == 'Passivo') & (df['Saldo Atual (D/C)'] == 'D')

    df.loc[cond_ativo_c, 'ViradaBool'] = True
    df.loc[cond_ativo_c, 'Motivo'] = "Ativo com saldo Credor (C)"

    df.loc[cond_passivo_d, 'ViradaBool'] = True
    df.loc[cond_passivo_d, 'Motivo'] = "Passivo com saldo Devedor (D)"

    # Agora, a coluna 'Virada' mostrará "Sim" ou "Não" para o usuário
    df['Virada'] = df['ViradaBool'].map({True: 'Sim', False: 'Não'})

    return df

def gerar_download_excel(df: pd.DataFrame, nome_arquivo: str) -> None:
    """
    Gera um botão de download para o DataFrame em formato XLSX dentro do Streamlit.
    Recebe o nome do arquivo (e.g. 'todas_contas.xlsx' ou 'contas_viradas.xlsx').
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

        # Contagens de contas viradas
        total_ativo_viradas = df[(df['Tipo'] == 'Ativo') & (df['ViradaBool'] == True)].shape[0]
        total_passivo_viradas = df[(df['Tipo'] == 'Passivo') & (df['ViradaBool'] == True)].shape[0]

        msg_ativo = f"**{total_ativo_viradas}** contas de Ativo viradas"
        msg_passivo = f"**{total_passivo_viradas}** contas de Passivo viradas"

        if total_ativo_viradas > 0 or total_passivo_viradas > 0:
            st.error("Foram encontradas contas viradas no balancete:")
            st.write(msg_ativo + " | " + msg_passivo)
        else:
            st.success("Nenhuma conta virada identificada!")
            st.write(msg_ativo + " | " + msg_passivo)

        st.write("---")

        # Exibir tabela completa
        st.subheader("Tabela Completa de Contas")

        # Destaque em vermelho para linhas viradas
        styled_df = df.style.apply(
            lambda row: [
                'background-color: #ffcccc' if row['ViradaBool'] else ''
                for _ in row
            ],
            axis=1
        )
        st.dataframe(styled_df)

        # Filtro para mostrar apenas contas viradas
        st.write("### Tabela de Contas Viradas")
        df_viradas = df[df['ViradaBool'] == True]
        if not df_viradas.empty:
            st.dataframe(df_viradas)
        else:
            st.info("Não há contas viradas para mostrar.")

        st.write("---")
        
        # Botão para exportar todas as contas
        gerar_download_excel(df, "todas_contas.xlsx")

        # Se existirem viradas, permitir exportar só elas
        if not df_viradas.empty:
            gerar_download_excel(df_viradas, "contas_viradas.xlsx")

if __name__ == "__main__":
    main()