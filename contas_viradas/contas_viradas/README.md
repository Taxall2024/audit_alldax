# Audit AllDAX - Verificador de Contas Viradas

Este módulo faz parte do projeto Audit AllDAX e é responsável por identificar contas contábeis "viradas" em balancetes.

## Funcionalidades

- Análise de balancetes em formato HTML
- Identificação automática de contas viradas
- Exportação dos resultados em Excel
- Interface web amigável usando Streamlit

## Como usar

1. Execute o aplicativo Streamlit:
```bash
streamlit run contas_viradas/contas_viradas.py
```
2. Faça o upload de um arquivo HTML contendo o balancete.
3. Clique no botão "Analisar Balancete" para iniciar a análise.
4. Faça o download dos resultados em formato Excel.

## Regras de Negócio

- Contas do Ativo (1.x.x.x) com saldo Credor são marcadas como viradas
- Contas do Passivo (2.x.x.x) com saldo Devedor são marcadas como viradas