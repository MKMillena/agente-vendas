# agente-vendas

# Agente de Automação de Vendas (Excel -> Excel)

Este projeto é uma aplicação Streamlit que automatiza o cruzamento de dados entre uma planilha de vendas e uma base de vendedores.

## Como Usar

1.  **Acesse a Aplicação**: (Link do seu deploy no Streamlit Cloud)
2.  **Passo 1**: Faça o upload da **Base de Vendedores** (Excel).
    *   Estrutura esperada: Colunas alternadas `Vendedor | Cliente`.
    *   O sistema varre todas as colunas procurando pares.
3.  **Passo 2**: Faça o upload da **Planilha de Vendas** (Excel).
    *   Estrutura esperada: Colunas `Data Aprovação`, `Clientes`, `Valor Total`.
4.  **Download**: Baixe o relatório gerado (`Relatorio_Vendas_Consolidado.xlsx`) com a coluna `Vendedor` preenchida.

## Instalação Local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Como Fazer o Deploy (Streamlit Community Cloud)

1.  Suba este código para um repositório no **GitHub**.
2.  Acesse [share.streamlit.io](https://share.streamlit.io/).
3.  Clique em **New app**.
4.  Selecione o repositório, a branch (`main`) e o arquivo principal (`app.py`).
5.  Clique em **Deploy**.
