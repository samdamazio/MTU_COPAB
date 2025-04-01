import pandas as pd

# -------------------------------
# Carrega os dados da planilha_copab
# -------------------------------

# Tabela MESTRE:
# - Sheet: "MESTRE"
# - Os dados começam na linha 7 (skiprows=6) e o PN está na coluna H (índice 7)
df_mestre = pd.read_excel("restricted_data/planilha_copab.xlsx", sheet_name="MESTRE", skiprows=6)
# Cria uma coluna 'PN' extraída da coluna H (índice 7)
df_mestre["PN"] = df_mestre.iloc[:, 7]

# Tabela Cotação:
# - Sheet: "Cotação"
# - Os dados começam na linha 4 (skiprows=3) e o PN está na coluna A (índice 0)
df_cotacao = pd.read_excel("restricted_data/planilha_copab.xlsx", sheet_name="Cotação", skiprows=3)
df_cotacao["PN"] = df_cotacao.iloc[:, 0]

# -------------------------------
# Carrega os dados da planilha_mtu
# -------------------------------

# Os dados começam na linha 2 (skiprows=1) e o PN está na coluna A (índice 0)
df_mtu = pd.read_excel("restricted_data/planilha_mtu.xlsx", skiprows=1)
df_mtu["PN"] = df_mtu.iloc[:, 0]

# -------------------------------
# Realiza a junção (merge) dos dados com base no PN
# -------------------------------

# Realiza merge entre a tabela MESTRE e a planilha_mtu
merged_mestre = pd.merge(df_mestre, df_mtu, on="PN", how="inner", suffixes=("_mestre", "_mtu"))

# Realiza merge entre a tabela Cotação e a planilha_mtu
merged_cotacao = pd.merge(df_cotacao, df_mtu, on="PN", how="inner", suffixes=("_cotacao", "_mtu"))

# Combina os dois resultados em um único DataFrame
df_final = pd.concat([merged_mestre, merged_cotacao], ignore_index=True)

# -------------------------------
# Salva o resultado em uma nova planilha
# -------------------------------
df_final.to_excel("restricted_data/planilha_final.xlsx", index=False)

print("Planilha final gerada com sucesso!")
