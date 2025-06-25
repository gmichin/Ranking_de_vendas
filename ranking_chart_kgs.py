import pandas as pd
import os

# Caminho do arquivo e aba específica
file_path = r"C:\Users\gmass\Downloads\Margem_250531 - wapp - V3.xlsx"
sheet_name = "Base (3,5%)"

# Ler a planilha, começando da linha 9 (já que os dados começam em A9)
df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)

# Selecionar apenas as colunas necessárias
cols = ['CODPRODUTO', 'DESCRICAO', 'QTDE REAL']
df = df[cols]

# Agrupar por CODPRODUTO e DESCRICAO, calculando a média de QTDE REAL
grouped = df.groupby(['CODPRODUTO', 'DESCRICAO'])['QTDE REAL'].mean().reset_index()

# Ordenar por QTDE REAL em ordem decrescente
sorted_df = grouped.sort_values('QTDE REAL', ascending=False).reset_index(drop=True)

# Adicionar coluna de colocação (ranking)
sorted_df.insert(0, 'Colocação', range(1, len(sorted_df) + 1))

# Selecionar apenas as colunas que queremos no output
output_df = sorted_df[['Colocação', 'DESCRICAO', 'QTDE REAL']]

# Dividir o ranking em grupos de 10
chunks = [output_df[i:i+10] for i in range(0, len(output_df), 10)]

# Criar um arquivo Excel com várias abas para cada grupo
output_path = os.path.join(os.path.expanduser('~'), 'Downloads', 'Ranking_Produtos.xlsx')

with pd.ExcelWriter(output_path) as writer:
    for i, chunk in enumerate(chunks, 1):
        # Nome da aba será "Ranking 1-10", "Ranking 11-20", etc.
        start = (i-1)*10 + 1
        end = i*10 if i*10 <= len(output_df) else len(output_df)
        sheet_name = f'Ranking {start}-{end}'
        chunk.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Arquivo salvo com sucesso em: {output_path}")