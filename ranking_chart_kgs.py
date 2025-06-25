import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os

# Configuração
file_path = r"C:\Users\gmass\Downloads\Margem_250531 - wapp - V3.xlsx"
sheet_name = "Base (3,5%)"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
output_path = os.path.join(output_dir, 'Ranking_Produtos_Tonelagem.pdf')
items_per_page = 10

# Verificar e remover arquivo existente
if os.path.exists(output_path):
    try:
        os.remove(output_path)
    except PermissionError:
        print(f"Erro: Feche o arquivo {output_path} antes de executar")
        exit()

# Ler e processar os dados
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[['CODPRODUTO', 'DESCRICAO', 'QTDE REAL']]
    df = df[df['QTDE REAL'] >= 0]
    
    # Processar dados
    grouped = df.groupby(['CODPRODUTO', 'DESCRICAO'])['QTDE REAL'].sum().reset_index()
    grouped['QTDE REAL'] = grouped['QTDE REAL'].round(3)
    sorted_df = grouped.sort_values('QTDE REAL', ascending=False).reset_index(drop=True)
    sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
    
    # Criar PDF
    with PdfPages(output_path) as pdf:
        for i in range(0, len(sorted_df), items_per_page):
            chunk = sorted_df.iloc[i:i+items_per_page]
            
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 8.5))
            fig.suptitle(f'Ranking de Produtos {i+1}-{min(i+items_per_page, len(sorted_df))}', fontsize=14)
            
            # Tabela
            ax1.axis('off')
            table_data = chunk[['Posição', 'DESCRICAO', 'QTDE REAL']].values
            table = ax1.table(
                cellText=table_data,
                colLabels=['Posição', 'Descrição', 'Tonelagem (kg)'],
                loc='center',
                cellLoc='center'
            )
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.8)
            
            # Gráfico
            wedges, texts, autotexts = ax2.pie(
                chunk['QTDE REAL'],
                labels=chunk['DESCRICAO'],
                autopct='%1.1f%%',
                startangle=140,
                textprops={'fontsize': 8}
            )
            ax2.set_title('Distribuição Percentual', fontsize=12)
            
            plt.tight_layout(rect=[0, 0, 1, 0.95])
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)  # Fechar figura explicitamente
            
    print(f"Relatório gerado com sucesso em: {output_path}")

except Exception as e:
    print(f"Erro ao gerar relatório: {str(e)}")
    if os.path.exists(output_path):
        os.remove(output_path)  # Remover arquivo corrompido em caso de erro