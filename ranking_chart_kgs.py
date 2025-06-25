import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import shutil
from datetime import datetime
import psutil

def check_disk_space(path, min_space_gb=1):
    """Verifica se há espaço suficiente em disco"""
    usage = shutil.disk_usage(os.path.dirname(path))
    return usage.free > (min_space_gb * 1024**3)

def clean_matplotlib_memory():
    """Limpa a memória do matplotlib"""
    plt.close('all')
    import gc
    gc.collect()

# Configuração
file_path = r"C:\Users\gmass\Downloads\Margem_250531 - wapp - V3.xlsx"
sheet_name = "Base (3,5%)"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
output_path = os.path.join(output_dir, 'Ranking_Produtos_Tonelagem.pdf')
temp_path = os.path.join(output_dir, 'temp_Ranking.pdf')
items_per_page = 10

try:
    # Verificar espaço em disco
    if not check_disk_space(output_path):
        raise RuntimeError("Espaço insuficiente em disco para gerar o relatório")

    # Verificar e remover arquivos existentes
    for path in [output_path, temp_path]:
        if os.path.exists(path):
            try:
                os.remove(path)
            except PermissionError:
                raise RuntimeError(f"Feche o arquivo {path} antes de executar")

    # Ler e processar os dados
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
        ['CODPRODUTO', 'DESCRICAO', 'QTDE REAL', 'DATA']]
    df = df[df['QTDE REAL'] >= 0]
    
    # Converter DATA para datetime
    df['DATA'] = pd.to_datetime(df['DATA'])
    
    # Processar dados para o ranking
    grouped = df.groupby(['CODPRODUTO', 'DESCRICAO'])['QTDE REAL'].sum().reset_index()
    grouped['QTDE REAL'] = grouped['QTDE REAL'].round(3)
    sorted_df = grouped.sort_values('QTDE REAL', ascending=False).reset_index(drop=True)
    sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
    
    # Processar dados para série temporal
    time_series = df.groupby(['CODPRODUTO', 'DESCRICAO', 'DATA'])['QTDE REAL'].sum().reset_index()
    
    # Criar PDF temporário primeiro
    with PdfPages(temp_path) as pdf:
        for i in range(0, len(sorted_df), items_per_page):
            clean_matplotlib_memory()
            
            chunk = sorted_df.iloc[i:i+items_per_page]
            produtos_na_pagina = chunk['CODPRODUTO'].tolist()
            
            fig = plt.figure(figsize=(11, 14), constrained_layout=True)
            gs = fig.add_gridspec(3, 1)
            ax1 = fig.add_subplot(gs[0])  # Tabela
            ax2 = fig.add_subplot(gs[1])  # Pizza
            ax3 = fig.add_subplot(gs[2])  # Linha
            
            fig.suptitle(f'Ranking de Produtos {i+1}-{min(i+items_per_page, len(sorted_df))}', 
                        fontsize=14, y=1.02)
            
            # Tabela
            ax1.axis('off')
            table_data = chunk[['Posição', 'DESCRICAO', 'QTDE REAL']].values
            table = ax1.table(
                cellText=table_data,
                colLabels=['Posição', 'Descrição', 'Tonelagem (kg)'],
                loc='center',
                cellLoc='center',
                colWidths=[0.1, 0.6, 0.3]
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1, 1.3)
            
            # Gráfico de Pizza - Legenda acima
            ax2.set_title('Distribuição Percentual', fontsize=10, pad=10)
            wedges, texts, autotexts = ax2.pie(
                chunk['QTDE REAL'],
                autopct=lambda p: f'{p:.1f}%\n({p*sum(chunk["QTDE REAL"])/100:.1f} kg)',
                startangle=140,
                textprops={'fontsize': 7},
                wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                pctdistance=0.85
            )
            
            # Legenda otimizada para usar toda a largura
            n_cols = min(4, len(chunk))  # Máximo de 4 colunas, mas ajusta automaticamente
            ax2.legend(wedges, chunk['DESCRICAO'],
                      title="Produtos",
                      loc="upper center",
                      bbox_to_anchor=(0.5, -0.05),  # Posicionada abaixo do gráfico
                      ncol=n_cols,
                      fontsize=7,
                      title_fontsize=8,
                      frameon=False)
            
            # Gráfico de Linha - Legenda acima
            ts_filtered = time_series[time_series['CODPRODUTO'].isin(produtos_na_pagina)]
            colors = [w.get_facecolor() for w in wedges]
            
            ax3.set_title('Evolução Temporal', fontsize=10, pad=10)
            ax3.set_ylabel('Tonelagem (kg)', fontsize=8)
            
            lines = []  # Para armazenar as linhas para a legenda
            labels = []  # Para armazenar os labels
            for idx, (produto, group) in enumerate(ts_filtered.groupby('CODPRODUTO')):
                group = group.sort_values('DATA')
                line, = ax3.plot(group['DATA'], group['QTDE REAL'], 
                               marker='o', linestyle='-', 
                               color=colors[idx], 
                               markersize=4, linewidth=1.5)
                lines.append(line)
                labels.append(group['DESCRICAO'].iloc[0])
            
            # Legenda otimizada para usar toda a largura
            n_cols = min(4, len(lines))  # Máximo de 4 colunas, mas ajusta automaticamente
            ax3.legend(lines, labels,
                      loc="upper center",
                      bbox_to_anchor=(0.5, -0.2),  # Posicionada abaixo do gráfico
                      ncol=n_cols,
                      fontsize=7,
                      frameon=False)
            
            ax3.grid(True, linestyle=':', alpha=0.5)
            plt.setp(ax3.get_xticklabels(), rotation=30, ha='right', fontsize=7)
            plt.setp(ax3.get_yticklabels(), fontsize=7)
            
            # Ajustar layout para acomodar as legendas
            plt.subplots_adjust(hspace=0.5)
            
            # Configurações do PDF
            pdf.savefig(fig, dpi=150, bbox_inches='tight', pad_inches=0.5)
            plt.close(fig)
            
    # Renomear arquivo temporário para final
    os.rename(temp_path, output_path)
    print(f"Relatório gerado com sucesso em: {output_path}")
    print(f"Tamanho do arquivo: {os.path.getsize(output_path)/1024/1024:.2f} MB")

except Exception as e:
    print(f"ERRO: {str(e)}")
    # Remover arquivos temporários em caso de erro
    for path in [output_path, temp_path]:
        if os.path.exists(path):
            try:
                os.remove(path)
            except:
                pass
    # Sugestões para solução
    if "PermissionError" in str(e):
        print("\nDICA: Feche todos os visualizadores de PDF antes de executar")
    elif "MemoryError" in str(e):
        print("\nDICA: Reduza items_per_page ou libere memória RAM")
    elif "disk" in str(e).lower():
        print("\nDICA: Libere espaço em disco (pelo menos 1GB recomendado)")