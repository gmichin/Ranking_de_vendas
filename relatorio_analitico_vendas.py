import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import shutil
from datetime import timedelta
from matplotlib import patheffects

# Dicionário para traduzir os meses para português
MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

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
items_per_page = 5

try:
    # Ler os dados primeiro para obter as datas
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
        ['CODPRODUTO', 'DESCRICAO', 'QTDE REAL', 'DATA']]
    df = df[df['QTDE REAL'] >= 0]
    
    # Converter DATA para datetime e obter mês/ano
    df['DATA'] = pd.to_datetime(df['DATA'])
    primeiro_mes = df['DATA'].iloc[0].month
    primeiro_ano = df['DATA'].iloc[0].year
    
    # Obter nome do mês em português
    nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
    
    # Criar nome do arquivo com as variáveis
    output_filename = f"Relatório Analítico de Vendas - {nome_mes} {primeiro_ano} - {items_per_page} em {items_per_page}.pdf"
    output_path = os.path.join(output_dir, output_filename)
    temp_path = os.path.join(output_dir, f"temp_{output_filename}")

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

    # Processar dados para o ranking
    grouped = df.groupby(['CODPRODUTO', 'DESCRICAO'])['QTDE REAL'].sum().reset_index()
    grouped['QTDE REAL'] = grouped['QTDE REAL'].round(3)
    sorted_df = grouped.sort_values('QTDE REAL', ascending=False).reset_index(drop=True)
    sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
    
    # Processar dados para série temporal - Agrupando por semana corretamente
    time_series = df.copy()
    time_series['SEMANA'] = time_series['DATA'].dt.to_period('W').dt.start_time
    time_series = time_series.groupby(['CODPRODUTO', 'DESCRICAO', 'SEMANA'])['QTDE REAL'].sum().reset_index()
    
    # Criar PDF temporário primeiro
    with PdfPages(temp_path) as pdf:
        for i in range(0, len(sorted_df), items_per_page):
            clean_matplotlib_memory()
            
            chunk = sorted_df.iloc[i:i+items_per_page]
            produtos_na_pagina = chunk['CODPRODUTO'].tolist()
            
            fig = plt.figure(figsize=(11, 16), constrained_layout=True)  # Aumentado para 16 para acomodar o novo gráfico
            gs = fig.add_gridspec(4, 1)  # Agora são 4 linhas
            ax1 = fig.add_subplot(gs[0])  # Tabela
            ax2 = fig.add_subplot(gs[1])  # Pizza
            ax3 = fig.add_subplot(gs[2])  # Linha
            ax4 = fig.add_subplot(gs[3])  # Barras - novo gráfico
            
            fig.suptitle(f'Ranking de Produtos {i+1}-{min(i+items_per_page, len(sorted_df))}', 
                        fontsize=14, y=1.02)
            
            # Tabela (mesmo código anterior)
            ax1.axis('off')
            table_data = chunk[['Posição', 'CODPRODUTO', 'DESCRICAO', 'QTDE REAL']].values
            table = ax1.table(
                cellText=table_data,
                colLabels=['Posição', 'Código', 'Descrição', 'Tonelagem (kg)'],
                loc='center',
                cellLoc='center',
                colWidths=[0.1, 0.1, 0.5, 0.3]
            )
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1, 1.3)
            
            # Gráfico de Pizza 
            ax2.set_title('Distribuição Percentual', fontsize=10, pad=10)
            
            # Opções: hsv
            colormap_name = 'jet'
            
            # Criar um colormap apenas para os produtos desta página
            num_produtos = len(produtos_na_pagina)
            colors = plt.colormaps[colormap_name].resampled(num_produtos)
            
            # Aumentar a saturação e brilho das cores
            def boost_color(color, saturation_factor=1.3, brightness_factor=1.1):
                import colorsys
                r, g, b, a = color
                h, s, v = colorsys.rgb_to_hsv(r, g, b)
                s = min(1.0, s * saturation_factor)
                v = min(1.0, v * brightness_factor)
                r, g, b = colorsys.hsv_to_rgb(h, s, v)
                return (r, g, b, a)
            
            # Aplicar cores boosteadas
            boosted_colors = [boost_color(colors(i)) for i in range(num_produtos)]
            
            # Adicione esta função
            def get_contrast_color(color):
                """Retorna preto ou branco dependendo do brilho da cor de fundo"""
                r, g, b, a = color
                brightness = (0.299 * r + 0.587 * g + 0.114 * b)
                return 'black' if brightness > 0.5 else 'white'

            # Modifique a criação do gráfico de pizza:
            wedges, texts, autotexts = ax2.pie(
                chunk['QTDE REAL'],
                autopct=lambda p: f'{p:.1f}%\n({p*sum(chunk["QTDE REAL"])/100:.1f} kg)',
                startangle=140,
                textprops={
                    'fontsize': 7,
                    'color': 'white',  # Cor principal do texto
                    'path_effects': [
                        patheffects.withStroke(linewidth=2, foreground='black'),  # Contorno
                        patheffects.Normal()  # Garante que o texto principal seja visível
                    ]
                },
                wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                pctdistance=0.85,
                colors=boosted_colors
            )

            # Aplicar cores dinâmicas
            # Aplicar cores dinâmicas para melhor contraste
            for text, wedge in zip(autotexts, wedges):
                # Mantém o texto branco com contorno preto
                text.set_color('white')
                text.set_path_effects([
                    patheffects.withStroke(linewidth=2, foreground='black'),
                    patheffects.Normal()
                ])

            # Criar mapeamento de cores para os produtos desta página
            product_colors = {prod: boosted_colors[i] for i, prod in enumerate(produtos_na_pagina)}
            
            # Legenda otimizada para usar toda a largura
            n_cols = min(4, len(chunk))
            ax2.legend(wedges, chunk['DESCRICAO'],
                      loc="upper center",
                      bbox_to_anchor=(0.5, -0.05),
                      ncol=n_cols,
                      fontsize=7,
                      title_fontsize=8,
                      frameon=False)
            
            # Gráfico de Linha (agora com um ponto por semana)
            ts_filtered = time_series[time_series['CODPRODUTO'].isin(produtos_na_pagina)]
            
            ax3.set_title('Evolução Temporal (por semana)', fontsize=10, pad=10)
            ax3.set_ylabel('Tonelagem (kg)', fontsize=8)
            
            lines = []
            labels = []
            for produto, group in ts_filtered.groupby('CODPRODUTO'):
                group = group.sort_values('SEMANA')
                line_color = product_colors[produto]  # Usar a mesma cor do gráfico de pizza

                line, = ax3.plot(group['SEMANA'], group['QTDE REAL'], 
                               marker='o', linestyle='-', 
                               color=line_color,
                               markersize=4, linewidth=1.5)
                lines.append(line)
                labels.append(group['DESCRICAO'].iloc[0])

                for x, y in zip(group['SEMANA'], group['QTDE REAL']):
                    annotation = ax3.annotate(f'{y:.1f}', 
                                            xy=(x, y),
                                            xytext=(0, 5),
                                            textcoords='offset points',
                                            ha='center', va='bottom',
                                            fontsize=6,
                                            color='white')  # Texto branco

                    # Adiciona contorno com a cor da linha
                    annotation.set_path_effects([
                        patheffects.withStroke(linewidth=2, foreground=line_color),
                        patheffects.Normal()
                    ])
            # Formatando o eixo x para mostrar o período semanal
            ax3.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d/%m'))
            ax3.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(byweekday=plt.matplotlib.dates.MO))
            
            # Criando labels que mostram o período da semana
            tick_labels = []
            for semana in ax3.get_xticks():
                semana_inicio = plt.matplotlib.dates.num2date(semana)
                semana_fim = semana_inicio + timedelta(days=6)
                tick_labels.append(f"{semana_inicio.strftime('%d/%m')}\na\n{semana_fim.strftime('%d/%m')}")
            
            ax3.set_xticklabels(tick_labels)
            
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
            
            # NOVO GRÁFICO DE BARRAS
            ax4.set_title('Quantidade por Produto', fontsize=10, pad=10)
            ax4.set_ylabel('Tonelagem (kg)', fontsize=8)
            
            # Criar as barras com as mesmas cores dos outros gráficos
            bars = ax4.bar(
                chunk['DESCRICAO'],
                chunk['QTDE REAL'],
                color=[product_colors[p] for p in produtos_na_pagina]
            )
            
            # Adicionar os valores em cima de cada barra
            for bar in bars:
                height = bar.get_height()
                ax4.annotate(f'{height:.1f}',
                            xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3),  # 3 points vertical offset
                            textcoords="offset points",
                            ha='center', va='bottom',
                            fontsize=8)
            
            # Rotacionar os labels do eixo x para melhor legibilidade
            plt.setp(ax4.get_xticklabels(), rotation=15, ha='right', fontsize=8)
            plt.setp(ax4.get_yticklabels(), fontsize=7)
            
            # Adicionar grid horizontal
            ax4.grid(True, axis='y', linestyle=':', alpha=0.5)
            
            # Ajustar layout para acomodar todos os gráficos
            plt.subplots_adjust(hspace=0.7)
            
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