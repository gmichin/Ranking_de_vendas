import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import shutil
from datetime import timedelta
from matplotlib import patheffects
import re

# Suprimir warnings do matplotlib
warnings.filterwarnings("ignore", category=UserWarning)

# Dicionário para traduzir os meses para português
MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

TOTAL_COLOR = '#2e8b57'

def check_disk_space(path, min_space_gb=1):
    """Verifica se há espaço suficiente em disco"""
    usage = shutil.disk_usage(os.path.dirname(path))
    return usage.free > (min_space_gb * 1024**3)

def clean_matplotlib_memory():
    """Limpa a memória do matplotlib"""
    plt.close('all')
    import gc
    gc.collect()

def format_currency(value):
    """Formata um valor como moeda brasileira (R$)"""
    return f"R${value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def clean_currency(value):
    """
    Converte valores monetários em formato string para float, tratando:
    - Positivos: R$ 1.234,56 ou 1.234,56 ou 1234.56
    - Negativos: R$ (1.234,56) ou (1.234,56) ou -1.234,56 ou R$ -1.234,56
    """
    if pd.isna(value):
        return value
        
    if isinstance(value, (int, float)):
        return float(value)
        
    if isinstance(value, str):
        original = value
        # Padroniza o formato removendo R$, espaços e caracteres não numéricos exceto ,.-
        value = value.replace('R$', '').strip()
        
        # Verifica se é negativo (diferentes formatos)
        negative = (value.startswith('-') or 
                   '(' in value or 
                   ')' in value or
                   value.startswith('(') and value.endswith(')'))
        
        # Remove todos os caracteres não numéricos exceto pontos e vírgulas
        cleaned = re.sub(r'[^\d,-.]', '', value)
        
        # Remove parênteses se existirem
        cleaned = cleaned.replace('(', '').replace(')', '')
        
        # Trata casos onde o sinal negativo está no meio (formato inválido)
        if '-' in cleaned[1:]:
            cleaned = cleaned.replace('-', '')
        
        # Determina o separador decimal
        if ',' in cleaned and '.' in cleaned:
            # Se tem ambos, assume que vírgula é decimal
            cleaned = cleaned.replace('.', '').replace(',', '.')
        elif ',' in cleaned:
            # Só tem vírgula, verifica se é decimal
            parts = cleaned.split(',')
            if len(parts) > 1 and len(parts[-1]) == 2:  # Assume valor monetário com centavos
                cleaned = cleaned.replace('.', '').replace(',', '.')
            else:
                cleaned = cleaned.replace(',', '.')
        
        try:
            num = float(cleaned)
            return -abs(num) if negative else abs(num)
        except ValueError:
            print(f"Valor não convertido: {original}")
            return None
    return value

def format_currency(value):
    """Formata um valor como moeda brasileira (R$) tratando negativos"""
    abs_value = abs(value)
    formatted = f"R${abs_value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{formatted}" if value < 0 else formatted

def create_report_directory(output_dir, month_name, year):
    """Cria o diretório para os relatórios se não existir"""
    dir_name = f"Relatório Analítico de Vendas - {month_name} {year}"
    report_dir = os.path.join(output_dir, dir_name)
    
    if not os.path.exists(report_dir):
        os.makedirs(report_dir)
    
    return report_dir

def generate_report(file_path, sheet_name, output_dir, metric_column, metric_name, unit, items_per_page=5):
    """Gera um relatório PDF para uma métrica específica"""
    try:
        # Ler os dados primeiro para obter as datas
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
            ['CODPRODUTO', 'DESCRICAO', metric_column, 'DATA']]
        
        # Converter DATA para datetime e obter mês/ano
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        
        # Obter nome do mês em português
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        
        # Criar diretório para os relatórios
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        
        # Criar nome do arquivo com as variáveis
        output_filename = f"Relatório Analítico de Vendas - {metric_name} - {nome_mes} {primeiro_ano} - {items_per_page} em {items_per_page}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        temp_path = os.path.join(report_dir, f"temp_{output_filename}")

        print(f"Gerando - {output_filename}")
        
        # ETAPA 1: Tratamento de valores monetários
        if metric_column == 'Fat Liquido':
            df[metric_column] = df[metric_column].apply(clean_currency)
            df = df[df[metric_column].notna()]
        
        # ETAPA 2: Filtragem e tratamento específico por métrica
        df = df[df[metric_column].notna()]
        
        # Cálculo do TOTAL corrigido para cada métrica
        if metric_name == 'Tonelagem':
            # Para tonelagem, somamos os absolutos dos valores (não subtraímos negativos)
            total_metric = df[metric_column].sum()
        elif metric_name == 'Margem':
            # Para margem, mantemos o cálculo original (já está em %)
            total_metric = df[metric_column].sum() * 100  # Média das margens em %
        else:
            # Para faturamento, soma simples
            total_metric = df[metric_column].sum()
        
        # Formatar o total corretamente
        if metric_name == 'Faturamento':
            total_text = format_currency(total_metric)
        elif metric_name == 'Margem':
            total_text = f"{total_metric:.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        else:  # Tonelagem
            total_text = f"{total_metric:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # ETAPA 3: Processamento do ranking - tratamento específico por métrica
        if metric_name == 'Tonelagem':
            df_ranking = df.copy()
            # Não aplicamos mais abs() aqui para manter os valores reais
            grouped = df_ranking.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            
            # Verificar se há valores negativos para o gráfico de pizza
            if (grouped[metric_column] < 0).any():
                warnings.warn(f"Não é possível criar gráfico em pizza para {metric_name} com valores negativos. "
                             f"Os valores negativos serão exibidos nas tabelas e outros gráficos, "
                             f"mas o gráfico de pizza será omitido.")
        elif metric_name == 'Margem':
            grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            grouped[metric_column] = grouped[metric_column] * 100  # Convertemos para %
        else:
            # Para faturamento, soma simples
            grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
        
        # Restante do código permanece igual...
        latest_descriptions = df.sort_values('DATA', ascending=False).drop_duplicates('CODPRODUTO')[['CODPRODUTO', 'DESCRICAO']]
        grouped = pd.merge(grouped, latest_descriptions, on='CODPRODUTO', how='left')
        
        # Ordenação e preparação do DataFrame final
        sorted_df = grouped.sort_values(metric_column, ascending=False).reset_index(drop=True)
        sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
        
        # Processar dados para série temporal
        time_series = df.copy()
        if metric_name == 'Tonelagem':
            time_series[metric_column] = time_series[metric_column].abs()  # Valores absolutos para gráficos
        elif metric_name == 'Margem':
            time_series[metric_column] = time_series[metric_column] * 100  # Converter para %

        time_series = df.copy()
        time_series['SEMANA'] = time_series['DATA'].dt.to_period('W').dt.start_time
        time_series = time_series.groupby(['CODPRODUTO', 'SEMANA'])[metric_column].sum().reset_index()
        
         # Criar PDF temporário primeiro
        with PdfPages(temp_path) as pdf:
            # Página de título - MODIFICADA PARA INCLUIR O TOTAL
            fig_title = plt.figure(figsize=(11, 16))
            plt.text(0.5, 0.5, f"RELATÓRIO ANALÍTICO DE VENDAS - {metric_name.upper()}", 
                    fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                    fontsize=18, ha='center', va='center', fontweight='normal')
            plt.text(0.5, 0.4, f"{metric_name.upper()} (TOTAL): {total_text}", 
                    fontsize=16, ha='center', va='center', fontweight='bold', color=TOTAL_COLOR)
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight')
            plt.close(fig_title)
            
            # Páginas de conteúdo
            for i in range(0, len(sorted_df), items_per_page):
                chunk = sorted_df.iloc[i:i+items_per_page]
                produtos_na_pagina = chunk['CODPRODUTO'].tolist()
                
                fig = plt.figure(figsize=(11, 16))
                gs = fig.add_gridspec(4, 1)
                ax1 = fig.add_subplot(gs[0])
                ax2 = fig.add_subplot(gs[1])
                ax3 = fig.add_subplot(gs[2])
                ax4 = fig.add_subplot(gs[3])
                
                fig.suptitle(f'Ranking de Produtos {i+1}-{min(i+items_per_page, len(sorted_df))}', 
                            fontsize=14, y=1.02)
                
                # Tabela
                ax1.axis('off')
                
                # Formatar os valores para exibição na tabela
                display_values = chunk[metric_column].copy()
                if metric_name == 'Faturamento':
                    display_values = display_values.apply(format_currency)
                elif metric_name == 'Margem':
                    display_values = display_values.apply(lambda x: f"{x:.2f}%")
                else:  # Tonelagem
                    display_values = display_values.apply(lambda x: f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                table_data = chunk[['Posição', 'CODPRODUTO', 'DESCRICAO']].copy()
                table_data[metric_column] = display_values
                table = ax1.table(
                    cellText=table_data.values,
                    colLabels=['Posição', 'Código', 'Descrição', f'{metric_name} ({unit})'],
                    loc='center',
                    cellLoc='center',
                    colWidths=[0.1, 0.1, 0.5, 0.3]
                )
                table.auto_set_font_size(False)
                table.set_fontsize(8)
                table.scale(1, 1.3)
                
                # Gráfico de Pizza 
                ax2.set_title('Distribuição Percentual', fontsize=10, pad=10)
                colormap_name = 'jet'
                num_produtos = len(produtos_na_pagina)
                colors = plt.colormaps[colormap_name].resampled(num_produtos)

                # Verificar se há valores negativos ou soma zero
                if (chunk[metric_column] < 0).any():
                    ax2.text(0.5, 0.5, 'Gráfico não disponível:\nvalores negativos presentes', 
                            ha='center', va='center', fontsize=10)
                    ax2.axis('off')
                elif chunk[metric_column].sum() <= 0:
                    ax2.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico', 
                            ha='center', va='center', fontsize=10)
                    ax2.axis('off')
                else:
                    if metric_name == 'Faturamento':
                        autopct_format = lambda p: f'{p:.1f}%\n({format_currency(p*sum(chunk[metric_column])/100)})'
                    elif metric_name == 'Margem':
                        autopct_format = lambda p: f'{p:.1f}%\n({p*sum(chunk[metric_column])/100:.2f}%)'
                    else:
                        autopct_format = lambda p: f'{p:.1f}%\n({p*sum(chunk[metric_column])/100:,.1f} {unit})'.replace(",", "X").replace(".", ",").replace("X", ".")

                    wedges, texts, autotexts = ax2.pie(
                        chunk[metric_column],
                        autopct=autopct_format,
                        startangle=140,
                        textprops={'fontsize': 7, 'color': 'white'},
                        wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                        pctdistance=0.85,
                        colors=[colors(i) for i in range(num_produtos)]
                    )

                    for text, wedge in zip(autotexts, wedges):
                        wedge_color = wedge.get_facecolor()
                        text.set_color('white')
                        text.set_path_effects([
                            patheffects.withStroke(linewidth=2, foreground=wedge_color),
                            patheffects.Normal()
                        ])

                    n_cols = min(4, len(chunk))
                    ax2.legend(wedges, chunk['DESCRICAO'],
                              loc="upper center",
                              bbox_to_anchor=(0.5, -0.05),
                              ncol=n_cols,
                              fontsize=7,
                              title_fontsize=8,
                              frameon=False)
                
                # Gráfico de Linha
                ts_filtered = time_series[time_series['CODPRODUTO'].isin(produtos_na_pagina)]
                ts_filtered = pd.merge(ts_filtered, latest_descriptions, on='CODPRODUTO', how='left')
                
                ax3.set_title('Evolução Temporal (por semana)', fontsize=10, pad=10)
                ax3.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                
                lines = []
                labels = []
                product_colors = {prod: colors(i) for i, prod in enumerate(produtos_na_pagina)}
                
                for produto, group in ts_filtered.groupby('CODPRODUTO'):
                    group = group.sort_values('SEMANA')
                    line_color = product_colors[produto]

                    line, = ax3.plot(group['SEMANA'], group[metric_column], 
                                   marker='o', linestyle='-', 
                                   color=line_color,
                                   markersize=4, linewidth=1.5)
                    lines.append(line)
                    labels.append(group['DESCRICAO'].iloc[0])

                    for x, y in zip(group['SEMANA'], group[metric_column]):
                        if metric_name == 'Faturamento':
                            label = format_currency(y)
                        elif metric_name == 'Margem':
                            label = f'{y:.2f}%'
                        else:
                            label = f'{y:,.1f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                        
                        # Na função generate_report(), modificar a parte das anotações do gráfico de linhas:
                        annotation = ax3.annotate(label, 
                            xy=(x, y),
                            xytext=(0, 5),
                            textcoords='offset points',
                            ha='center', va='bottom',
                            fontsize=6,
                            color="white")  # Alterado para verde
                        annotation.set_path_effects([
                            patheffects.withStroke(linewidth=2, foreground=line_color),
                            patheffects.Normal()
                        ])
                
                ax3.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d/%m'))
                ax3.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(byweekday=plt.matplotlib.dates.MO))
                
                tick_labels = []
                for semana in ax3.get_xticks():
                    semana_inicio = plt.matplotlib.dates.num2date(semana)
                    semana_fim = semana_inicio + timedelta(days=6)
                    tick_labels.append(f"{semana_inicio.strftime('%d/%m')}\na\n{semana_fim.strftime('%d/%m')}")
                
                ax3.set_xticklabels(tick_labels)
                
                n_cols = min(4, len(lines))
                ax3.legend(lines, labels,
                          loc="upper center",
                          bbox_to_anchor=(0.5, -0.2),
                          ncol=n_cols,
                          fontsize=7,
                          frameon=False)
                
                ax3.grid(True, linestyle=':', alpha=0.5)
                plt.setp(ax3.get_xticklabels(), rotation=30, ha='right', fontsize=7)
                plt.setp(ax3.get_yticklabels(), fontsize=7)
                
                # Gráfico de Barras
                ax4.set_title(f'{metric_name} por Produto', fontsize=10, pad=10)
                ax4.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                
                bars = ax4.bar(
                    chunk['DESCRICAO'],
                    chunk[metric_column],
                    color=[product_colors[p] for p in produtos_na_pagina]
                )
                
                for bar in bars:
                    height = bar.get_height()
                    bar_color = bar.get_facecolor()
                    
                    if metric_name == 'Faturamento':
                        label = format_currency(height)
                    elif metric_name == 'Margem':
                        label = f'{height:.2f}%'
                    else:
                        label = f'{height:,.1f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                    
                    annotation = ax4.annotate(label,
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom',
                        fontsize=8,
                        color="white") 
                    annotation.set_path_effects([
                        patheffects.withStroke(linewidth=2, foreground=bar_color),
                        patheffects.Normal()
                    ])
                
                plt.setp(ax4.get_xticklabels(), rotation=15, ha='right', fontsize=8)
                plt.setp(ax4.get_yticklabels(), fontsize=7)
                ax4.grid(True, axis='y', linestyle=':', alpha=0.5)
                
                plt.tight_layout(rect=[0, 0, 1, 0.95])
                pdf.savefig(fig, dpi=150, bbox_inches='tight', pad_inches=0.5)
                plt.close(fig)
                
        # Renomear arquivo temporário para final
        os.rename(temp_path, output_path)
        print(f"Finalizado - {output_filename}")

    except Exception as e:
        print(f"ERRO ao gerar relatório de {metric_name}: {str(e)}")
        for path in [output_path, temp_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

def generate_general_report(file_path, sheet_name, output_dir):
    """Gera um relatório geral com estatísticas básicas e gráficos comparativos"""
    try:
        # Ler os dados do Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)
        
        # Verificar se as colunas necessárias existem
        required_columns = ['RAZAO', 'VENDEDOR', 'CODPRODUTO', 'DATA', 'QTDE REAL', 'Fat Liquido', 'Margem']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Colunas faltando no arquivo Excel: {', '.join(missing_columns)}")
        
        # Obter mês/ano para o nome do arquivo
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        
        # Criar diretório para os relatórios
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        
        # Criar nome do arquivo
        output_filename = f"Relatório Analítico de Vendas - Geral - {nome_mes} {primeiro_ano}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        
        print(f"Gerando - {output_filename}")
        
        # Verificar espaço em disco
        if not check_disk_space(output_path):
            raise RuntimeError("Espaço insuficiente em disco para gerar o relatório")

        # Verificar e remover arquivo existente
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except PermissionError:
                raise RuntimeError(f"Feche o arquivo {output_path} antes de executar")
        
        # Calcular as métricas básicas
        qtde_clientes = df['RAZAO'].nunique()
        qtde_vendedores = df['VENDEDOR'].nunique()
        qtde_produtos = df['CODPRODUTO'].nunique()
        
        # Criar dados para a tabela (sem cabeçalhos)
        table_data = [
            ['Qtde Clientes', qtde_clientes],
            ['Qtde Vendedores', qtde_vendedores],
            ['Qtde Produtos', qtde_produtos]
        ]
        
        # Preparar dados para os gráficos de pizza
        def prepare_pie_data(metric_column, metric_name):
            group_cols = ['CODPRODUTO']
            if 'DESCRICAO' in df.columns:
                group_cols.append('DESCRICAO')
            
            grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            
            if metric_name == 'Margem':
                grouped[metric_column] = grouped[metric_column] * 100
            
            sorted_df = grouped.sort_values(metric_column, ascending=False).reset_index(drop=True)
            top20 = sorted_df.head(20)
            resto = sorted_df.iloc[20:]
            
            total_top20 = top20[metric_column].sum()
            total_resto = resto[metric_column].sum()
            
            return {
                'labels': ['Top 20', 'Resto'],
                'values': [total_top20, total_resto],
                'title': f'Top 20 produtos vs Resto - {metric_name}',
                'total': total_top20 + total_resto
            }
        
        pie_data = {
            'Tonelagem': prepare_pie_data('QTDE REAL', 'Tonelagem'),
            'Faturamento': prepare_pie_data('Fat Liquido', 'Faturamento'),
            'Margem': prepare_pie_data('Margem', 'Margem')
        }
       
        # Criar PDF
        with PdfPages(output_path) as pdf:
            # [Página de título permanece a mesma]
            
           # Página com tabela e gráficos
            fig_content = plt.figure(figsize=(11, 16))
            gs = fig_content.add_gridspec(4, 1, height_ratios=[1, 1, 1, 1], hspace=0.5)
            
            # Tabela na primeira parte
            ax_table = fig_content.add_subplot(gs[0])
            ax_table.axis('off')
            
            table = ax_table.table(
                cellText=table_data,
                loc='center',
                cellLoc='center',
                colWidths=[0.5, 0.5]
            )
            
            table.auto_set_font_size(False)
            table.set_fontsize(12)
            table.scale(1, 2)
            ax_table.set_title("Estatísticas Gerais de Vendas", fontsize=14, pad=20)
            
            def format_value(value, metric_name):
                if metric_name == 'Faturamento':
                    return format_currency(value)
                elif metric_name == 'Margem':
                    return f"{value:.2f}%"
                else:
                    return f"{value:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            colors = ['#1f77b4', '#ff7f0e']
            legend_labels = ['Top 20 Produtos', 'Resto dos Produtos']
            
            for i, (metric_name, data) in enumerate(pie_data.items()):
                ax_pie = fig_content.add_subplot(gs[i+1])
                
                # Ajustar posição para dar mais espaço
                ax_pie.set_position([0.1, ax_pie.get_position().y0, 0.8, ax_pie.get_position().height * 0.8])
                
                # Título principal
                ax_pie.set_title(data['title'], fontsize=12, pad=25, y=1.08)
                
                # Total (verde)
                ax_pie.text(0.5, 1.02, f"Total: {format_value(data['total'], metric_name)}", 
                           fontsize=11, ha='center', va='bottom', 
                           color=TOTAL_COLOR, transform=ax_pie.transAxes,
                           bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=3))
                
                # Gráfico de pizza (sem autopct)
                wedges, texts = ax_pie.pie(
                    data['values'],
                    startangle=90,
                    colors=colors,
                    wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                    radius=0.8
                )
                
                # Adicionar os valores fora do gráfico
                bbox_props = dict(boxstyle="round,pad=0.3", fc="white", ec="gray", lw=0.5, alpha=0.8)
                kw = dict(arrowprops=dict(arrowstyle="-"), bbox=bbox_props, zorder=0, va="center")
                
                for j, (wedge, value) in enumerate(zip(wedges, data['values'])):
                    ang = (wedge.theta2 - wedge.theta1)/2. + wedge.theta1
                    y = np.sin(np.deg2rad(ang))
                    x = np.cos(np.deg2rad(ang))
                    
                    # Calcular porcentagem
                    percentage = 100 * value / data['total']
                    formatted_value = format_value(value, metric_name)
                    
                    # Texto com porcentagem e valor
                    text = f"{percentage:.1f}%\n({formatted_value})"
                    
                    # Posicionar fora do gráfico
                    horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
                    connectionstyle = f"angle,angleA=0,angleB={ang}"
                    kw["arrowprops"].update({"connectionstyle": connectionstyle})
                    
                    ax_pie.annotate(text, xy=(x, y), xytext=(1.3*np.sign(x), 1.3*y),
                                   horizontalalignment=horizontalalignment,
                                   fontsize=9, **kw)
                
                # Legenda
                ax_pie.legend(wedges, legend_labels,
                             loc='lower center',
                             bbox_to_anchor=(0.5, -0.3),
                             ncol=2,
                             fontsize=10,
                             frameon=False)
            
            plt.tight_layout()
            pdf.savefig(fig_content, bbox_inches='tight')
            plt.close(fig_content)
            
        print(f"Finalizado - {output_filename}")
        
    except Exception as e:
        print(f"ERRO ao gerar relatório geral: {str(e)}")
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
            
# Configuração principal
file_path = r"C:\Users\win11\OneDrive\Documentos\Margens de fechamento\Margem_250630 - FEC - wapp V5.xlsx"
sheet_name = "Base (3,5%)"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
items_per_page = 5

# Definir as métricas que queremos analisar
metrics = [
    {'column': 'QTDE REAL', 'name': 'Tonelagem', 'unit': 'kg'},
    {'column': 'Fat Liquido', 'name': 'Faturamento', 'unit': 'R$'},
    {'column': 'Margem', 'name': 'Margem', 'unit': '%'}
]


# Gerar relatório geral primeiro
generate_general_report(
    file_path=file_path,
    sheet_name=sheet_name,
    output_dir=output_dir
)

# Gerar todos os relatórios específicos
for metric in metrics:
    generate_report(
        file_path=file_path,
        sheet_name=sheet_name,
        output_dir=output_dir,
        metric_column=metric['column'],
        metric_name=metric['name'],
        unit=metric['unit'],
        items_per_page=items_per_page
    )