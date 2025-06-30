import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import shutil
from datetime import timedelta
from matplotlib import patheffects
import re

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

def generate_report(file_path, sheet_name, output_dir, metric_column, metric_name, unit, items_per_page=5):
    """Gera um relatório PDF para uma métrica específica"""
    try:
        # Ler os dados primeiro para obter as datas
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
            ['CODPRODUTO', 'DESCRICAO', metric_column, 'DATA']]
        
       # ETAPA 1: Tratamento de valores monetários
        if metric_column == 'Fat Liquido':
            df[metric_column] = df[metric_column].apply(clean_currency)
            df = df[df[metric_column].notna()]
            
            print("\nValores negativos APÓS limpeza:")
            print(df[df[metric_column] < 0].head())

        # ETAPA 2: Filtragem - APENAS para Tonelagem removemos negativos
        df = df[df[metric_column].notna()]
        if metric_name == 'Tonelagem':
            df = df[df[metric_column] >= 0]
        
        # ETAPA 3: Processamento do ranking
        latest_descriptions = df.sort_values('DATA', ascending=False).drop_duplicates('CODPRODUTO')[['CODPRODUTO', 'DESCRICAO']]
        grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
        
        print(f"\nValores negativos NO AGRUPAMENTO:")
        print(grouped[grouped[metric_column] < 0].head())
        
        # Converter Margem para porcentagem (se estava em decimal)
        if metric_name == 'Margem':
            df[metric_column] = df[metric_column] * 100
        
        # Converter DATA para datetime e obter mês/ano
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        
        # Obter nome do mês em português
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        
        # Criar nome do arquivo com as variáveis
        output_filename = f"Relatório Analítico de Vendas - {metric_name} - {nome_mes} {primeiro_ano} - {items_per_page} em {items_per_page}.pdf"
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
        # Primeiro, encontre a descrição mais recente para cada produto
        latest_descriptions = df.sort_values('DATA', ascending=False).drop_duplicates('CODPRODUTO')[['CODPRODUTO', 'DESCRICAO']]
        
       # Processar dados para o ranking
        grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
        
        # Adicionar verificação explícita para negativos
        print(f"\nVerificação pós-agrupamento - Valores negativos existem: {any(grouped[metric_column] < 0)}")
        
        # Aplicar formatação específica para cada métrica
        if metric_name == 'Faturamento':
            grouped[metric_column] = grouped[metric_column].round(2)
        elif metric_name == 'Margem':
            grouped[metric_column] = grouped[metric_column].round(2)  # 2 casas decimais para Margem
        else:  # Tonelagem
            grouped[metric_column] = grouped[metric_column].round(3)
        
        # Junte com as descrições mais recentes
        grouped = pd.merge(grouped, latest_descriptions, on='CODPRODUTO', how='left')
        
        # Ordene e prepare o DataFrame final
        sorted_df = grouped.sort_values(metric_column, ascending=False).reset_index(drop=True)
        sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
        
        # Processar dados para série temporal - Agrupando por semana corretamente
        time_series = df.copy()
        time_series['SEMANA'] = time_series['DATA'].dt.to_period('W').dt.start_time
        time_series = time_series.groupby(['CODPRODUTO', 'SEMANA'])[metric_column].sum().reset_index()
        
        # Criar PDF temporário primeiro
        with PdfPages(temp_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16))
            plt.text(0.5, 0.5, f"RELATÓRIO ANALÍTICO DE VENDAS - {metric_name.upper()}", 
                     fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                     fontsize=18, ha='center', va='center', fontweight='normal')
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight')
            plt.close(fig_title)
            
            # Páginas de conteúdo
            # ETAPA 4: Verificação durante a geração de páginas
            for i in range(0, len(sorted_df), items_per_page):
                chunk = sorted_df.iloc[i:i+items_per_page]
            
                if any(chunk[metric_column] < 0):
                    print(f"\nPágina {i//items_per_page + 1} contém negativos:")
                    print(chunk[chunk[metric_column] < 0])
            
                produtos_na_pagina = chunk['CODPRODUTO'].tolist()
                
                fig = plt.figure(figsize=(11, 16))  # Removido constrained_layout
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
                
                # Verificar se há valores válidos para o gráfico de pizza
                if chunk[metric_column].sum() > 0:
                    # Formatar rótulos de porcentagem de acordo com a métrica
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
                else:
                    ax2.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico', 
                            ha='center', va='center', fontsize=10)
                    ax2.axis('off')
                
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

                    # Formatar anotações de acordo com a métrica
                    for x, y in zip(group['SEMANA'], group[metric_column]):
                        if metric_name == 'Faturamento':
                            label = format_currency(y)
                        elif metric_name == 'Margem':
                            label = f'{y:.2f}%'
                        else:
                            label = f'{y:,.1f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                        
                        annotation = ax3.annotate(label, 
                                                xy=(x, y),
                                                xytext=(0, 5),
                                                textcoords='offset points',
                                                ha='center', va='bottom',
                                                fontsize=6,
                                                color='white')
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
                    
                    # Formatar anotações de acordo com a métrica
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
                                            color='white')
                    annotation.set_path_effects([
                        patheffects.withStroke(linewidth=2, foreground=bar_color),
                        patheffects.Normal()
                    ])
                
                plt.setp(ax4.get_xticklabels(), rotation=15, ha='right', fontsize=8)
                plt.setp(ax4.get_yticklabels(), fontsize=7)
                ax4.grid(True, axis='y', linestyle=':', alpha=0.5)
                
                # Ajustar layout manualmente
                plt.tight_layout(rect=[0, 0, 1, 0.95])  # Ajuste para o suptitle
                
                # Configurações do PDF
                pdf.savefig(fig, dpi=150, bbox_inches='tight', pad_inches=0.5)
                plt.close(fig)
                
        # Renomear arquivo temporário para final
        os.rename(temp_path, output_path)
        print(f"Relatório de {metric_name} gerado com sucesso em: {output_path}")
        print(f"Tamanho do arquivo: {os.path.getsize(output_path)/1024/1024:.2f} MB")
        # Verificar se há valores negativos após o processamento
        negatives = df[df[metric_column] < 0]
        if not negatives.empty:
            print(f"\nProdutos com valores negativos que serão incluídos:")
            print(negatives[['CODPRODUTO', 'DESCRICAO', metric_column]].head())
        else:
            print("\nNenhum valor negativo encontrado após processamento")
        
        # Verificar valores extremos
        print("\nValores extremos:")
        print(f"Top 5 positivos: {df.nlargest(5, metric_column)[[metric_column, 'DESCRICAO']]}")
        print(f"Top 5 negativos: {df.nsmallest(5, metric_column)[[metric_column, 'DESCRICAO']]}")

    except Exception as e:
        print(f"ERRO ao gerar relatório de {metric_name}: {str(e)}")
        for path in [output_path, temp_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

# Configuração principal
file_path = r"C:\Users\win11\Documents\Andrey Enviou\Margem_250531 - wapp - V3.xlsx"
sheet_name = "Base (3,5%)"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
items_per_page = 5

# Definir as métricas que queremos analisar
metrics = [
    {'column': 'QTDE REAL', 'name': 'Tonelagem', 'unit': 'kg'},
    {'column': 'Fat Liquido', 'name': 'Faturamento', 'unit': 'R$'},
    {'column': 'Margem', 'name': 'Margem', 'unit': '%'}
]


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
        
        # Criar nome do arquivo
        output_filename = f"Relatório Analítico de Vendas - Geral - {nome_mes} {primeiro_ano}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        
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
            # Processar dados para o ranking
            # Se não houver coluna DESCRICAO, usar apenas CODPRODUTO
            group_cols = ['CODPRODUTO']
            if 'DESCRICAO' in df.columns:
                group_cols.append('DESCRICAO')
            
            grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            
            # Converter Margem para porcentagem se necessário
            if metric_name == 'Margem':
                grouped[metric_column] = grouped[metric_column] * 100
            
            # Ordenar e pegar top 20
            sorted_df = grouped.sort_values(metric_column, ascending=False).reset_index(drop=True)
            top20 = sorted_df.head(20)
            resto = sorted_df.iloc[20:]
            
            # Calcular totais
            total_top20 = top20[metric_column].sum()
            total_resto = resto[metric_column].sum()
            
            return {
                'labels': ['Top 20', 'Resto'],
                'values': [total_top20, total_resto],
                'title': f'Top 20 produtos vs Resto - {metric_name}',
                'total': total_top20 + total_resto
            }
        
        # Preparar dados para cada métrica
        pie_data = {
            'Tonelagem': prepare_pie_data('QTDE REAL', 'Tonelagem'),
            'Faturamento': prepare_pie_data('Fat Liquido', 'Faturamento'),
            'Margem': prepare_pie_data('Margem', 'Margem')
        }
        
        # Criar PDF
        with PdfPages(output_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16))
            plt.text(0.5, 0.5, "RELATÓRIO ANALÍTICO DE VENDAS - GERAL", 
                     fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                     fontsize=18, ha='center', va='center', fontweight='normal')
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight')
            plt.close(fig_title)
            
            # Página com tabela e gráficos
            fig_content = plt.figure(figsize=(11, 16))
            
            # Criar grid para organizar os elementos
            gs = fig_content.add_gridspec(4, 1, height_ratios=[1, 1, 1, 1])
            
            # Tabela na primeira parte
            ax_table = fig_content.add_subplot(gs[0])
            ax_table.axis('off')
            
            # Criar tabela sem cabeçalhos
            table = ax_table.table(
                cellText=table_data,
                loc='center',
                cellLoc='center',
                colWidths=[0.5, 0.5]
            )
            
            # Formatar tabela
            table.auto_set_font_size(False)
            table.set_fontsize(12)
            table.scale(1, 2)
            
            # Adicionar título à tabela
            ax_table.set_title("Estatísticas Gerais de Vendas", fontsize=14, pad=20)
            
            # Função para formatar valores
            def format_value(value, metric_name):
                if metric_name == 'Faturamento':
                    return format_currency(value)
                elif metric_name == 'Margem':
                    return f"{value:.2f}%"
                else:  # Tonelagem
                    return f"{value:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            # Cores para os gráficos
            colors = ['#1f77b4', '#ff7f0e']
            legend_labels = ['Top 20 Produtos', 'Resto dos Produtos']
            
            # Criar os três gráficos de pizza
            for i, (metric_name, data) in enumerate(pie_data.items()):
                ax_pie = fig_content.add_subplot(gs[i+1])
                
                # Plotar o gráfico de pizza
                wedges, texts, autotexts = ax_pie.pie(
                    data['values'],
                    autopct=lambda p: f'{p:.1f}%\n({format_value(p * data["total"] / 100, metric_name)})',
                    startangle=90,
                    colors=colors,
                    textprops={'fontsize': 9, 'color': 'white'},
                    wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                    pctdistance=0.85
                )
                
                # Ajustar formatação dos textos
                for text, wedge in zip(autotexts, wedges):
                    text.set_color('white')
                    text.set_path_effects([
                        patheffects.withStroke(linewidth=2, foreground=wedge.get_facecolor()),
                        patheffects.Normal()
                    ])
                
                # Adicionar título
                ax_pie.set_title(f"{data['title']} - {nome_mes} {primeiro_ano}", fontsize=12, pad=10)
                
                # Adicionar legenda abaixo do gráfico
                ax_pie.legend(wedges, legend_labels,
                             loc='lower center',
                             bbox_to_anchor=(0.5, -0.2),
                             ncol=2,
                             fontsize=10,
                             frameon=False)
            
            # Ajustar layout
            plt.tight_layout()
            
            # Salvar página
            pdf.savefig(fig_content, bbox_inches='tight')
            plt.close(fig_content)
            
        print(f"Relatório Geral gerado com sucesso em: {output_path}")
        
    except Exception as e:
        print(f"ERRO ao gerar relatório geral: {str(e)}")
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
            
# Adicione esta chamada após o loop que gera os outros relatórios
generate_general_report(
    file_path=file_path,
    sheet_name=sheet_name,
    output_dir=output_dir
)

# Gerar todos os relatórios
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
