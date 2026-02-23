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
    """Limpa a memória do matplotlib de forma mais agressiva"""
    plt.close('all')
    import gc
    gc.collect()
    # Forçar coleta de lixo
    gc.collect(generation=2)
    import matplotlib
    matplotlib.pyplot.close('all')

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
    dir_name = f"Ranking de Vendas - {month_name} {year}"
    report_dir = os.path.join(output_dir, dir_name)
    
    if not os.path.exists(report_dir):
        os.makedirs(report_dir)
    
    return report_dir

def generate_report(file_path, sheet_name, output_dir, metric_column, metric_name, unit, items_per_page=5):
    """Gera um relatório PDF para uma métrica específica, agrupando produtos conforme definido"""
    temp_path = None
    output_path = None
    
    try:
        # 1. Limpar memória antes de começar
        clean_matplotlib_memory()
        
        # 2. Definir os grupos de produtos (todos em maiúsculas)
        product_groups = {
            'ACEM': [1924, 8006, 1940, 1878, 8101, 1841],
            'ALCATRA C/ MAMINHA': [8001, 1836, 1965, 1800],
            'BARRIGA': [1833, 1639, 1544, 1674, 1863, 1845, 1385, 1898, 1913, 1513, 
                        1444, 1434, 1960, 1954, 5200, 2042, 2043, 2047, 2051],
            'BUCHO': [1567, 1816, 1856, 1480, 1527, 1903, 1855, 1958],
            'BACON MANTA': [869, 981],
            'BANHA SUINA': [1605, 1139],
            'BATATA': [1767, 1872],
            'BOLACHA': [1649, 1644, 1643, 1647, 1645, 1648],
            'BOLINHO': [1709, 1707, 1708, 1941, 1999],
            'CARNE TEMPERADA': [1720, 1623, 1618],
            'CARRE': [1568, 1355, 1443, 1817, 1464, 1640, 1533, 1286, 1518, 1653, 1216, 
                      1316, 906, 1210, 1908, 1221, 1177, 1612, 1634, 917, 1689, 1511,
                      1955],
            'CONTRA FILÉ': [1901, 1922, 1840, 1947, 1894, 1899, 1905, 1503, 1824],
            'CORAÇÃO DE ALCATRA': [1830, 1939],
            'COSTELA BOV': [1768, 1825, 1931, 1814, 1890, 1890],
            'COSTELA MINGA': [1973, 1982],
            'COSTELA SUINA CONGELADA': [1478, 1595, 1506, 1081, 1592, 1412, 1641, 1888, 
                                        1522, 1638, 1607, 1517, 1461, 1416, 1760, 1877, 
                                        1664, 1053, 1314, 1617, 1599, 1896, 1857, 1179,
                                        1324, 1529, 1421, 1323, 1879, 1052, 1051, 1354,
                                        905,  1384, 1086, 1174, 1150, 1758, 1320, 1829,
                                        1665, 1327, 1442, 1431, 1704, 1736, 1445, 1321,
                                        1884, 1535, 8007, 2050],
            'COXÃO DURO': [1920, 8003, 1803, 1949, 1795],
            'COXÃO MOLE': [1831, 8002, 1948, 1976, 1375],
            'COXINHA DA ASA': [1604, 1546, 8005, 1722, 2038],
            'CUPIM A': [1772],
            'CUPIM B': [1804, 1456, 1926, 1984],
            'FIGADO': [1808, 1455, 1818, 1910, 1823, 1537, 1505, 1408, 1373, 1458, 1508,
                       1525, 1454, 1801, 1528, 1530, 1502, 1945, 1967, 1998, 1978, 1983,
                       2018, 2035, 2026],
            'FILÉ MIGNON': [1812, 1919],
            'FRALDA': [1797, 1925],
            'HAMBURGUER': [1009, 1866, 1010],
            'HOT POCKET': [1987],
            'JERKED': [1893, 1943, 1880, 1886, 1851],
            'LAGARTO': [1849, 1396, 1895, 1813],
            'LASANHA': [1003, 1691, 1997, 1002, 1991],
            'LINGUIÇA CALABRESA AURORA': [788, 1974],
            'LINGUIÇA CALABRESA SADIA': [1339, 807, 1848, 1847],
            'LINGUIÇA CALABRESA PAMPLONA': [9165, 910],
            'MEIO DA ASA': [2311, 1937, 2014, 2039, 2052],
            'MINI CHICKEN': [1024, 1994],
            'MINI LASANHA': [1992, 1985],
            'MOCOTÓ': [1539, 1460, 1342, 1540, 1675, 1850, 1827, 1821, 1853, 1407, 1723,
                       1585, 1407, 1723, 1585, 1843, 1584, 762, 1534, 1883, 1509, 1601,
                       1962, 2049, 2028],
            'MUSSARELA': [2000, 947, 1807, 1914],
            'NUGGETS': [1007, 1995],
            'PATINHO': [1805, 1874, 8000, 1938, 9166, 1966],
            'PALETA': [1953, 1964, 1923, 1975],
            'PEITO BOV': [1815, 1875, 1789, 1952],
            'PERNIL SUINO C/OSSO C/PELE': [1942, 1635, 1724, 1570, 1756, 1303],
            'PICANHA B': [1946, 1950],
            'PIZZA': [1989, 1990],
            'PURURUCA 60G': [1288, 1289, 1287],
            'RABO BOV': [1828, 1839, 1876, 1861, 1116, 1705, 1531, 1906, 1826, 1911, 1882,
                        1571, 1335, 1963, 1909, 1473, 1481],
            'SALAME UAI': [1495, 1500, 1496, 1497, 1498, 1499],
            'STEAK FGO': [1718, 1996],
            'TAPIOCA DA TERRINHA': [1929, 1930],
            'YOPRO': [1698, 1701, 1587, 1700, 86754, 1586, 9675]
        }
        
        # Inverter o dicionário para mapear código para nome do grupo
        code_to_group = {}
        for group_name, codes in product_groups.items():
            for code in codes:
                code_to_group[code] = group_name

        # Ler os dados primeiro para obter as datas
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
            ['CODPRODUTO', 'DESCRICAO', 'DATA', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']]
        
        # Adicionar coluna de grupo ao DataFrame
        df['GRUPO'] = df['CODPRODUTO'].map(code_to_group)
        
        # Converter DATA para datetime e obter mês/ano
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        
        # Obter nome do mês em português
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        
        # Criar diretório para os relatórios
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        
        # Definir a coluna métrica correta
        if metric_name == 'Tonelagem':
            metric_column = 'QTDE REAL'
        elif metric_name == 'Faturamento':
            metric_column = 'Fat Liquido'
        elif metric_name == 'Margem':
            metric_column = 'Margem Calculada'
        
        # Criar nome do arquivo
        output_filename = f"Ranking de Vendas - {metric_name} - {nome_mes} {primeiro_ano} - {items_per_page} em {items_per_page}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        temp_path = os.path.join(report_dir, f"temp_{output_filename}")

        print(f"Gerando - {output_filename}")
        
        # 3. ETAPA 1: Tratamento de valores monetários
        if metric_name in ['Faturamento', 'Margem']:
            df['Fat Liquido'] = df['Fat Liquido'].apply(clean_currency)
            df['Lucro / Prej.'] = df['Lucro / Prej.'].apply(clean_currency)
            # Remover linhas com valores NaN
            df = df[df['Fat Liquido'].notna() & df['Lucro / Prej.'].notna()]
        
        # 4. ETAPA 2: Preparação específica por métrica
        if metric_name == 'Tonelagem':
            df = df[df['QTDE REAL'].notna()]
        elif metric_name == 'Faturamento':
            df = df[df['Fat Liquido'].notna()]
            df['Fat Liquido'] = df['Fat Liquido'].astype(float)
        elif metric_name == 'Margem':
            # Calcular margem e zerar quando faturamento for negativo
            df['Margem Calculada'] = np.where(
                df['Fat Liquido'] <= 0, 0,
                (df['Lucro / Prej.'] / df['Fat Liquido']) * 100
            )
            df = df[df['Margem Calculada'].notna()]
        
        # *** CORREÇÃO PRINCIPAL: Criar uma única tabela com produtos e grupos ***
        # Para os grupos, vamos usar o nome do grupo como identificador único
        df_combined = df.copy()
        
        # Para produtos em grupos, substituir CODPRODUTO pelo nome do grupo
        df_combined['ID_AGRUPADO'] = df_combined.apply(
            lambda row: row['GRUPO'] if pd.notna(row['GRUPO']) else str(row['CODPRODUTO']), 
            axis=1
        )
        
        # Manter descrição original para produtos individuais, usar nome do grupo para grupos
        df_combined['DESCRICAO_AGRUPADA'] = df_combined.apply(
            lambda row: row['GRUPO'] if pd.notna(row['GRUPO']) else row['DESCRICAO'], 
            axis=1
        )
        
        # *** ETAPA 3: Agregar dados por ID_AGRUPADO (grupos e produtos individuais) ***
        if metric_name == 'Tonelagem':
            aggregated = df_combined.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
                'QTDE REAL': 'sum',
                'DATA': 'count'
            }).reset_index()
            aggregated.rename(columns={
                'QTDE REAL': metric_column,
                'DATA': 'Qtde de vendas'
            }, inplace=True)
            aggregated['CODPRODUTO'] = aggregated.apply(
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'],
                axis=1
            )
            
        elif metric_name == 'Faturamento':
            aggregated = df_combined.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
                'Fat Liquido': 'sum',
                'DATA': 'count'
            }).reset_index()
            aggregated.rename(columns={
                'Fat Liquido': metric_column,
                'DATA': 'Qtde de vendas'
            }, inplace=True)
            aggregated['CODPRODUTO'] = aggregated.apply(
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'],
                axis=1
            )
            
        elif metric_name == 'Margem':
            aggregated = df_combined.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
                'Lucro / Prej.': 'sum',
                'Fat Liquido': 'sum',
                'DATA': 'count'
            }).reset_index()
            aggregated.rename(columns={'DATA': 'Qtde de vendas'}, inplace=True)
            aggregated[metric_column] = np.where(
                aggregated['Fat Liquido'] <= 0, 0,
                (aggregated['Lucro / Prej.'] / aggregated['Fat Liquido']) * 100
            )
            aggregated['CODPRODUTO'] = aggregated.apply(
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'],
                axis=1
            )
        
        # Renomear colunas para consistência
        aggregated.rename(columns={
            'ID_AGRUPADO': 'GRUPO_ID',
            'DESCRICAO_AGRUPADA': 'DESCRICAO'
        }, inplace=True)
        
        # Ordenar e numerar as posições
        sorted_df = aggregated.sort_values(metric_column, ascending=False).reset_index(drop=True)
        sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
        
        # *** Cálculo do TOTAL (incluindo grupos) ***
        if metric_name == 'Tonelagem':
            total_metric = df['QTDE REAL'].sum()
        elif metric_name == 'Faturamento':
            total_metric = df['Fat Liquido'].sum()
        elif metric_name == 'Margem':
            total_lucro = df['Lucro / Prej.'].sum()
            total_fat = df['Fat Liquido'].sum()
            total_metric = 0 if total_fat <= 0 else (total_lucro / total_fat) * 100
        
        # Formatar o total
        if metric_name == 'Faturamento':
            total_text = format_currency(total_metric)
        elif metric_name == 'Margem':
            total_text = f"{total_metric:.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        else:  # Tonelagem
            total_text = f"{total_metric:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # *** ETAPA 4: Processar dados para série temporal (agora considerando grupos corretamente) ***
        time_series = df_combined.copy()
        time_series['SEMANA'] = time_series['DATA'].dt.to_period('W').dt.start_time
        
        # Usar ID_AGRUPADO para agrupar
        if metric_name == 'Tonelagem':
            time_series_agg = time_series.groupby(['ID_AGRUPADO', 'SEMANA'])['QTDE REAL'].sum().reset_index()
            time_series_agg.rename(columns={'QTDE REAL': metric_column}, inplace=True)
        elif metric_name == 'Faturamento':
            time_series_agg = time_series.groupby(['ID_AGRUPADO', 'SEMANA'])['Fat Liquido'].sum().reset_index()
            time_series_agg.rename(columns={'Fat Liquido': metric_column}, inplace=True)
        elif metric_name == 'Margem':
            time_series_agg = time_series.groupby(['ID_AGRUPADO', 'SEMANA']).agg({
                'Lucro / Prej.': 'sum',
                'Fat Liquido': 'sum'
            }).reset_index()
            time_series_agg[metric_column] = np.where(
                time_series_agg['Fat Liquido'] <= 0, 0,
                (time_series_agg['Lucro / Prej.'] / time_series_agg['Fat Liquido']) * 100
            )
        
        # Mapear IDs para descrições
        id_to_description = dict(zip(aggregated['GRUPO_ID'], aggregated['DESCRICAO']))
        time_series_agg['DESCRICAO'] = time_series_agg['ID_AGRUPADO'].map(id_to_description)
        
        # 5. OTIMIZAÇÃO: Criar PDF com menos uso de memória
        # Usar backend não-interativo para reduzir uso de memória
        plt.switch_backend('agg')
        
        # Criar PDF temporário
        with PdfPages(temp_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16), dpi=100)  # Reduzir DPI
            plt.text(0.5, 0.5, f"RANKING DE VENDAS - {metric_name.upper()}", 
                    fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                    fontsize=18, ha='center', va='center', fontweight='normal')
            plt.text(0.5, 0.4, f"{metric_name.upper()} (TOTAL): {total_text}", 
                    fontsize=16, ha='center', va='center', fontweight='bold', color=TOTAL_COLOR)
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight', dpi=100)
            plt.close(fig_title)
            clean_matplotlib_memory()
            
            # Páginas de conteúdo - processar em lotes menores
            total_pages = (len(sorted_df) + items_per_page - 1) // items_per_page
            
            for page_num in range(total_pages):
                i = page_num * items_per_page
                chunk = sorted_df.iloc[i:i+items_per_page]
                produtos_na_pagina = chunk['GRUPO_ID'].tolist()
                
                # Criar figura com tamanho otimizado
                fig = plt.figure(figsize=(11, 16), dpi=100)
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
                
                # *** CORREÇÃO: Mostrar "GRUPO" na coluna de código para grupos ***
                table_data = chunk[['Posição', 'CODPRODUTO', 'DESCRICAO', 'Qtde de vendas']].copy()
                table_data[metric_column] = display_values
                table = ax1.table(
                    cellText=table_data.values,
                    colLabels=['Posição', 'Tipo', 'Descrição', 'Qtde de vendas', f'{metric_name} ({unit})'],
                    loc='center',
                    cellLoc='center',
                    colWidths=[0.08, 0.08, 0.4, 0.2, 0.24]
                )
                table.auto_set_font_size(False)
                table.set_fontsize(8)
                table.scale(1, 1.3)
                
                # Gráfico de Pizza - simplificar
                ax2.set_title('Distribuição Percentual', fontsize=10, pad=10)
                
                # Verificar se há valores negativos ou soma zero
                if (chunk[metric_column] < 0).any() or chunk[metric_column].sum() <= 0:
                    ax2.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico', 
                            ha='center', va='center', fontsize=10, color='red', fontweight='bold')
                    ax2.axis('off')
                else:
                    try:
                        # Usar cores mais simples
                        colors = plt.cm.Set3(np.linspace(0, 1, len(chunk)))
                        
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
                            textprops={'fontsize': 7},
                            wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'},
                            pctdistance=0.85,
                            colors=colors
                        )
                        
                        n_cols = min(4, len(chunk))
                        ax2.legend(wedges, chunk['DESCRICAO'],
                                  loc="upper center",
                                  bbox_to_anchor=(0.5, -0.05),
                                  ncol=n_cols,
                                  fontsize=7,
                                  title_fontsize=8,
                                  frameon=False)
                    except Exception:
                        ax2.text(0.5, 0.5, 'Erro no gráfico', 
                                ha='center', va='center', fontsize=9, color='red', fontweight='bold')
                        ax2.axis('off')
                
                # Gráfico de Linha - simplificar
                ax3.set_title('Evolução Temporal (por semana)', fontsize=10, pad=10)
                ax3.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                
                # Filtrar dados de série temporal para os produtos desta página
                ts_filtered = time_series_agg[time_series_agg['ID_AGRUPADO'].isin(produtos_na_pagina)]
                
                if ts_filtered.empty:
                    ax3.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico de linha', 
                            ha='center', va='center', fontsize=10, color='red', fontweight='bold')
                    ax3.axis('off')
                else:
                    # Usar menos cores
                    line_colors = plt.cm.tab10(np.linspace(0, 1, min(10, len(produtos_na_pagina))))
                    
                    for idx, (produto, group) in enumerate(ts_filtered.groupby('ID_AGRUPADO')):
                        group = group.sort_values('SEMANA')
                        color_idx = idx % len(line_colors)
                        
                        line, = ax3.plot(group['SEMANA'], group[metric_column], 
                                       marker='o', linestyle='-', 
                                       color=line_colors[color_idx],
                                       markersize=4, linewidth=1.5)
                    
                    ax3.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d/%m'))
                    ax3.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(byweekday=plt.matplotlib.dates.MO))
                    ax3.grid(True, linestyle=':', alpha=0.5)
                    plt.setp(ax3.get_xticklabels(), rotation=30, ha='right', fontsize=7)
                    plt.setp(ax3.get_yticklabels(), fontsize=7)
                
                # Gráfico de Barras - simplificar
                ax4.set_title(f'{metric_name} por Produto', fontsize=10, pad=10)
                ax4.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                
                # Usar cores simples para barras
                bar_colors = plt.cm.Set3(np.linspace(0, 1, len(chunk)))
                
                bars = ax4.bar(
                    chunk['DESCRICAO'],
                    chunk[metric_column],
                    color=bar_colors
                )
                
                # Adicionar rótulos apenas se não houver muitos produtos
                if len(chunk) <= 10:
                    for bar in bars:
                        height = bar.get_height()
                        
                        if metric_name == 'Faturamento':
                            label = format_currency(height)
                        elif metric_name == 'Margem':
                            label = f'{height:.2f}%'
                        else:
                            label = f'{height:,.1f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                        
                        ax4.text(bar.get_x() + bar.get_width() / 2, height,
                                label, ha='center', va='bottom',
                                fontsize=8)
                
                plt.setp(ax4.get_xticklabels(), rotation=15, ha='right', fontsize=8)
                plt.setp(ax4.get_yticklabels(), fontsize=7)
                ax4.grid(True, axis='y', linestyle=':', alpha=0.5)
                
                plt.tight_layout(rect=[0, 0, 1, 0.95])
                pdf.savefig(fig, dpi=100, bbox_inches='tight', pad_inches=0.5)  # Reduzir DPI
                plt.close(fig)
                clean_matplotlib_memory()  # Limpar memória entre páginas
                
        # Renomear arquivo temporário
        if os.path.exists(temp_path):
            os.rename(temp_path, output_path)
        print(f"Finalizado - {output_filename}")

    except Exception as e:
        print(f"ERRO ao gerar relatório de {metric_name}: {str(e)}")
        import traceback
        traceback.print_exc()
        # Limpar arquivos temporários
        for path in [output_path, temp_path]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

def generate_general_report(file_path, sheet_name, output_dir):
    """Gera um relatório geral com estatísticas básicas e gráficos comparativos"""
    try:
        # 1. Limpar memória antes de começar
        clean_matplotlib_memory()
        plt.switch_backend('agg')  # Usar backend não-interativo
        
        # Ler os dados do Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)
        
        # Verificar se as colunas necessárias existem
        required_columns = ['RAZAO', 'VENDEDOR', 'CODPRODUTO', 'DATA', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']
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
        output_filename = f"Ranking de Vendas - Geral - {nome_mes} {primeiro_ano}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        
        print(f"Gerando - {output_filename}")
        
        # Calcular as métricas básicas
        qtde_clientes = df['RAZAO'].nunique()
        qtde_vendedores = df['VENDEDOR'].nunique()
        qtde_produtos = df['CODPRODUTO'].nunique()
        
        # Criar dados para a tabela
        table_data = [
            ['Qtde Clientes', qtde_clientes],
            ['Qtde Vendedores', qtde_vendedores],
            ['Qtde Produtos', qtde_produtos]
        ]
        
        # Criar PDF com configurações otimizadas
        with PdfPages(output_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16), dpi=100)
            plt.text(0.5, 0.5, "RANKING DE VENDAS - GERAL", 
                    fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                    fontsize=18, ha='center', va='center', fontweight='normal')
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight', dpi=100)
            plt.close(fig_title)
            clean_matplotlib_memory()
            
        print(f"Finalizado - {output_filename}")
        
    except Exception as e:
        print(f"ERRO ao gerar relatório geral: {str(e)}")
        if 'output_path' in locals() and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass

def generate_consolidated_excel(file_path, sheet_name, output_dir):
    """Gera um arquivo Excel consolidado com todas as métricas por produto e por grupos especiais"""
    output_path = None
    
    try:
        # 1. Ler os dados do Excel incluindo a coluna QTDE
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=8)[
            ['CODPRODUTO', 'DESCRICAO', 'DATA', 'QTDE', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']]
        
        # 2. Definir os grupos de produtos (todos em maiúsculas)
        product_groups = {
            'ACEM': [1924, 8006, 1940, 1878, 8101, 1841],
            'ALCATRA C/ MAMINHA': [8001, 1836, 1965, 1800],
            'BARRIGA': [1833, 1639, 1544, 1674, 1863, 1845, 1385, 1898, 1913, 1513, 
                        1444, 1434, 1960, 1954, 5200, 2042, 2043, 2047, 2051],
            'BUCHO': [1567, 1816, 1856, 1480, 1527, 1903, 1855, 1958],
            'BACON MANTA': [869, 981],
            'BANHA SUINA': [1605, 1139],
            'BATATA': [1767, 1872],
            'BOLACHA': [1649, 1644, 1643, 1647, 1645, 1648],
            'BOLINHO': [1709, 1707, 1708, 1941, 1999],
            'CARNE TEMPERADA': [1720, 1623, 1618],
            'CARRE': [1568, 1355, 1443, 1817, 1464, 1640, 1533, 1286, 1518, 1653, 1216, 
                      1316, 906, 1210, 1908, 1221, 1177, 1612, 1634, 917, 1689, 1511,
                      1955],
            'CONTRA FILÉ': [1901, 1922, 1840, 1947, 1894, 1899, 1905, 1503, 1824],
            'CORAÇÃO DE ALCATRA': [1830, 1939],
            'COSTELA BOV': [1768, 1825, 1931, 1814, 1890, 1890],
            'COSTELA MINGA': [1973, 1982],
            'COSTELA SUINA CONGELADA': [1478, 1595, 1506, 1081, 1592, 1412, 1641, 1888, 
                                        1522, 1638, 1607, 1517, 1461, 1416, 1760, 1877, 
                                        1664, 1053, 1314, 1617, 1599, 1896, 1857, 1179,
                                        1324, 1529, 1421, 1323, 1879, 1052, 1051, 1354,
                                        905,  1384, 1086, 1174, 1150, 1758, 1320, 1829,
                                        1665, 1327, 1442, 1431, 1704, 1736, 1445, 1321,
                                        1884, 1535, 8007, 2050],
            'COXÃO DURO': [1920, 8003, 1803, 1949, 1795],
            'COXÃO MOLE': [1831, 8002, 1948, 1976, 1375],
            'COXINHA DA ASA': [1604, 1546, 8005, 1722, 2038],
            'CUPIM A': [1772],
            'CUPIM B': [1804, 1456, 1926, 1984],
            'FIGADO': [1808, 1455, 1818, 1910, 1823, 1537, 1505, 1408, 1373, 1458, 1508,
                       1525, 1454, 1801, 1528, 1530, 1502, 1945, 1967, 1998, 1978, 1983,
                       2018, 2035, 2026],
            'FILÉ MIGNON': [1812, 1919],
            'FRALDA': [1797, 1925],
            'HAMBURGUER': [1009, 1866, 1010],
            'HOT POCKET': [1987],
            'JERKED': [1893, 1943, 1880, 1886, 1851],
            'LAGARTO': [1849, 1396, 1895, 1813],
            'LASANHA': [1003, 1691, 1997, 1002, 1991],
            'LINGUIÇA CALABRESA AURORA': [788, 1974],
            'LINGUIÇA CALABRESA SADIA': [1339, 807, 1848, 1847],
            'LINGUIÇA CALABRESA PAMPLONA': [9165, 910],
            'MEIO DA ASA': [2311, 1937, 2014, 2039, 2052],
            'MINI CHICKEN': [1024, 1994],
            'MINI LASANHA': [1992, 1985],
            'MOCOTÓ': [1539, 1460, 1342, 1540, 1675, 1850, 1827, 1821, 1853, 1407, 1723,
                       1585, 1407, 1723, 1585, 1843, 1584, 762, 1534, 1883, 1509, 1601,
                       1962, 2049, 2028],
            'MUSSARELA': [2000, 947, 1807, 1914],
            'NUGGETS': [1007, 1995],
            'PATINHO': [1805, 1874, 8000, 1938, 9166, 1966],
            'PALETA': [1953, 1964, 1923, 1975],
            'PEITO BOV': [1815, 1875, 1789, 1952],
            'PERNIL SUINO C/OSSO C/PELE': [1942, 1635, 1724, 1570, 1756, 1303],
            'PICANHA B': [1946, 1950],
            'PIZZA': [1989, 1990],
            'PURURUCA 60G': [1288, 1289, 1287],
            'RABO BOV': [1828, 1839, 1876, 1861, 1116, 1705, 1531, 1906, 1826, 1911, 1882,
                        1571, 1335, 1963, 1909, 1473, 1481],
            'SALAME UAI': [1495, 1500, 1496, 1497, 1498, 1499],
            'STEAK FGO': [1718, 1996],
            'TAPIOCA DA TERRINHA': [1929, 1930],
            'YOPRO': [1698, 1701, 1587, 1700, 86754, 1586, 9675]
        }
        
        # Inverter o dicionário para mapear código para nome do grupo
        code_to_group = {}
        for group_name, codes in product_groups.items():
            for code in codes:
                code_to_group[code] = group_name
        
        # 3. Adicionar coluna de grupo ao DataFrame
        df['GRUPO'] = df['CODPRODUTO'].map(code_to_group)
        
        # 4. Calcular o peso unitário (QTDE REAL / QTDE)
        df['PESO_UNITARIO'] = df['QTDE REAL'] / df['QTDE']
        
        # 5. Processar datas e criar caminho de saída
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        output_filename = f"Ranking de Vendas - {nome_mes} {primeiro_ano}.xlsx"
        output_path = os.path.join(report_dir, output_filename)
        
        print(f"Gerando arquivo Excel consolidado - {output_filename}")

        # 6. Limpeza e preparação dos dados
        df['Fat Liquido'] = df['Fat Liquido'].apply(clean_currency)
        df['Lucro / Prej.'] = df['Lucro / Prej.'].apply(clean_currency)
        df = df.dropna(subset=['Fat Liquido', 'Lucro / Prej.', 'QTDE REAL', 'QTDE', 'PESO_UNITARIO'])

        # 7. Obter descrições mais recentes para produtos individuais
        latest_descriptions = df.sort_values('DATA').drop_duplicates('CODPRODUTO', keep='last')[['CODPRODUTO', 'DESCRICAO']]

        # 8. Criar DataFrames separados para produtos agrupados e individuais
        # DataFrame para produtos em grupos
        grouped_df = df[df['GRUPO'].notna()].copy()
        
        # DataFrame para produtos individuais (não estão em nenhum grupo)
        individual_df = df[df['GRUPO'].isna()].copy()
        
        # 9. Agregar dados para os grupos
        group_aggregated = grouped_df.groupby('GRUPO').agg(
            TONELAGEM_KG=('QTDE REAL', 'sum'),
            FATURAMENTO_RS=('Fat Liquido', 'sum'),
            LUCRO_RS=('Lucro / Prej.', 'sum'),
            QTDE_VENDAS=('DATA', 'count'),
            QTDE_TOTAL=('QTDE', 'sum')
        ).reset_index()
        
        # Calcular métricas para os grupos
        group_aggregated['PESO_MEDIO'] = group_aggregated['TONELAGEM_KG'] / group_aggregated['QTDE_TOTAL']
        
        group_aggregated['MARGEM_PERC'] = np.where(
            group_aggregated['FATURAMENTO_RS'] == 0, 0,
            (group_aggregated['LUCRO_RS'] / group_aggregated['FATURAMENTO_RS']) * 100
        ).round(2)
        
        # Adicionar colunas para consistência
        group_aggregated['CODPRODUTO'] = "VÁRIOS PROD."
        group_aggregated['DESCRICAO'] = group_aggregated['GRUPO'].str.upper()
        
        # 10. Agregar dados para produtos individuais
        individual_aggregated = individual_df.groupby('CODPRODUTO').agg(
            TONELAGEM_KG=('QTDE REAL', 'sum'),
            FATURAMENTO_RS=('Fat Liquido', 'sum'),
            LUCRO_RS=('Lucro / Prej.', 'sum'),
            QTDE_VENDAS=('DATA', 'count'),
            QTDE_TOTAL=('QTDE', 'sum')
        ).reset_index()
        
        # Calcular métricas para produtos individuais
        individual_aggregated['PESO_MEDIO'] = individual_aggregated['TONELAGEM_KG'] / individual_aggregated['QTDE_TOTAL']
        
        individual_aggregated['MARGEM_PERC'] = np.where(
            individual_aggregated['FATURAMENTO_RS'] == 0, 0,
            (individual_aggregated['LUCRO_RS'] / individual_aggregated['FATURAMENTO_RS']) * 100
        ).round(2)
        
        # Adicionar descrições aos produtos individuais (em maiúsculas)
        individual_aggregated = pd.merge(
            individual_aggregated, 
            latest_descriptions, 
            on='CODPRODUTO', 
            how='left'
        )
        individual_aggregated['DESCRICAO'] = individual_aggregated['DESCRICAO'].str.upper()
        
        # 11. Combinar todos os dados
        final_df = pd.concat([
            group_aggregated[[
                'CODPRODUTO', 'DESCRICAO', 'PESO_MEDIO', 'TONELAGEM_KG',
                'FATURAMENTO_RS', 'MARGEM_PERC', 'LUCRO_RS', 'QTDE_VENDAS'
            ]],
            individual_aggregated[[
                'CODPRODUTO', 'DESCRICAO', 'PESO_MEDIO', 'TONELAGEM_KG',
                'FATURAMENTO_RS', 'MARGEM_PERC', 'LUCRO_RS', 'QTDE_VENDAS'
            ]]
        ], ignore_index=True)
        
        # Ordenar por tonelagem decrescente
        final_df = final_df.sort_values('TONELAGEM_KG', ascending=False)
        
        # 12. Converter a margem para formato decimal (dividir por 100)
        final_df['MARGEM_PERC'] = final_df['MARGEM_PERC'] / 100
        
        # 13. Gerar o arquivo Excel
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Consolidado', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Consolidado']
            
            # Configurar formatos
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'top',
                'fg_color': '#000000', 'font_color': 'white',
                'border': 1, 'align': 'center', 'font_size': 10
            })
            
            # Formato para porcentagem (ajustado para mostrar 2 decimais)
            percent_format = workbook.add_format({'num_format': '0.00%'})
            
            # Aplicar formatação às colunas
            worksheet.set_column('A:A', 12)  # CODPRODUTO
            worksheet.set_column('B:B', 40)  # DESCRICAO
            worksheet.set_column('C:C', 15, workbook.add_format({'num_format': '#,##0.000'}))  # PESO_MEDIO
            worksheet.set_column('D:D', 15, workbook.add_format({'num_format': '#,##0.000'}))  # TONELAGEM
            worksheet.set_column('E:E', 18, workbook.add_format({'num_format': 'R$ #,##0.00'}))  # FATURAMENTO
            worksheet.set_column('F:F', 12, percent_format)  # MARGEM (formatado como porcentagem)
            worksheet.set_column('G:G', 15, workbook.add_format({'num_format': 'R$ #,##0.00'}))  # LUCRO
            worksheet.set_column('H:H', 12, workbook.add_format({'num_format': '#,##0'}))  # QTDE_VENDAS
            
            # Escrever cabeçalhos
            headers = [
                'CÓDIGO', 'DESCRIÇÃO', 'PESO MÉDIO (KG)',
                'TONELAGEM (KG)', 'FATURAMENTO (R$)', 'MARGEM (%)',
                'LUCRO/PREJ. (R$)', 'QTDE VENDAS'
            ]
            for col_num, value in enumerate(headers):
                worksheet.write(2, col_num, value, header_format)
            
            # Configurações finais
            worksheet.freeze_panes(3, 0)
            worksheet.set_landscape()
            
        print(f"Finalizado - {output_filename}")
        
    except Exception as e:
        print(f"ERRO ao gerar Excel consolidado: {str(e)}")
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass

# Configuração principal
file_path = r"C:\Users\win11\Downloads\260222_MRG - wapp.xlsx"
sheet_name = "Base (3,5%)"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
items_per_page = 5

# Definir as métricas que queremos analisar
metrics = [
    {'column': 'QTDE REAL', 'name': 'Tonelagem', 'unit': 'kg'},
    {'column': 'Fat Liquido', 'name': 'Faturamento', 'unit': 'R$'},
    {'column': 'Margem', 'name': 'Margem', 'unit': '%'}
]

# Executar geradores em sequência com limpeza de memória entre eles
print("Iniciando geração de relatórios...")

# Gerar relatório geral primeiro
try:
    generate_general_report(
        file_path=file_path,
        sheet_name=sheet_name,
        output_dir=output_dir
    )
    clean_matplotlib_memory()
except Exception as e:
    print(f"Erro no relatório geral: {e}")

# Gerar Excel consolidado
try:
    generate_consolidated_excel(
        file_path=file_path,
        sheet_name=sheet_name,
        output_dir=output_dir
    )
    clean_matplotlib_memory()
except Exception as e:
    print(f"Erro no Excel consolidado: {e}")

# Gerar todos os relatórios específicos
for metric in metrics:
    try:
        generate_report(
            file_path=file_path,
            sheet_name=sheet_name,
            output_dir=output_dir,
            metric_column=metric['column'],
            metric_name=metric['name'],
            unit=metric['unit'],
            items_per_page=items_per_page
        )
        # Limpar memória entre cada relatório
        clean_matplotlib_memory()
    except Exception as e:
        print(f"Erro no relatório de {metric['name']}: {e}")
        clean_matplotlib_memory()

print("Processamento concluído!")