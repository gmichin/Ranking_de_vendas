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
import locale

# Tentar configurar locale para português
try:
    locale.setlocale(locale.LC_NUMERIC, 'pt_BR.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_NUMERIC, 'Portuguese_Brazil.1252')
    except:
        pass

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
    gc.collect(generation=2)
    import matplotlib
    matplotlib.pyplot.close('all')

def convert_br_to_float(value):
    """Converte formato brasileiro (1.234,56 ou 1,234.56) para float (1234.56)"""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        # Remove R$ e espaços
        value = value.replace('R$', '').strip()
        value = value.replace(' ', '').strip()
        
        # Verifica se tem vírgula e ponto
        if ',' in value and '.' in value:
            # Caso brasileiro: 1.234,56 (ponto milhar, vírgula decimal)
            if value.rfind(',') > value.rfind('.'):
                # Última vírgula é decimal, pontos são milhar
                value = value.replace('.', '').replace(',', '.')
            else:
                # Último ponto é decimal, vírgulas são milhar
                value = value.replace(',', '')
        elif ',' in value:
            # Só tem vírgula - pode ser decimal ou milhar
            # Se tiver 3 dígitos após a vírgula, provavelmente é milhar
            parts = value.split(',')
            if len(parts) > 1 and len(parts[-1]) == 3:
                # Vírgula é separador de milhar
                value = value.replace(',', '')
            else:
                # Vírgula é decimal
                value = value.replace(',', '.')
        
        try:
            return float(value)
        except:
            print(f"  Aviso: Não foi possível converter '{value}'")
            return 0.0
    return 0.0

def format_currency(value):
    """Formata um valor como moeda brasileira (R$)"""
    return f"R${value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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
        clean_matplotlib_memory()
        
        # Definir os grupos de produtos
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
            'COXINHA DA ASA': [1604, 1546, 8005, 1722, 2038, 1616, 3065],
            'CUPIM A': [1772],
            'CUPIM B': [1804, 1456, 1926, 1984],
            'FIGADO': [1808, 1455, 1818, 1910, 1823, 1537, 1505, 1408, 1373, 1458, 1508,
                       1525, 1454, 1801, 1528, 1530, 1502, 1945, 1967, 1998, 1978, 1983,
                       2018, 2035, 2026, 3012],
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
            'LINGUIÇA CALABRESA PERDIGAO': [1423, 880],
            'MEIO DA ASA': [2311, 1937, 2014, 2039, 2052],
            'MINI CHICKEN': [1024, 1994],
            'MINI LASANHA': [1992, 1985],
            'MOCOTÓ': [1539, 1460, 1342, 1540, 1675, 1850, 1827, 1821, 1853, 1407, 1723,
                       1585, 1407, 1723, 1585, 1843, 1584, 762, 1534, 1883, 1509, 1601,
                       1962, 2049, 2028, 2989, 2081, 2989],
            'MUSSARELA': [2000, 947, 1807, 1914],
            'NUGGETS': [1007, 1995],
            'PATINHO': [1805, 1874, 8000, 1938, 9166, 1966],
            'PALETA': [1953, 1964, 1923, 1975],
            'PEITO BOV': [1815, 1875, 1789, 1952],
            'PERNIL SUINO C/OSSO C/PELE': [1942, 1635, 1724, 1570, 1756, 1303, 3093],
            'PICANHA B': [1946, 1950],
            'PIZZA': [1989, 1990],
            'RABO BOV': [1828, 1839, 1876, 1861, 1116, 1705, 1531, 1906, 1826, 1911, 1882,
                        1571, 1335, 1963, 1909, 1473, 1481, 2079, 859, 1747, 3013, 2424, 3066],
            'SALAME UAI': [1495, 1500, 1496, 1497, 1498, 1499],
            'STEAK FGO': [1718, 1996],
            'TAPIOCA DA TERRINHA': [1929, 1930],
            'YOPRO': [1698, 1701, 1587, 1700, 86754, 1586, 9675]
        }
        
        code_to_group = {}
        for group_name, codes in product_groups.items():
            for code in codes:
                code_to_group[code] = group_name

        # Ler os dados
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=9)[
            ['CODPRODUTO', 'DESCRICAO', 'DATA', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']]
        
        # CONVERSÃO: Converter todos os valores numéricos que podem vir como string
        df['QTDE REAL'] = df['QTDE REAL'].apply(convert_br_to_float)
        df['Fat Liquido'] = df['Fat Liquido'].apply(convert_br_to_float)
        df['Lucro / Prej.'] = df['Lucro / Prej.'].apply(convert_br_to_float)
        
        # Debug: imprimir totais para verificar
        print(f"  DEBUG {metric_name}:")
        print(f"    Total QTDE REAL (original) = {df['QTDE REAL'].sum():.3f}")
        print(f"    Total Fat Liquido (original) = {df['Fat Liquido'].sum():.2f}")
        print(f"    Total Lucro (original) = {df['Lucro / Prej.'].sum():.2f}")
        
        df['GRUPO'] = df['CODPRODUTO'].map(code_to_group)
        df['DATA'] = pd.to_datetime(df['DATA'])
        
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        
        output_filename = f"Ranking de Vendas - {metric_name} - {nome_mes} {primeiro_ano} - {items_per_page} em {items_per_page}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        temp_path = os.path.join(report_dir, f"temp_{output_filename}")

        print(f"Gerando - {output_filename}")
        
        # IMPORTANTE: Cada métrica deve usar seu próprio DataFrame sem filtros cruzados
        # Criar uma cópia para trabalhar
        df_work = df.copy()
        
        # Filtrar apenas para a métrica específica
        if metric_name == 'Tonelagem':
            # Para tonelagem, não filtrar por faturamento
            df_work = df_work[df_work['QTDE REAL'].notna()]
        elif metric_name == 'Faturamento':
            df_work = df_work[df_work['Fat Liquido'].notna()]
            df_work = df_work[df_work['Fat Liquido'] != 0]
        elif metric_name == 'Margem':
            df_work = df_work[df_work['Fat Liquido'].notna()]
            df_work = df_work[df_work['Fat Liquido'] != 0]
            df_work['Margem Calculada'] = np.where(
                df_work['Fat Liquido'] <= 0, 0,
                (df_work['Lucro / Prej.'] / df_work['Fat Liquido']) * 100
            )
        
        print(f"    Após filtro: {len(df_work)} linhas restantes")
        
        # Criar colunas de agrupamento
        df_work['ID_AGRUPADO'] = df_work.apply(
            lambda row: row['GRUPO'] if pd.notna(row['GRUPO']) else str(row['CODPRODUTO']), 
            axis=1
        )
        df_work['DESCRICAO_AGRUPADA'] = df_work.apply(
            lambda row: row['GRUPO'] if pd.notna(row['GRUPO']) else row['DESCRICAO'], 
            axis=1
        )
        
        if metric_name == 'Tonelagem':
            aggregated = df_work.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
                'QTDE REAL': 'sum',
                'DATA': 'count'
            }).reset_index()
            aggregated.rename(columns={'QTDE REAL': metric_column, 'DATA': 'Qtde de vendas'}, inplace=True)
            aggregated['CODPRODUTO'] = aggregated.apply(
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'], axis=1
            )
        elif metric_name == 'Faturamento':
            aggregated = df_work.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
                'Fat Liquido': 'sum',
                'DATA': 'count'
            }).reset_index()
            aggregated.rename(columns={'Fat Liquido': metric_column, 'DATA': 'Qtde de vendas'}, inplace=True)
            aggregated['CODPRODUTO'] = aggregated.apply(
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'], axis=1
            )
        elif metric_name == 'Margem':
            aggregated = df_work.groupby(['ID_AGRUPADO', 'DESCRICAO_AGRUPADA']).agg({
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
                lambda row: 'GRUPO' if row['ID_AGRUPADO'] in product_groups.keys() else row['ID_AGRUPADO'], axis=1
            )
        
        aggregated.rename(columns={'ID_AGRUPADO': 'GRUPO_ID', 'DESCRICAO_AGRUPADA': 'DESCRICAO'}, inplace=True)
        sorted_df = aggregated.sort_values(metric_column, ascending=False).reset_index(drop=True)
        sorted_df.insert(0, 'Posição', range(1, len(sorted_df)+1))
        
        # Calcular total usando o df_work (já filtrado para a métrica)
        if metric_name == 'Tonelagem':
            total_metric = df_work['QTDE REAL'].sum()
        elif metric_name == 'Faturamento':
            total_metric = df_work['Fat Liquido'].sum()
        elif metric_name == 'Margem':
            total_lucro = df_work['Lucro / Prej.'].sum()
            total_fat = df_work['Fat Liquido'].sum()
            total_metric = 0 if total_fat <= 0 else (total_lucro / total_fat) * 100
        
        print(f"    TOTAL FINAL = {total_metric:.2f}")
        
        if metric_name == 'Faturamento':
            total_text = format_currency(total_metric)
        elif metric_name == 'Margem':
            total_text = f"{total_metric:.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            total_text = f"{total_metric:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        # Série temporal - usar o df_work que já tem as colunas ID_AGRUPADO
        time_series = df_work.copy()
        time_series['SEMANA'] = time_series['DATA'].dt.to_period('W').dt.start_time
        
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
        
        id_to_description = dict(zip(aggregated['GRUPO_ID'], aggregated['DESCRICAO']))
        time_series_agg['DESCRICAO'] = time_series_agg['ID_AGRUPADO'].map(id_to_description)
        
        plt.switch_backend('agg')
        
        with PdfPages(temp_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16), dpi=100)
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
            
            total_pages = (len(sorted_df) + items_per_page - 1) // items_per_page
            
            for page_num in range(total_pages):
                i = page_num * items_per_page
                chunk = sorted_df.iloc[i:i+items_per_page]
                produtos_na_pagina = chunk['GRUPO_ID'].tolist()
                
                fig = plt.figure(figsize=(11, 16), dpi=100)
                gs = fig.add_gridspec(4, 1)
                ax1 = fig.add_subplot(gs[0])
                ax2 = fig.add_subplot(gs[1])
                ax3 = fig.add_subplot(gs[2])
                ax4 = fig.add_subplot(gs[3])
                
                fig.suptitle(f'Ranking de Produtos {i+1}-{min(i+items_per_page, len(sorted_df))}', fontsize=14, y=1.02)
                ax1.axis('off')
                
                display_values = chunk[metric_column].copy()
                if metric_name == 'Faturamento':
                    display_values = display_values.apply(format_currency)
                elif metric_name == 'Margem':
                    display_values = display_values.apply(lambda x: f"{x:.2f}%")
                else:
                    display_values = display_values.apply(lambda x: f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
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
                
                ax2.set_title('Distribuição Percentual', fontsize=10, pad=10)
                if (chunk[metric_column] < 0).any() or chunk[metric_column].sum() <= 0:
                    ax2.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico', ha='center', va='center', fontsize=10, color='red')
                    ax2.axis('off')
                else:
                    try:
                        colors = plt.cm.Set3(np.linspace(0, 1, len(chunk)))
                        if metric_name == 'Faturamento':
                            autopct_format = lambda p: f'{p:.1f}%\n({format_currency(p*sum(chunk[metric_column])/100)})'
                        elif metric_name == 'Margem':
                            autopct_format = lambda p: f'{p:.1f}%\n({p*sum(chunk[metric_column])/100:.2f}%)'
                        else:
                            autopct_format = lambda p: f'{p:.1f}%\n({p*sum(chunk[metric_column])/100:,.1f} {unit})'.replace(",", "X").replace(".", ",").replace("X", ".")
                        wedges, texts, autotexts = ax2.pie(chunk[metric_column], autopct=autopct_format, startangle=140, textprops={'fontsize': 7}, wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'}, pctdistance=0.85, colors=colors)
                        n_cols = min(4, len(chunk))
                        ax2.legend(wedges, chunk['DESCRICAO'], loc="upper center", bbox_to_anchor=(0.5, -0.05), ncol=n_cols, fontsize=7, frameon=False)
                    except Exception:
                        ax2.text(0.5, 0.5, 'Erro no gráfico', ha='center', va='center', fontsize=9, color='red')
                        ax2.axis('off')
                
                ax3.set_title('Evolução Temporal (por semana)', fontsize=10, pad=10)
                ax3.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                ts_filtered = time_series_agg[time_series_agg['ID_AGRUPADO'].isin(produtos_na_pagina)]
                if ts_filtered.empty:
                    ax3.text(0.5, 0.5, 'Dados insuficientes\npara o gráfico de linha', ha='center', va='center', fontsize=10, color='red')
                    ax3.axis('off')
                else:
                    line_colors = plt.cm.tab10(np.linspace(0, 1, min(10, len(produtos_na_pagina))))
                    for idx, (produto, group) in enumerate(ts_filtered.groupby('ID_AGRUPADO')):
                        group = group.sort_values('SEMANA')
                        ax3.plot(group['SEMANA'], group[metric_column], marker='o', linestyle='-', color=line_colors[idx % len(line_colors)], markersize=4, linewidth=1.5)
                    ax3.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d/%m'))
                    ax3.xaxis.set_major_locator(plt.matplotlib.dates.WeekdayLocator(byweekday=plt.matplotlib.dates.MO))
                    ax3.grid(True, linestyle=':', alpha=0.5)
                    plt.setp(ax3.get_xticklabels(), rotation=30, ha='right', fontsize=7)
                
                ax4.set_title(f'{metric_name} por Produto', fontsize=10, pad=10)
                ax4.set_ylabel(f'{metric_name} ({unit})', fontsize=8)
                bar_colors = plt.cm.Set3(np.linspace(0, 1, len(chunk)))
                bars = ax4.bar(chunk['DESCRICAO'], chunk[metric_column], color=bar_colors)
                if len(chunk) <= 10:
                    for bar in bars:
                        height = bar.get_height()
                        if metric_name == 'Faturamento':
                            label = format_currency(height)
                        elif metric_name == 'Margem':
                            label = f'{height:.2f}%'
                        else:
                            label = f'{height:,.1f}'.replace(",", "X").replace(".", ",").replace("X", ".")
                        ax4.text(bar.get_x() + bar.get_width() / 2, height, label, ha='center', va='bottom', fontsize=8)
                plt.setp(ax4.get_xticklabels(), rotation=15, ha='right', fontsize=8)
                ax4.grid(True, axis='y', linestyle=':', alpha=0.5)
                
                plt.tight_layout(rect=[0, 0, 1, 0.95])
                pdf.savefig(fig, dpi=100, bbox_inches='tight', pad_inches=0.5)
                plt.close(fig)
                clean_matplotlib_memory()
        
        if os.path.exists(temp_path):
            os.rename(temp_path, output_path)
        print(f"  Finalizado - {output_filename}")

    except Exception as e:
        print(f"ERRO ao gerar relatório de {metric_name}: {str(e)}")
        import traceback
        traceback.print_exc()
        for path in [output_path, temp_path]:
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass

def generate_general_report(file_path, sheet_name, output_dir):
    """Gera um relatório geral com estatísticas básicas e gráficos comparativos"""
    try:
        # Ler os dados do Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=9)
        
        # Converter valores
        df['QTDE REAL'] = df['QTDE REAL'].apply(convert_br_to_float)
        df['Fat Liquido'] = df['Fat Liquido'].apply(convert_br_to_float)
        df['Lucro / Prej.'] = df['Lucro / Prej.'].apply(convert_br_to_float)
        
        # Verificar colunas
        required_columns = ['RAZAO', 'VENDEDOR', 'CODPRODUTO', 'DATA', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Colunas faltando: {', '.join(missing_columns)}")
        
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        
        output_filename = f"Ranking de Vendas - Geral - {nome_mes} {primeiro_ano}.pdf"
        output_path = os.path.join(report_dir, output_filename)
        
        print(f"Gerando - {output_filename}")
        
        # Calcular métricas
        qtde_clientes = df['RAZAO'].nunique()
        qtde_vendedores = df['VENDEDOR'].nunique()
        qtde_produtos = df['CODPRODUTO'].nunique()
        
        table_data = [
            ['Qtde Clientes', qtde_clientes],
            ['Qtde Vendedores', qtde_vendedores],
            ['Qtde Produtos', qtde_produtos]
        ]
        
        def prepare_pie_data(metric_name):
            if metric_name == 'Tonelagem':
                metric_column = 'QTDE REAL'
                grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            elif metric_name == 'Faturamento':
                metric_column = 'Fat Liquido'
                grouped = df.groupby('CODPRODUTO')[metric_column].sum().reset_index()
            elif metric_name == 'Margem':
                grouped = df.groupby('CODPRODUTO').agg({
                    'Lucro / Prej.': 'sum',
                    'Fat Liquido': 'sum'
                }).reset_index()
                grouped['Margem'] = (grouped['Lucro / Prej.'] / grouped['Fat Liquido']) * 100
                grouped = grouped.replace([np.inf, -np.inf], 0)
                metric_column = 'Margem'

            sorted_df = grouped.sort_values(metric_column, ascending=False).reset_index(drop=True)
            top20 = sorted_df.head(20)
            resto = sorted_df.iloc[20:]

            if metric_name == 'Margem':
                total_top20_lucro = top20['Lucro / Prej.'].sum()
                total_top20_fat = top20['Fat Liquido'].sum()
                total_top20 = (total_top20_lucro / total_top20_fat) * 100 if total_top20_fat != 0 else 0

                total_resto_lucro = resto['Lucro / Prej.'].sum()
                total_resto_fat = resto['Fat Liquido'].sum()
                total_resto = (total_resto_lucro / total_resto_fat) * 100 if total_resto_fat != 0 else 0

                total_geral_lucro = df['Lucro / Prej.'].sum()
                total_geral_fat = df['Fat Liquido'].sum()
                total_geral = (total_geral_lucro / total_geral_fat) * 100 if total_geral_fat != 0 else 0

                lucro_total = abs(total_top20_lucro) + abs(total_resto_lucro)
                if lucro_total > 0:
                    perc_top20 = 100 * abs(total_top20_lucro) / lucro_total
                    perc_resto = 100 * abs(total_resto_lucro) / lucro_total

                return {
                    'labels': ['Top 20 produtos', f'Outros {len(resto)} produtos'],
                    'values': [abs(total_top20_lucro), abs(total_resto_lucro)],
                    'display_values': [total_top20, total_resto],
                    'title': f'Top 20 produtos vs Outros {len(resto)} produtos - {metric_name}',
                    'total': total_geral,
                    'is_margin': True
                }
            else:
                total_top20 = top20[metric_column].sum()
                total_resto = resto[metric_column].sum()
                total_geral = total_top20 + total_resto

                return {
                    'labels': ['Top 20 produtos', f'Outros {len(resto)} produtos'],
                    'values': [total_top20, total_resto],
                    'title': f'Top 20 produtos vs Outros {len(resto)} produtos - {metric_name}',
                    'total': total_geral,
                    'is_margin': False
                }
        
        pie_data = {
            'Tonelagem': prepare_pie_data('Tonelagem'),
            'Faturamento': prepare_pie_data('Faturamento'),
            'Margem': prepare_pie_data('Margem')
        }
        
        with PdfPages(output_path) as pdf:
            # Página de título
            fig_title = plt.figure(figsize=(11, 16))
            plt.text(0.5, 0.5, "RANKING DE VENDAS - GERAL", 
                    fontsize=24, ha='center', va='center', fontweight='bold')
            plt.text(0.5, 0.45, f"{nome_mes} {primeiro_ano}", 
                    fontsize=18, ha='center', va='center', fontweight='normal')
            plt.axis('off')
            pdf.savefig(fig_title, bbox_inches='tight')
            plt.close(fig_title)
            
            # Página de conteúdo
            fig_content = plt.figure(figsize=(11, 16))
            gs = fig_content.add_gridspec(4, 1, height_ratios=[1, 1, 1, 1], hspace=0.5)
            
            ax_table = fig_content.add_subplot(gs[0])
            ax_table.axis('off')
            table = ax_table.table(cellText=table_data, loc='center', cellLoc='center', colWidths=[0.5, 0.5])
            table.auto_set_font_size(False)
            table.set_fontsize(12)
            table.scale(1, 2)
            ax_table.set_title("Estatísticas Gerais de Vendas", fontsize=14, pad=20)
            
            colors = ['#1f77b4', '#ff7f0e']
            
            for i, (metric_name, data) in enumerate(pie_data.items()):
                ax_pie = fig_content.add_subplot(gs[i+1])
                ax_pie.set_position([0.1, ax_pie.get_position().y0, 0.8, ax_pie.get_position().height * 0.8])
                ax_pie.set_title(data['title'], fontsize=12, pad=25, y=1.08)
                
                if metric_name == 'Faturamento':
                    total_text = format_currency(data['total'])
                elif metric_name == 'Margem':
                    total_text = f"{data['total']:.2f}%"
                else:
                    total_text = f"{data['total']:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
                
                ax_pie.text(0.5, 1.02, f"Total: {total_text}", fontsize=11, ha='center', va='bottom', 
                           color=TOTAL_COLOR, transform=ax_pie.transAxes)
                
                wedges, texts = ax_pie.pie(data['values'], startangle=90, colors=colors,
                                          wedgeprops={'linewidth': 0.5, 'edgecolor': 'white'}, radius=0.8)
                
                ax_pie.legend(wedges, data['labels'], loc='lower center', bbox_to_anchor=(0.5, -0.3),
                             ncol=2, fontsize=10, frameon=False)
            
            plt.tight_layout()
            pdf.savefig(fig_content, bbox_inches='tight')
            plt.close(fig_content)
            
        print(f"Finalizado - {output_filename}")
        
    except Exception as e:
        print(f"ERRO no relatório geral: {str(e)}")
        import traceback
        traceback.print_exc()

def generate_consolidated_excel(file_path, sheet_name, output_dir):
    """Gera arquivo Excel consolidado"""
    output_path = None
    
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=9)[
            ['CODPRODUTO', 'DESCRICAO', 'DATA', 'QTDE', 'QTDE REAL', 'Fat Liquido', 'Lucro / Prej.']]
        
        # Converter valores
        df['QTDE REAL'] = df['QTDE REAL'].apply(convert_br_to_float)
        df['Fat Liquido'] = df['Fat Liquido'].apply(convert_br_to_float)
        df['Lucro / Prej.'] = df['Lucro / Prej.'].apply(convert_br_to_float)
        
        # Grupos
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
            'COXINHA DA ASA': [1604, 1546, 8005, 1722, 2038, 1616, 3065],
            'CUPIM A': [1772],
            'CUPIM B': [1804, 1456, 1926, 1984],
            'FIGADO': [1808, 1455, 1818, 1910, 1823, 1537, 1505, 1408, 1373, 1458, 1508,
                       1525, 1454, 1801, 1528, 1530, 1502, 1945, 1967, 1998, 1978, 1983,
                       2018, 2035, 2026, 3012],
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
            'LINGUIÇA CALABRESA PERDIGAO': [1423, 880],
            'MEIO DA ASA': [2311, 1937, 2014, 2039, 2052],
            'MINI CHICKEN': [1024, 1994],
            'MINI LASANHA': [1992, 1985],
            'MOCOTÓ': [1539, 1460, 1342, 1540, 1675, 1850, 1827, 1821, 1853, 1407, 1723,
                       1585, 1407, 1723, 1585, 1843, 1584, 762, 1534, 1883, 1509, 1601,
                       1962, 2049, 2028, 1989, 2081, 2989],
            'MUSSARELA': [2000, 947, 1807, 1914],
            'NUGGETS': [1007, 1995],
            'PATINHO': [1805, 1874, 8000, 1938, 9166, 1966],
            'PALETA': [1953, 1964, 1923, 1975],
            'PEITO BOV': [1815, 1875, 1789, 1952],
            'PERNIL SUINO C/OSSO C/PELE': [1942, 1635, 1724, 1570, 1756, 1303, 3093],
            'PICANHA B': [1946, 1950],
            'PIZZA': [1989, 1990],
            'RABO BOV': [1828, 1839, 1876, 1861, 1116, 1705, 1531, 1906, 1826, 1911, 1882,
                        1571, 1335, 1963, 1909, 1473, 1481, 2079, 8599, 1747, 3013, 2424, 3066],
            'SALAME UAI': [1495, 1500, 1496, 1497, 1498, 1499],
            'STEAK FGO': [1718, 1996],
            'TAPIOCA DA TERRINHA': [1929, 1930],
            'YOPRO': [1698, 1701, 1587, 1700, 86754, 1586, 9675]
        }
        
        code_to_group = {}
        for group_name, codes in product_groups.items():
            for code in codes:
                code_to_group[code] = group_name
        
        df['GRUPO'] = df['CODPRODUTO'].map(code_to_group)
        df['PESO_UNITARIO'] = np.where(df['QTDE'] != 0, df['QTDE REAL'] / df['QTDE'], 0)
        
        df['DATA'] = pd.to_datetime(df['DATA'])
        primeiro_mes = df['DATA'].iloc[0].month
        primeiro_ano = df['DATA'].iloc[0].year
        nome_mes = MESES_PT.get(primeiro_mes, f'Mês {primeiro_mes}')
        report_dir = create_report_directory(output_dir, nome_mes, primeiro_ano)
        output_filename = f"Ranking de Vendas - {nome_mes} {primeiro_ano}.xlsx"
        output_path = os.path.join(report_dir, output_filename)
        
        print(f"Gerando Excel consolidado - {output_filename}")
        
        df = df.dropna(subset=['Fat Liquido', 'Lucro / Prej.', 'QTDE REAL', 'QTDE'])
        
        latest_descriptions = df.sort_values('DATA').drop_duplicates('CODPRODUTO', keep='last')[['CODPRODUTO', 'DESCRICAO']]
        
        grouped_df = df[df['GRUPO'].notna()].copy()
        individual_df = df[df['GRUPO'].isna()].copy()
        
        group_aggregated = grouped_df.groupby('GRUPO').agg(
            TONELAGEM_KG=('QTDE REAL', 'sum'),
            FATURAMENTO_RS=('Fat Liquido', 'sum'),
            LUCRO_RS=('Lucro / Prej.', 'sum'),
            QTDE_VENDAS=('DATA', 'count'),
            QTDE_TOTAL=('QTDE', 'sum')
        ).reset_index()
        
        group_aggregated['PESO_MEDIO'] = np.where(
            group_aggregated['QTDE_TOTAL'] != 0,
            group_aggregated['TONELAGEM_KG'] / group_aggregated['QTDE_TOTAL'],
            0
        )
        
        group_aggregated['MARGEM_PERC'] = np.where(
            group_aggregated['FATURAMENTO_RS'] != 0,
            (group_aggregated['LUCRO_RS'] / group_aggregated['FATURAMENTO_RS']) * 100,
            0
        ).round(2)
        
        group_aggregated['CODPRODUTO'] = "VÁRIOS PROD."
        group_aggregated['DESCRICAO'] = group_aggregated['GRUPO'].str.upper()
        
        individual_aggregated = individual_df.groupby('CODPRODUTO').agg(
            TONELAGEM_KG=('QTDE REAL', 'sum'),
            FATURAMENTO_RS=('Fat Liquido', 'sum'),
            LUCRO_RS=('Lucro / Prej.', 'sum'),
            QTDE_VENDAS=('DATA', 'count'),
            QTDE_TOTAL=('QTDE', 'sum')
        ).reset_index()
        
        individual_aggregated['PESO_MEDIO'] = np.where(
            individual_aggregated['QTDE_TOTAL'] != 0,
            individual_aggregated['TONELAGEM_KG'] / individual_aggregated['QTDE_TOTAL'],
            0
        )
        
        individual_aggregated['MARGEM_PERC'] = np.where(
            individual_aggregated['FATURAMENTO_RS'] != 0,
            (individual_aggregated['LUCRO_RS'] / individual_aggregated['FATURAMENTO_RS']) * 100,
            0
        ).round(2)
        
        individual_aggregated = pd.merge(individual_aggregated, latest_descriptions, on='CODPRODUTO', how='left')
        individual_aggregated['DESCRICAO'] = individual_aggregated['DESCRICAO'].str.upper()
        
        final_df = pd.concat([
            group_aggregated[['CODPRODUTO', 'DESCRICAO', 'PESO_MEDIO', 'TONELAGEM_KG', 'FATURAMENTO_RS', 'MARGEM_PERC', 'LUCRO_RS', 'QTDE_VENDAS']],
            individual_aggregated[['CODPRODUTO', 'DESCRICAO', 'PESO_MEDIO', 'TONELAGEM_KG', 'FATURAMENTO_RS', 'MARGEM_PERC', 'LUCRO_RS', 'QTDE_VENDAS']]
        ], ignore_index=True)
        
        final_df = final_df.sort_values('TONELAGEM_KG', ascending=False)
        final_df['MARGEM_PERC'] = final_df['MARGEM_PERC'] / 100
        
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, sheet_name='Consolidado', index=False, startrow=2)
            
            workbook = writer.book
            worksheet = writer.sheets['Consolidado']
            
            header_format = workbook.add_format({'bold': True, 'fg_color': '#000000', 'font_color': 'white', 'border': 1, 'align': 'center', 'font_size': 10})
            percent_format = workbook.add_format({'num_format': '0.00%'})
            
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 40)
            worksheet.set_column('C:C', 15, workbook.add_format({'num_format': '#,##0.000'}))
            worksheet.set_column('D:D', 15, workbook.add_format({'num_format': '#,##0.000'}))
            worksheet.set_column('E:E', 18, workbook.add_format({'num_format': 'R$ #,##0.00'}))
            worksheet.set_column('F:F', 12, percent_format)
            worksheet.set_column('G:G', 15, workbook.add_format({'num_format': 'R$ #,##0.00'}))
            worksheet.set_column('H:H', 12, workbook.add_format({'num_format': '#,##0'}))
            
            headers = ['CÓDIGO', 'DESCRIÇÃO', 'PESO MÉDIO (KG)', 'TONELAGEM (KG)', 'FATURAMENTO (R$)', 'MARGEM (%)', 'LUCRO/PREJ. (R$)', 'QTDE VENDAS']
            for col_num, value in enumerate(headers):
                worksheet.write(2, col_num, value, header_format)
            
            worksheet.freeze_panes(3, 0)
            worksheet.set_landscape()
            
        print(f"Finalizado - {output_filename}")
        
    except Exception as e:
        print(f"ERRO no Excel: {str(e)}")
        import traceback
        traceback.print_exc()

# Configuração principal
file_path = r"C:\Users\win11\Downloads\260602_MRG - wapp.xlsx"
sheet_name = "FEC_PQ"
output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
items_per_page = 5

metrics = [
    {'column': 'QTDE REAL', 'name': 'Tonelagem', 'unit': 'kg'},
    {'column': 'Fat Liquido', 'name': 'Faturamento', 'unit': 'R$'},
    {'column': 'Margem', 'name': 'Margem', 'unit': '%'}
]

print("Iniciando geração de relatórios...")
print("=" * 60)

# Gerar relatório geral
try:
    generate_general_report(file_path, sheet_name, output_dir)
    clean_matplotlib_memory()
except Exception as e:
    print(f"Erro no relatório geral: {e}")

# Gerar Excel consolidado
try:
    generate_consolidated_excel(file_path, sheet_name, output_dir)
    clean_matplotlib_memory()
except Exception as e:
    print(f"Erro no Excel consolidado: {e}")

# Gerar relatórios específicos
for metric in metrics:
    try:
        generate_report(file_path, sheet_name, output_dir, metric['column'], metric['name'], metric['unit'], items_per_page)
        clean_matplotlib_memory()
    except Exception as e:
        print(f"Erro no relatório de {metric['name']}: {e}")
        clean_matplotlib_memory()

print("=" * 60)
print("Processamento concluído!")