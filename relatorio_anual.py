import os
import json
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import PyPDF2
import re
from datetime import datetime

# Configurações iniciais
caminho_base = r"C:\Users\win11\OneDrive\Documentos\Relatórios Analítico de Vendas\2025"
meses = ['Maio', 'Junho']
ano = '2025'
output_pdf = os.path.join(os.path.expanduser('~'), 'Downloads', 'Relatorio_Consolidado_Vendas.pdf')

# Dicionário para armazenar os dados
dados_consolidados = {}

# Função para extrair texto de PDF
def extrair_texto_pdf(caminho_pdf):
    texto = ""
    with open(caminho_pdf, 'rb') as arquivo:
        leitor = PyPDF2.PdfReader(arquivo)
        for pagina in leitor.pages:
            texto += pagina.extract_text()
    return texto

def extrair_dados(texto):
    dados = {}
    
    # Extração básica
    padrao_clientes = r"Qtde Clientes\s*([\d,.]+)"
    padrao_vendedores = r"Qtde Vendedores\s*([\d,.]+)"
    padrao_produtos = r"Qtde Produtos\s*([\d,.]+)"
    
    match_clientes = re.search(padrao_clientes, texto)
    match_vendedores = re.search(padrao_vendedores, texto)
    match_produtos = re.search(padrao_produtos, texto)
    
    if match_clientes:
        dados['Qtde Clientes'] = float(match_clientes.group(1).replace('.', '').replace(',', '.'))
    if match_vendedores:
        dados['Qtde Vendedores'] = float(match_vendedores.group(1).replace('.', '').replace(',', '.'))
    if match_produtos:
        dados['Qtde Produtos'] = float(match_produtos.group(1).replace('.', '').replace(',', '.'))
    
    # Extração de totais
    padrao_faturamento = r"T otal:\s*R\$([\d\.]+,\d+)T op 20 produtos vs Resto - Faturamento"
    padrao_tonelagem = r"T otal:\s*([\d\.,]+)T op 20 produtos vs Resto - T onelagem"
    padrao_margem = r"T otal:\s*([\d\.]+)%T op 20 produtos vs Resto - Margem"
    
    if match_faturamento := re.search(padrao_faturamento, texto):
        dados['total_faturamento'] = float(match_faturamento.group(1).replace('.', '').replace(',', '.'))
    
    if match_tonelagem := re.search(padrao_tonelagem, texto):
        dados['total_tonelagem'] = float(match_tonelagem.group(1).replace('.', '').replace(',', '.'))
    
    if match_margem := re.search(padrao_margem, texto):
        dados['total_margem'] = float(match_margem.group(1))
    
    # Extração Top 20 vs Resto
    padrao_top20_ton = r"\((\d+\.\d+,\d+)\)\d+\.\d+%"
    padrao_resto_ton = r"\((\d+\.\d+,\d+)\)T otal:"
    padrao_top20_fat = r"\(R\$([\d\.]+,\d+)\)\d+\.\d+%"
    padrao_resto_fat = r"\(R\$([\d\.]+,\d+)\)T otal:"
    padrao_top20_mar = r"\((\d+\.\d+)%\)\d+\.\d+%"
    padrao_resto_mar = r"\((\d+\.\d+)%\)T otal:"
    
    if match := re.search(padrao_top20_ton, texto):
        dados['top20_ton'] = float(match.group(1).replace('.', '').replace(',', '.'))
    if match := re.search(padrao_resto_ton, texto):
        dados['resto_ton'] = float(match.group(1).replace('.', '').replace(',', '.'))
    
    if match := re.search(padrao_top20_fat, texto):
        dados['top20_fat'] = float(match.group(1).replace('.', '').replace(',', '.'))
    if match := re.search(padrao_resto_fat, texto):
        dados['resto_fat'] = float(match.group(1).replace('.', '').replace(',', '.'))
    
    if match := re.search(padrao_top20_mar, texto):
        dados['top20_mar'] = float(match.group(1))
    if match := re.search(padrao_resto_mar, texto):
        dados['resto_mar'] = float(match.group(1))
    
    return dados

# Processar cada arquivo PDF
for mes in meses:
    nome_arquivo = f"Relatório Analítico de Vendas - Geral - {mes} {ano}.pdf"
    caminho_pdf = os.path.join(caminho_base, mes, nome_arquivo)
    
    if os.path.exists(caminho_pdf):
        texto = extrair_texto_pdf(caminho_pdf)
        dados = extrair_dados(texto)
        dados_consolidados[mes] = dados

# Configurações de layout
PAGE_WIDTH = 11.69
PAGE_HEIGHT = 8.27
FONT_SIZE = 10
DPI = 300

def formatar_valor(valor, tipo):
    if pd.isna(valor):
        return "N/A"
    if tipo == 'tonelagem':
        return f"{valor:,.3f} kg".replace(',', 'X').replace('.', ',').replace('X', '.')
    elif tipo == 'faturamento':
        return f"R${valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    elif tipo == 'margem':
        return f"{valor:.2f}%".replace('.', ',')
    return str(valor)

# Criar PDF
with PdfPages(output_pdf) as pdf:
    plt.rcParams.update({
        'font.size': FONT_SIZE,
        'figure.dpi': DPI,
        'figure.constrained_layout.use': True
    })
    
    # Página de título com data e hora de geração
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.6, 'Relatório Consolidado de Vendas', 
             ha='center', va='center', fontsize=16, fontweight='bold')
    plt.text(0.5, 0.5, f'Período: Maio a Junho de {ano}', 
             ha='center', va='center', fontsize=12)
    plt.text(0.5, 0.4, f'Gerado em: {data_hora}', 
             ha='center', va='center', fontsize=10)
    plt.axis('off')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção de Quantidades
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '1. Análise de Quantidades', 
             ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
    plt.axis('off')
    
    # Tabela de quantidades (valores inteiros)
    ax_table = fig.add_axes([0.1, 0.55, 0.8, 0.35])
    ax_table.axis('off')
    
    df_quantidades = pd.DataFrame.from_dict(dados_consolidados, orient='index')[['Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos']]
    # Converter para inteiros
    df_quantidades = df_quantidades.astype(int)
    
    tabela = ax_table.table(
        cellText=df_quantidades.reset_index().values,
        colLabels=['Mês'] + df_quantidades.columns.tolist(),
        loc='center',
        cellLoc='center'
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(FONT_SIZE)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#f2f2f2')
    
    # Gráfico de quantidades com valores nos pontos
    ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
    lines = df_quantidades.plot(ax=ax_graph, marker='o')
    ax_graph.set_title('Evolução Mensal das Quantidades', fontweight='bold')
    ax_graph.grid(True, linestyle='--', alpha=0.7)
    ax_graph.legend()
    
    # Adicionar valores nos pontos do gráfico
    for line in lines.get_lines():
        x = line.get_xdata()
        y = line.get_ydata()
        for xi, yi in zip(x, y):
            ax_graph.text(xi, yi, f'{int(yi)}', 
                         ha='center', va='bottom', fontsize=FONT_SIZE)
    
    # Ajustar limite superior para não cortar os valores
    ylim = ax_graph.get_ylim()
    ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção de Totais - Página 1 (Tabela + 1 gráfico)
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '2. Análise de Totais', 
             ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
    plt.axis('off')
    
    # Tabela de totais (corrigir formatação de margem)
    ax_table = fig.add_axes([0.1, 0.55, 0.8, 0.35])
    ax_table.axis('off')
    
    df_totais = pd.DataFrame.from_dict(dados_consolidados, orient='index')[['total_tonelagem', 'total_faturamento', 'total_margem']]
    tabela_dados = []
    for mes, row in df_totais.iterrows():
        linha = [mes]
        linha.append(formatar_valor(row['total_tonelagem'], 'tonelagem'))
        linha.append(formatar_valor(row['total_faturamento'], 'faturamento'))
        linha.append(formatar_valor(row['total_margem'], 'margem'))
        tabela_dados.append(linha)
    
    tabela = ax_table.table(
        cellText=tabela_dados,
        colLabels=['Mês', 'Tonelagem', 'Faturamento', 'Margem'],
        loc='center',
        cellLoc='center'
    )
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(FONT_SIZE)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#e6e6e6')
    
    # Gráfico de faturamento com espaço superior
    ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
    bars = df_totais['total_faturamento'].plot(kind='bar', ax=ax_graph, color='#2ca02c')
    ax_graph.set_title('Faturamento Total', fontweight='bold')
    ax_graph.set_ylabel('R$')
    ax_graph.grid(True, linestyle='--', alpha=0.7, axis='y')
    
    # Adicionar valores nas barras
    for i, valor in enumerate(df_totais['total_faturamento']):
        ax_graph.text(i, valor, formatar_valor(valor, 'faturamento'), 
                     ha='center', va='bottom')
    
    # Ajustar limite superior
    ylim = ax_graph.get_ylim()
    ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção de Totais - Página 2 (2 gráficos restantes)
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '2. Análise de Totais (continuação)', 
             ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
    plt.axis('off')
    
    # Gráfico de tonelagem com espaço superior
    ax_graph1 = fig.add_axes([0.1, 0.55, 0.8, 0.35])
    df_totais['total_tonelagem'].plot(kind='bar', ax=ax_graph1, color='#1f77b4')
    ax_graph1.set_title('Tonelagem Total', fontweight='bold')
    ax_graph1.set_ylabel('kg')
    ax_graph1.grid(True, linestyle='--', alpha=0.7, axis='y')
    
    # Adicionar valores nas barras
    for i, valor in enumerate(df_totais['total_tonelagem']):
        ax_graph1.text(i, valor, formatar_valor(valor, 'tonelagem'), 
                      ha='center', va='bottom')
    
    # Ajustar limite superior
    ylim = ax_graph1.get_ylim()
    ax_graph1.set_ylim(ylim[0], ylim[1] * 1.1)
    
    # Gráfico de margem com espaço superior
    ax_graph2 = fig.add_axes([0.1, 0.1, 0.8, 0.35])
    df_totais['total_margem'].plot(kind='bar', ax=ax_graph2, color='#ff7f0e')
    ax_graph2.set_title('Margem Total', fontweight='bold')
    ax_graph2.set_ylabel('%')
    ax_graph2.grid(True, linestyle='--', alpha=0.7, axis='y')
    
    # Adicionar valores nas barras
    for i, valor in enumerate(df_totais['total_margem']):
        ax_graph2.text(i, valor, formatar_valor(valor, 'margem'), 
                      ha='center', va='bottom')
    
    # Ajustar limite superior
    ylim = ax_graph2.get_ylim()
    ax_graph2.set_ylim(ylim[0], ylim[1] * 1.1)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção Top 20 vs Resto - Tabelas separadas por métrica
    for tipo, titulo, unidade, fmt in [('ton', 'Tonelagem', 'kg', 'tonelagem'),
                                     ('fat', 'Faturamento', 'R$', 'faturamento'),
                                     ('mar', 'Margem', '%', 'margem')]:
        
        fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
        plt.text(0.5, 0.95, f'3. Análise Top 20 vs Resto - {titulo}', 
                 ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
        plt.axis('off')
        
        # Tabela
        ax_table = fig.add_axes([0.1, 0.5, 0.8, 0.4])
        ax_table.axis('off')
        
        tabela_dados = []
        for mes in dados_consolidados:
            dados = dados_consolidados[mes]
            linha = [mes]
            linha.append(formatar_valor(dados.get(f'top20_{tipo}', None), fmt))
            linha.append(formatar_valor(dados.get(f'resto_{tipo}', None), fmt))
            tabela_dados.append(linha)
        
        tabela = ax_table.table(
            cellText=tabela_dados,
            colLabels=['Mês', 'Top 20', 'Resto'],
            loc='center',
            cellLoc='center'
        )
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(FONT_SIZE)
        tabela.scale(1, 1.5)
        
        for (i, j), cell in tabela.get_celld().items():
            if i == 0:
                cell.set_text_props(fontweight='bold')
                cell.set_facecolor('#e6e6e6')
        
        # Gráfico com espaço superior
        ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
        
        top20 = [dados_consolidados[mes].get(f'top20_{tipo}', 0) for mes in meses]
        resto = [dados_consolidados[mes].get(f'resto_{tipo}', 0) for mes in meses]
        
        bar_width = 0.35
        x = range(len(meses))
        
        ax_graph.bar(x, top20, bar_width, label='Top 20', color='#1f77b4')
        ax_graph.bar([i + bar_width for i in x], resto, bar_width, label='Resto', color='#ff7f0e')
        
        ax_graph.set_title(f'{titulo} - Comparativo Top 20 vs Resto', fontweight='bold')
        ax_graph.set_ylabel(unidade)
        ax_graph.set_xticks([i + bar_width/2 for i in x])
        ax_graph.set_xticklabels(meses)
        ax_graph.legend()
        ax_graph.grid(True, linestyle='--', alpha=0.7, axis='y')
        
        # Adicionar valores nas barras
        for i in x:
            ax_graph.text(i, top20[i], formatar_valor(top20[i], fmt), 
                         ha='center', va='bottom', fontsize=FONT_SIZE-1)
            ax_graph.text(i + bar_width, resto[i], formatar_valor(resto[i], fmt), 
                         ha='center', va='bottom', fontsize=FONT_SIZE-1)
        
        # Ajustar limite superior
        ylim = ax_graph.get_ylim()
        ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
        
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

# Exibir apenas o JSON consolidado e mensagem final
print(json.dumps(dados_consolidados, indent=4, ensure_ascii=False))
print(f"\nRelatório gerado com sucesso em: {output_pdf}")