import os
import json
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import PyPDF2
import re

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
    
    # Extração básica (mantida igual)
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
    
    # SOLUÇÃO DEFINITIVA PARA FATURAMENTO
    # Padrão para: "T otal: R$6.769.219,15T op 20 produtos vs Resto - Faturamento"
    padrao_faturamento = r"T otal:\s*R\$([\d\.]+,\d+)T op 20 produtos vs Resto - Faturamento"
    
    match_faturamento = re.search(padrao_faturamento, texto)
    if match_faturamento:
        try:
            valor = match_faturamento.group(1).replace('.', '').replace(',', '.')
            dados['total_faturamento'] = float(valor)
            print(f"DEBUG - Faturamento extraído com sucesso: {valor}")
        except Exception as e:
            print(f"Erro ao converter faturamento: {e}")

    # Padrões para tonelagem e margem (mantidos)
    padrao_tonelagem = r"T otal:\s*([\d\.,]+)T op 20 produtos vs Resto - T onelagem"
    padrao_margem = r"T otal:\s*([\d\.]+)%T op 20 produtos vs Resto - Margem"
    
    if match_tonelagem := re.search(padrao_tonelagem, texto):
        dados['total_tonelagem'] = float(match_tonelagem.group(1).replace('.', '').replace(',', '.'))
    
    if match_margem := re.search(padrao_margem, texto):
        dados['total_margem'] = float(match_margem.group(1))
    
    
    # PADRÕES CORRIGIDOS PARA TOP 20 VS RESTO
    # Tonelagem (já funcionando)
    padrao_top20_ton = r"\((\d+\.\d+,\d+)\)\d+\.\d+%"
    padrao_resto_ton = r"\((\d+\.\d+,\d+)\)T otal:"
    
    # FATURAMENTO CORRIGIDO - padrão para (R$X.XXX,XX)YY.YY%
    padrao_top20_fat = r"\(R\$([\d\.]+,\d+)\)\d+\.\d+%"
    padrao_resto_fat = r"\(R\$([\d\.]+,\d+)\)T otal:"
    
    # Margem (já funcionando)
    padrao_top20_mar = r"\((\d+\.\d+)%\)\d+\.\d+%"
    padrao_resto_mar = r"\((\d+\.\d+)%\)T otal:"
    
    # Extração dos novos dados
    try:
        # Tonelagem (mantido)
        if match := re.search(padrao_top20_ton, texto):
            dados['top20_ton'] = float(match.group(1).replace('.', '').replace(',', '.'))
        if match := re.search(padrao_resto_ton, texto):
            dados['resto_ton'] = float(match.group(1).replace('.', '').replace(',', '.'))
        
        # FATURAMENTO - CORREÇÃO PRINCIPAL
        if match := re.search(padrao_top20_fat, texto):
            valor = match.group(1).replace('.', '').replace(',', '.')
            dados['top20_fat'] = float(valor)
            print(f"DEBUG - Top20 Faturamento extraído: {valor}")
            
        if match := re.search(padrao_resto_fat, texto):
            valor = match.group(1).replace('.', '').replace(',', '.')
            dados['resto_fat'] = float(valor)
            print(f"DEBUG - Resto Faturamento extraído: {valor}")
        
        # Margem (mantido)
        if match := re.search(padrao_top20_mar, texto):
            dados['top20_mar'] = float(match.group(1))
        if match := re.search(padrao_resto_mar, texto):
            dados['resto_mar'] = float(match.group(1))
            
    except Exception as e:
        print(f"Erro ao extrair dados Top20/Resto: {e}")
    
    return dados

# Processar cada arquivo PDF
for mes in meses:
    nome_arquivo = f"Relatório Analítico de Vendas - Geral - {mes} {ano}.pdf"
    caminho_pdf = os.path.join(caminho_base, mes, nome_arquivo)
    
    if os.path.exists(caminho_pdf):
        texto = extrair_texto_pdf(caminho_pdf)
        print(f"\n--- Conteúdo extraído do PDF ({mes}) ---")
        print(texto[:1000] + "...")  # Mostra apenas o início do conteúdo para verificação
        
        dados = extrair_dados(texto)
        dados_consolidados[mes] = dados
        print(f"\nDados extraídos para {mes}:")
        print(json.dumps(dados, indent=4, ensure_ascii=False))
    else:
        print(f"Arquivo não encontrado: {caminho_pdf}")

# Exibir dados no console em formato JSON
print("\nDados consolidados extraídos dos relatórios:")
print(json.dumps(dados_consolidados, indent=4, ensure_ascii=False))

# Criar DataFrames para análise
df_quantidades = pd.DataFrame.from_dict({mes: {k: v for k, v in dados.items() if k.startswith('Qtde')} 
                                      for mes, dados in dados_consolidados.items()}, orient='index')
df_quantidades.index.name = 'Mês'
df_quantidades = df_quantidades.reset_index()

# Verificar se todas as colunas de totais existem nos dados
colunas_totais = ['total_tonelagem', 'total_faturamento', 'total_margem']
for mes in dados_consolidados:
    for coluna in colunas_totais:
        if coluna not in dados_consolidados[mes]:
            dados_consolidados[mes][coluna] = None  # Preencher com None se não existir

df_totais = pd.DataFrame.from_dict({mes: {k: v for k, v in dados.items() if k.startswith('total')} 
                                  for mes, dados in dados_consolidados.items()}, orient='index')
df_totais.index.name = 'Mês'
df_totais = df_totais.reset_index()

# Configurações de layout padronizadas
PAGE_WIDTH = 11.69  # Largura A4 em polegadas
PAGE_HEIGHT = 8.27   # Altura A4 em polegadas
FONT_SIZE = 10
DPI = 300

# Criar PDF com os resultados
with PdfPages(output_pdf) as pdf:
    # Configuração geral do estilo
    plt.rcParams.update({
        'font.size': FONT_SIZE,
        'figure.dpi': DPI,
        'figure.autolayout': True  # Ajusta automaticamente o layout
    })
    
    # 1. Página de título
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.6, 'Relatório Consolidado de Vendas', 
             ha='center', va='center', fontsize=16, fontweight='bold')
    plt.text(0.5, 0.5, f'Período: Maio a Junho de {ano}', 
             ha='center', va='center', fontsize=12)
    plt.axis('off')
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)
    
    # 2. Seção de Quantidades
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '1. Análise de Quantidades', 
             ha='center', va='center', fontsize=14, fontweight='bold')
    plt.axis('off')
    
    # Tabela de quantidades
    ax_table = fig.add_subplot(211)
    ax_table.axis('tight')
    ax_table.axis('off')
    
    tabela_dados = df_quantidades[['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos']].values
    
    tabela = ax_table.table(cellText=tabela_dados,
                      colLabels=['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos'],
                      loc='center',
                      cellLoc='center')
    
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(FONT_SIZE)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#f2f2f2')
    
    # Gráfico de quantidades
    ax_graph = fig.add_subplot(212)
    
    df_quantidades['Qtde Clientes'] = df_quantidades['Qtde Clientes'].astype(int)
    df_quantidades['Qtde Vendedores'] = df_quantidades['Qtde Vendedores'].astype(int)
    df_quantidades['Qtde Produtos'] = df_quantidades['Qtde Produtos'].astype(int)
    
    ax_graph.plot(df_quantidades['Mês'], df_quantidades['Qtde Clientes'], marker='o', label='Clientes')
    ax_graph.plot(df_quantidades['Mês'], df_quantidades['Qtde Vendedores'], marker='s', label='Vendedores')
    ax_graph.plot(df_quantidades['Mês'], df_quantidades['Qtde Produtos'], marker='^', label='Produtos')
    
    ax_graph.set_title('Evolução Mensal das Quantidades', fontweight='bold')
    ax_graph.set_xlabel('Mês')
    ax_graph.set_ylabel('Quantidade')
    ax_graph.grid(True, linestyle='--', alpha=0.7)
    ax_graph.legend()
    
    for i, row in df_quantidades.iterrows():
        ax_graph.text(row['Mês'], row['Qtde Clientes'], str(row['Qtde Clientes']), 
                 ha='center', va='bottom')
        ax_graph.text(row['Mês'], row['Qtde Vendedores'], str(row['Qtde Vendedores']), 
                 ha='center', va='bottom')
        ax_graph.text(row['Mês'], row['Qtde Produtos'], str(row['Qtde Produtos']), 
                 ha='center', va='bottom')
    
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)
    
    # 3. Seção de Totais
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '2. Análise de Totais', 
             ha='center', va='center', fontsize=14, fontweight='bold')
    plt.axis('off')
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)

    # Função para formatar os valores
    def formatar_valor(valor, tipo):
        if pd.isna(valor):
            return "N/A"
        if tipo == 'tonelagem':
            return f"{valor:,.3f} kg".replace(',', 'X').replace('.', ',').replace('X', '.')
        elif tipo == 'faturamento':
            return f"R${valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        elif tipo == 'margem':
            return f"{valor:,.2f}%".replace('.', ',')
        return str(valor)

    # Tabela de totais
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    ax = fig.add_subplot(111)
    ax.axis('tight')
    ax.axis('off')
    
    # Preparar dados para tabela com valores formatados
    tabela_dados = []
    for _, row in df_totais.iterrows():
        linha = [row['Mês']]
        if 'total_tonelagem' in df_totais.columns:
            linha.append(formatar_valor(row['total_tonelagem'], 'tonelagem'))
        if 'total_faturamento' in df_totais.columns:
            linha.append(formatar_valor(row['total_faturamento'], 'faturamento'))
        if 'total_margem' in df_totais.columns:
            linha.append(formatar_valor(row['total_margem'], 'margem'))
        tabela_dados.append(linha)
    
    # Criar rótulos das colunas
    colunas = ['Mês']
    if 'total_tonelagem' in df_totais.columns:
        colunas.append('Tonelagem (kg)')
    if 'total_faturamento' in df_totais.columns:
        colunas.append('Faturamento')
    if 'total_margem' in df_totais.columns:
        colunas.append('Margem (%)')
    
    # Criar tabela
    tabela = ax.table(cellText=tabela_dados,
                      colLabels=colunas,
                      loc='center',
                      cellLoc='center')
    
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(FONT_SIZE)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#e6e6e6')
    
    plt.title('Totais por Mês', y=1.08, fontweight='bold')
    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)

    # Gráficos individuais para cada métrica
    for i, (col, titulo, unidade, tipo) in enumerate([
        ('total_tonelagem', 'Tonelagem Total', 'kg', 'tonelagem'),
        ('total_faturamento', 'Faturamento Total', 'R$', 'faturamento'),
        ('total_margem', 'Margem Total', '%', 'margem')
    ]):
        if col in df_totais.columns:
            fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
            ax = fig.add_subplot(111)
            
            valores = df_totais[col]
            barras = ax.bar(df_totais['Mês'], valores, color=['#1f77b4', '#2ca02c', '#ff7f0e'][i % 3])
            
            ax.set_title(titulo, fontweight='bold')
            ax.set_xlabel('Mês')
            ax.set_ylabel(unidade)
            ax.grid(True, linestyle='--', alpha=0.7, axis='y')
            
            # Adicionar valores nas barras
            for barra in barras:
                height = barra.get_height()
                ax.text(barra.get_x() + barra.get_width()/2., height,
                        formatar_valor(height, tipo),
                        ha='center', va='bottom')
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
            plt.close(fig)
        
    # 4. Seção Top 20 vs Resto
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '3. Análise Top 20 Produtos vs Resto', 
             ha='center', va='center', fontsize=14, fontweight='bold')
    plt.axis('off')
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)

    # Tabela comparativa
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    ax = fig.add_subplot(111)
    ax.axis('off')
    
    # Preparar dados
    cabecalho = [
        'Mês',
        'Tonelagem (kg)', '',  # Título e espaço para Top20/Resto
        'Faturamento', '',     # Título e espaço para Top20/Resto
        'Margem (%)', ''       # Título e espaço para Top20/Resto
    ]
    
    subcabecalho = [
        '',  # Espaço para o mês
        'Top 20', 'Resto',  # Tonelagem
        'Top 20', 'Resto',  # Faturamento
        'Top 20', 'Resto'   # Margem
    ]
    
    linhas = []
    for mes in dados_consolidados:
        dados = dados_consolidados[mes]
        linha = [mes]
        
        # Tonelagem
        if 'top20_ton' in dados and 'resto_ton' in dados:
            linha.extend([
                formatar_valor(dados['top20_ton'], 'tonelagem'),
                formatar_valor(dados['resto_ton'], 'tonelagem')
            ])
        else:
            linha.extend(['-', '-'])
        
        # Faturamento
        if 'top20_fat' in dados and 'resto_fat' in dados:
            linha.extend([
                formatar_valor(dados['top20_fat'], 'faturamento'),
                formatar_valor(dados['resto_fat'], 'faturamento')
            ])
        else:
            linha.extend(['-', '-'])
        
        # Margem
        if 'top20_mar' in dados and 'resto_mar' in dados:
            linha.extend([
                formatar_valor(dados['top20_mar'], 'margem'),
                formatar_valor(dados['resto_mar'], 'margem')
            ])
        else:
            linha.extend(['-', '-'])
        
        linhas.append(linha)

    # Criar tabela com duas linhas de cabeçalho
    tabela = ax.table(
        cellText=[subcabecalho] + linhas,
        colLabels=None,
        loc='center',
        cellLoc='center'
    )

    # Adicionar títulos das categorias manualmente
    ax.text(0.28, 0.93, 'Tonelagem (kg)', fontweight='bold', 
            ha='center', transform=ax.transAxes)
    ax.text(0.56, 0.93, 'Faturamento', fontweight='bold', 
            ha='center', transform=ax.transAxes)
    ax.text(0.84, 0.93, 'Margem (%)', fontweight='bold', 
            ha='center', transform=ax.transAxes)

    # Estilização
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(FONT_SIZE)
    tabela.scale(1, 1.5)

    # Estilo para cabeçalhos
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:  # Linha de subcabeçalho
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#e6e6e6')
        elif j == 0:  # Coluna de meses
            cell.set_facecolor('#f9f9f9')

    plt.tight_layout()
    pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
    plt.close(fig)
    
    # Gráficos comparativos
    for tipo, titulo, unidade, fmt in [
        ('ton', 'Tonelagem - Top 20 vs Resto', 'kg', 'tonelagem'),
        ('fat', 'Faturamento - Top 20 vs Resto', 'R$', 'faturamento'),
        ('mar', 'Margem - Top 20 vs Resto', '%', 'margem')
    ]:
        if f'top20_{tipo}' in dados_consolidados[meses[0]]:
            fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
            ax = fig.add_subplot(111)
            
            # Preparar dados
            labels = meses
            top20 = [dados_consolidados[mes][f'top20_{tipo}'] for mes in meses]
            resto = [dados_consolidados[mes][f'resto_{tipo}'] for mes in meses]
            
            # Gráfico de barras agrupadas
            bar_width = 0.35
            x = range(len(meses))
            
            ax.bar(x, top20, bar_width, label='Top 20', color='#1f77b4')
            ax.bar([i + bar_width for i in x], resto, bar_width, label='Resto', color='#ff7f0e')
            
            ax.set_title(titulo, fontweight='bold')
            ax.set_xlabel('Mês')
            ax.set_ylabel(unidade)
            ax.set_xticks([i + bar_width/2 for i in x])
            ax.set_xticklabels(meses)
            ax.legend()
            ax.grid(True, linestyle='--', alpha=0.7, axis='y')
            
            # Adicionar valores
            for i in x:
                ax.text(i, top20[i], formatar_valor(top20[i], fmt), ha='center', va='bottom')
                ax.text(i + bar_width, resto[i], formatar_valor(resto[i], fmt), ha='center', va='bottom')
            
            plt.tight_layout()
            pdf.savefig(fig, bbox_inches='tight', dpi=DPI)
            plt.close(fig)

print(f"\nRelatório completo gerado com sucesso em: {output_pdf}")