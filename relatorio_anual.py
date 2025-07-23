import os
import camelot
import warnings
import json
import re
import os
import json
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import numpy as np
import camelot
import warnings

# Ignorar warnings (opcional)
warnings.filterwarnings("ignore")

# Dicionário para armazenar os resultados
resultados = {}

# Caminho base
ano = "2025"
base_path = r"C:\Users\win11\OneDrive\Documentos\Relatórios Analítico de Vendas\2025"
meses = ["Maio", "Junho", "Julho"]

def limpar_numero(texto, is_margem=False):
    """Converte strings numéricas com símbolos em floats"""
    if not texto:
        return 0.0
    
    if is_margem:
        # Para margens, apenas remove o % e converte (já está com ponto decimal)
        texto = texto.replace('%', '').strip()
        try:
            return float(texto)
        except ValueError:
            return 0.0
    else:
        # Para outros valores, remove R$, pontos de milhar e substitui vírgula decimal
        texto = texto.replace('R$', '').replace('.', '').replace(',', '.').strip()
        try:
            return float(texto)
        except ValueError:
            return 0.0

def extrair_valores(tabela_df):
    """Extrai os valores da tabela formatada com base na estrutura específica"""
    dados = {}
    texto_tabela = tabela_df.to_string(index=False, header=False)
    
    # Extrair estatísticas gerais (valores inteiros)
    dados['qtde_clientes'] = int(re.search(r'Qtde Clientes\s+(\d+)', texto_tabela).group(1))
    dados['qtde_vendedores'] = int(re.search(r'Qtde Vendedores\s+(\d+)', texto_tabela).group(1))
    dados['qtde_produtos'] = int(re.search(r'Qtde Produtos\s+(\d+)', texto_tabela).group(1))
    
    # Função auxiliar para extrair valores entre parênteses
    def extrair_parenteses(texto, padrao):
        matches = re.findall(padrao + r'.*?\(([^)]+)\)', texto_tabela, re.DOTALL)
        return matches[-1] if matches else None
    
    # Extrair tonelagem (3 decimais)
    tonelagem_total = re.search(r'Tonelagem.*?Total:\s*([\d.,]+)', texto_tabela, re.DOTALL)
    if tonelagem_total:
        # Primeiro valor entre parênteses após porcentagem é Outros
        outros_ton = extrair_parenteses(texto_tabela, r'Tonelagem.*?\d+\.\d+%')
        # Segundo valor entre parênteses após porcentagem é Top 20
        top20_ton = extrair_parenteses(texto_tabela, r'Tonelagem.*?\d+\.\d+%.*?\d+\.\d+%')
        
        dados['tonelagem'] = {
            'total': round(limpar_numero(tonelagem_total.group(1)), 3),
            'top_20': round(limpar_numero(top20_ton) if top20_ton else 0.0, 3),
            'outros': round(limpar_numero(outros_ton) if outros_ton else 0.0, 3)
        }
    
    # Extrair faturamento (2 decimais)
    faturamento_total = re.search(r'Faturamento.*?Total:\s*([R\$ \d.,]+)', texto_tabela, re.DOTALL)
    if faturamento_total:
        # Primeiro valor entre parênteses após porcentagem é Outros
        outros_fat = extrair_parenteses(texto_tabela, r'Faturamento.*?\d+\.\d+%')
        # Segundo valor entre parênteses após porcentagem é Top 20
        top20_fat = extrair_parenteses(texto_tabela, r'Faturamento.*?\d+\.\d+%.*?\d+\.\d+%')
        
        dados['faturamento'] = {
            'total': round(limpar_numero(faturamento_total.group(1)), 2),
            'top_20': round(limpar_numero(top20_fat) if top20_fat else 0.0, 2),
            'outros': round(limpar_numero(outros_fat) if outros_fat else 0.0, 2)
        }
    
    # Extrair margem (valores já estão com ponto decimal)
    margem_total = re.search(r'Margem.*?Total:\s*([\d.,]+%)', texto_tabela, re.DOTALL)
    if margem_total:
        # Encontrar todos os valores de margem com seus espaçamentos
        margens = re.findall(r'(\s+)\(([\d.,]+%)\)', texto_tabela.split('Margem')[-1])
        
        if len(margens) >= 2:
            # Ordenar pelo espaçamento - menos espaços = top_20, mais espaços = outros
            margens.sort(key=lambda x: len(x[0]))
            
            dados['margem'] = {
                'total': round(limpar_numero(margem_total.group(1), True), 2),
                'top_20': round(limpar_numero(margens[0][1], True), 2),
                'outros': round(limpar_numero(margens[1][1], True), 2)
            }
    
    return dados

for mes in meses:
    pdf_path = os.path.join(base_path, mes, f"Relatório Analítico de Vendas - Geral - {mes} 2025.pdf")
    
    if os.path.exists(pdf_path):
        print(f"\n=== Extraindo tabelas de {mes} ===")
        
        try:
            tables = camelot.read_pdf(
                pdf_path,
                pages="all",
                flavor="stream",
                strip_text="\n"
            )
            
            dados_mes = {}
            for table in tables:
                if "Estatísticas Gerais de Vendas" in table.df.to_string():
                    dados_mes = extrair_valores(table.df)
                    resultados[mes] = dados_mes
                    break
            
        except Exception as e:
            print(f"Erro ao extrair tabelas: {e}")
    else:
        print(f"Arquivo não encontrado: {pdf_path}")

# Exibir o resultado em JSON
print("\n=== Resultado em JSON ===")
print(json.dumps(resultados, indent=4, ensure_ascii=False))

# Supondo que seu código original populou a variável 'resultados' com os dados
dados = resultados

# Agora processamos para criar o PDF
primeiro_mes = meses[0]
ultimo_mes = meses[-1]

# Caminho para salvar o PDF
downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
pdf_path = os.path.join(downloads_path, f"Relatório Analítico de Vendas {primeiro_mes} a {ultimo_mes} de {ano}.pdf")

# Criar o PDF
with PdfPages(pdf_path) as pdf:
    # Configurações gerais
    plt.rcParams['font.size'] = 10
    
    # Página 1 - Cabeçalho
    fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4
    fig.patch.set_visible(False)
    ax.axis('off')
    
    # Título principal
    titulo = f"Relatório Analítico de Vendas\n{primeiro_mes} a {ultimo_mes} de {ano}"
    ax.text(0.5, 0.9, titulo, ha='center', va='center', fontsize=16, fontweight='bold')
    
    # Data e hora de geração
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ax.text(0.5, 0.85, f"Gerado em: {data_hora}", ha='center', va='center', fontsize=10)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Página 2 - Quantidades
    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    fig.patch.set_visible(False)
    ax.axis('off')
    
    # Título da seção
    ax.text(0.1, 0.95, "1. Quantidades", fontsize=14, fontweight='bold')
    
    # Tabela de quantidades
    col_labels = meses
    row_labels = ['Clientes', 'Vendedores', 'Produtos']
    
    # Preparar dados da tabela diretamente do JSON gerado
    table_data = [
        [dados[mes]['qtde_clientes'] for mes in meses if mes in dados],
        [dados[mes]['qtde_vendedores'] for mes in meses if mes in dados],
        [dados[mes]['qtde_produtos'] for mes in meses if mes in dados]
    ]
    
    # Posicionar tabela
    table = ax.table(cellText=table_data, 
                    rowLabels=row_labels, 
                    colLabels=col_labels,
                    loc='center', 
                    cellLoc='center',
                    bbox=[0.1, 0.6, 0.8, 0.3])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)
    
    # Gráfico de linhas
    ax2 = fig.add_axes([0.1, 0.1, 0.8, 0.4])
    for i, row in enumerate(table_data):
        ax2.plot(col_labels, row, marker='o', label=row_labels[i])
    
    ax2.set_title('Evolução das Quantidades')
    ax2.set_ylabel('Quantidade')
    ax2.legend()
    ax2.grid(True, linestyle='--', alpha=0.7)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Página 3 - Totais
    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    fig.patch.set_visible(False)
    ax.axis('off')
    
    # Título da seção
    ax.text(0.1, 0.95, "2. Totais", fontsize=14, fontweight='bold')
    
    # Tabela de totais
    col_labels = [mes for mes in meses if mes in dados]
    row_labels = ['Tonelagem (kg)', 'Faturamento (R$)', 'Margem (%)']
    
    # Preparar dados da tabela diretamente do JSON
    table_data = [
        [dados[mes]['tonelagem']['total'] for mes in meses if mes in dados],
        [dados[mes]['faturamento']['total'] for mes in meses if mes in dados],
        [dados[mes]['margem']['total'] for mes in meses if mes in dados]
    ]
    
    # Formatando os valores para exibição
    formatted_data = [
        [f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in table_data[0]],
        [f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in table_data[1]],
        [f"{x:.2f}%" for x in table_data[2]]
    ]
    
    # Posicionar tabela
    table = ax.table(cellText=formatted_data, 
                    rowLabels=row_labels, 
                    colLabels=col_labels,
                    loc='center', 
                    cellLoc='center',
                    bbox=[0.1, 0.7, 0.8, 0.2])
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.5)
    
    # Gráficos de barras
    ax1 = fig.add_axes([0.1, 0.45, 0.8, 0.2])
    ax1.bar(col_labels, table_data[0], color='skyblue')
    ax1.set_title('Tonelagem Total (kg)')
    ax1.grid(True, linestyle='--', alpha=0.7)
    
    ax2 = fig.add_axes([0.1, 0.2, 0.8, 0.2])
    ax2.bar(col_labels, table_data[1], color='lightgreen')
    ax2.set_title('Faturamento Total (R$)')
    ax2.grid(True, linestyle='--', alpha=0.7)
    
    ax3 = fig.add_axes([0.1, -0.05, 0.8, 0.2])
    ax3.bar(col_labels, table_data[2], color='salmon')
    ax3.set_title('Margem Total (%)')
    ax3.grid(True, linestyle='--', alpha=0.7)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Página 4 - Top 20 vs Outros
    fig, ax = plt.subplots(figsize=(8.27, 11.69))
    fig.patch.set_visible(False)
    ax.axis('off')
    
    # Título da seção
    ax.text(0.1, 0.95, "3. Top 20 vs Outros", fontsize=14, fontweight='bold')
    
    # Tabela para Tonelagem
    ax.text(0.1, 0.9, "Tonelagem (kg)", fontsize=12, fontweight='bold')
    ton_data = [
        [dados[mes]['tonelagem']['top_20'] for mes in meses if mes in dados],
        [dados[mes]['tonelagem']['outros'] for mes in meses if mes in dados]
    ]
    ton_formatted = [
        [f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in ton_data[0]],
        [f"{x:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in ton_data[1]]
    ]
    
    table1 = ax.table(cellText=ton_formatted, 
                     rowLabels=['Top 20', 'Outros'], 
                     colLabels=col_labels,
                     loc='center', 
                     cellLoc='center',
                     bbox=[0.1, 0.75, 0.8, 0.1])
    table1.auto_set_font_size(False)
    table1.set_fontsize(8)
    
    # Tabela para Faturamento
    ax.text(0.1, 0.65, "Faturamento (R$)", fontsize=12, fontweight='bold')
    fat_data = [
        [dados[mes]['faturamento']['top_20'] for mes in meses if mes in dados],
        [dados[mes]['faturamento']['outros'] for mes in meses if mes in dados]
    ]
    fat_formatted = [
        [f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in fat_data[0]],
        [f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in fat_data[1]]
    ]
    
    table2 = ax.table(cellText=fat_formatted, 
                     rowLabels=['Top 20', 'Outros'], 
                     colLabels=col_labels,
                     loc='center', 
                     cellLoc='center',
                     bbox=[0.1, 0.5, 0.8, 0.1])
    table2.auto_set_font_size(False)
    table2.set_fontsize(8)
    
    # Tabela para Margem
    ax.text(0.1, 0.4, "Margem (%)", fontsize=12, fontweight='bold')
    marg_data = [
        [dados[mes]['margem']['top_20'] for mes in meses if mes in dados],
        [dados[mes]['margem']['outros'] for mes in meses if mes in dados]
    ]
    marg_formatted = [
        [f"{x:.2f}%" for x in marg_data[0]],
        [f"{x:.2f}%" for x in marg_data[1]]
    ]
    
    table3 = ax.table(cellText=marg_formatted, 
                     rowLabels=['Top 20', 'Outros'], 
                     colLabels=col_labels,
                     loc='center', 
                     cellLoc='center',
                     bbox=[0.1, 0.25, 0.8, 0.1])
    table3.auto_set_font_size(False)
    table3.set_fontsize(8)
    
    # Gráficos
    # Tonelagem
    ax1 = fig.add_axes([0.1, 0.15, 0.25, 0.08])
    width = 0.35
    x = np.arange(len(col_labels))
    ax1.bar(x - width/2, ton_data[0], width, label='Top 20')
    ax1.bar(x + width/2, ton_data[1], width, label='Outros')
    ax1.set_title('Tonelagem')
    ax1.set_xticks(x)
    ax1.set_xticklabels(col_labels, rotation=45, ha='right')
    ax1.legend(fontsize=6)
    
    # Faturamento
    ax2 = fig.add_axes([0.4, 0.15, 0.25, 0.08])
    ax2.bar(x - width/2, fat_data[0], width, label='Top 20')
    ax2.bar(x + width/2, fat_data[1], width, label='Outros')
    ax2.set_title('Faturamento')
    ax2.set_xticks(x)
    ax2.set_xticklabels(col_labels, rotation=45, ha='right')
    ax2.legend(fontsize=6)
    
    # Margem
    ax3 = fig.add_axes([0.7, 0.15, 0.25, 0.08])
    ax3.bar(x - width/2, marg_data[0], width, label='Top 20')
    ax3.bar(x + width/2, marg_data[1], width, label='Outros')
    ax3.set_title('Margem')
    ax3.set_xticks(x)
    ax3.set_xticklabels(col_labels, rotation=45, ha='right')
    ax3.legend(fontsize=6)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

print(f"Relatório gerado com sucesso em: {pdf_path}")