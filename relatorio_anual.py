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

# Configurações padrão
plt.rcParams.update({
    'figure.constrained_layout.use': True,
    'figure.constrained_layout.h_pad': 0.5,
    'figure.constrained_layout.w_pad': 0.5,
    'figure.constrained_layout.hspace': 0.2,
    'figure.constrained_layout.wspace': 0.2
})  

# Tamanhos padrão
fig_size = (8.27, 11.69) 
graph_size = (6, 3) 
table_size = (6, 2) 

# Margens padrão
left_margin = 0.1
right_margin = 0.9
bottom_margin = 0.1
top_margin = 0.9

# Criar o PDF
with PdfPages(pdf_path) as pdf:
    # Página 1 - Cabeçalho
    fig = plt.figure(figsize=fig_size)
    fig.patch.set_visible(False)
    ax = fig.add_subplot(111)
    ax.axis('off')
    
    # Título principal centralizado
    titulo = f"Relatório Analítico de Vendas\n{primeiro_mes} a {ultimo_mes} de {ano}"
    ax.text(0.5, 0.6, titulo, ha='center', va='center', fontsize=16, fontweight='bold')
    
    # Data e hora de geração
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    ax.text(0.5, 0.4, f"Gerado em: {data_hora}", ha='center', va='center', fontsize=10)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Função para criar uma nova página com layout padronizado
    def nova_pagina(titulo_principal, is_continuacao=False, num_itens=3):
        fig = plt.figure(figsize=fig_size, constrained_layout=True)
        if is_continuacao:
            fig.suptitle(f"{titulo_principal} (continuação)", fontweight='bold', y=0.98)
        else:
            fig.suptitle(titulo_principal, fontweight='bold', y=0.98)
        
        # Ajusta o espaçamento baseado no número de itens
        if num_itens < 3:
            fig.subplots_adjust(hspace=0.5)  # Aumenta o espaçamento vertical
        
        # Adiciona subplots vazios para manter o mesmo tamanho de página
        for i in range(num_itens, 3):
            ax = fig.add_subplot(3, 1, i+1)
            ax.axis('off')
        
        return fig
    
    # Função add_tabela atualizada
    def add_tabela(fig, posicao, dados, titulo, row_labels, col_labels, is_money=False, is_percent=False):
        ax = fig.add_subplot(3, 1, posicao)
        ax.axis('off')
        ax.set_title(titulo, pad=10)  # Adiciona padding ao título
        
        # Formata os dados (mantido igual)
        formatted_data = []
        for row in dados:
            formatted_row = []
            for val in row:
                if is_money:
                    formatted_row.append(f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                elif is_percent:
                    formatted_row.append(f"{val:.2f}%")
                elif isinstance(val, float):
                    if 'Tonelagem' in titulo:
                        formatted_row.append(f"{val:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    else:
                        formatted_row.append(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                else:
                    formatted_row.append(f"{val:,}".replace(",", "."))
            formatted_data.append(formatted_row)
        
        # Ajuste no bbox da tabela
        table = ax.table(cellText=formatted_data, 
                       rowLabels=row_labels, 
                       colLabels=col_labels,
                       loc='center', 
                       cellLoc='center',
                       bbox=[0.1, 0.1, 0.8, 0.7])  # Ajuste na altura (0.7 em vez de 0.8)
        
        table.auto_set_font_size(False)
        table.set_fontsize(8)  # Fonte um pouco menor
        table.scale(1, 1.3)   # Escala mais compacta
        
        return fig
    
    # Função add_grafico atualizada
    def add_grafico(fig, posicao, dados, titulo, cor, is_percent=False):
        ax = fig.add_subplot(3, 1, posicao)
        ax.clear()  # Limpa qualquer conteúdo anterior
        
        # Restante do código mantido igual
        x = np.arange(len(meses_validas))
        width = 0.6
        
        bars = ax.bar(x, dados, width, color=cor)
        ax.set_title(titulo, pad=10)  # Adiciona padding ao título
        ax.set_xticks(x)
        ax.set_xticklabels(meses_validas)
        ax.grid(True, linestyle='--', alpha=0.7)
        
        # Adicionar valores nas barras (mantido igual)
        for bar in bars:
            height = bar.get_height()
            if is_percent:
                texto = f"{height:.2f}%"
            elif isinstance(height, float):
                if titulo.startswith('Tonelagem'):
                    texto = f"{height:,.3f}".replace(",", "X").replace(".", ",").replace("X", ".")
                else:
                    texto = f"{height:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                texto = f"{height:,}".replace(",", ".")
            
            ax.text(bar.get_x() + bar.get_width()/2., height,
                    texto, ha='center', va='bottom', fontsize=7)  # Fonte menor
        
        return fig
    
    # Lista de meses com dados válidos
    meses_validas = [mes for mes in meses if mes in dados]
    
    # ========== PÁGINA 2 ==========
    fig = nova_pagina("1. Quantidades")
    
    # Tabela de quantidades
    table_data = [
        [dados[mes]['qtde_clientes'] for mes in meses_validas],
        [dados[mes]['qtde_vendedores'] for mes in meses_validas],
        [dados[mes]['qtde_produtos'] for mes in meses_validas]
    ]
    fig = add_tabela(fig, 1, table_data, "Quantidades", 
                    ['Clientes', 'Vendedores', 'Produtos'], meses_validas)
    
    # Gráfico de clientes
    fig = add_grafico(fig, 2, table_data[0], "Quantidade de Clientes", 'skyblue')
    
    # Gráfico de vendedores
    fig = add_grafico(fig, 3, table_data[1], "Quantidade de Vendedores", 'lightgreen')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 3 ==========
    fig = nova_pagina("1. Quantidades", is_continuacao=True, num_itens=1)

    # Gráfico de produtos
    fig = add_grafico(fig, 1, table_data[2], "Quantidade de Produtos", 'salmon')

    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 4 ==========
    fig = nova_pagina("2. Totais")
    
    # Tabela de totais
    table_data = [
        [dados[mes]['tonelagem']['total'] for mes in meses_validas],
        [dados[mes]['faturamento']['total'] for mes in meses_validas],
        [dados[mes]['margem']['total'] for mes in meses_validas]
    ]
    fig = add_tabela(fig, 1, table_data, "Totais", 
                    ['Tonelagem (kg)', 'Faturamento (R$)', 'Margem (%)'], 
                    meses_validas)
    
    # Gráfico de tonelagem
    fig = add_grafico(fig, 2, table_data[0], "Tonelagem Total (kg)", 'skyblue')
    
    # Gráfico de faturamento
    fig = add_grafico(fig, 3, table_data[1], "Faturamento Total (R$)", 'lightgreen')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 5 ==========
    fig = nova_pagina("2. Totais", is_continuacao=True, num_itens=1)

    # Gráfico de margem
    fig = add_grafico(fig, 1, table_data[2], "Margem Total (%)", 'salmon', is_percent=True)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 6 ==========
    fig = nova_pagina("3. Top 20 vs Outros")
    
    # Tabela tonelagem (Top 20 vs Outros)
    table_data_ton = [
        [dados[mes]['tonelagem']['top_20'] for mes in meses_validas],
        [dados[mes]['tonelagem']['outros'] for mes in meses_validas]
    ]
    fig = add_tabela(fig, 1, table_data_ton, "Tonelagem (kg) - Top 20 vs Outros", 
                    ['Top 20', 'Outros'], meses_validas)
    
    # Tabela faturamento (Top 20 vs Outros)
    table_data_fat = [
        [dados[mes]['faturamento']['top_20'] for mes in meses_validas],
        [dados[mes]['faturamento']['outros'] for mes in meses_validas]
    ]
    fig = add_tabela(fig, 2, table_data_fat, "Faturamento (R$) - Top 20 vs Outros", 
                    ['Top 20', 'Outros'], meses_validas, is_money=True)
    
    # Tabela margem (Top 20 vs Outros)
    table_data_marg = [
        [dados[mes]['margem']['top_20'] for mes in meses_validas],
        [dados[mes]['margem']['outros'] for mes in meses_validas]
    ]
    fig = add_tabela(fig, 3, table_data_marg, "Margem (%) - Top 20 vs Outros", 
                    ['Top 20', 'Outros'], meses_validas, is_percent=True)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 7 ==========
    fig = nova_pagina("3. Top 20 vs Outros", is_continuacao=True)
    
    # Gráfico tonelagem (Top 20 vs Outros)
    width = 0.35
    x = np.arange(len(meses_validas))
    
    ax1 = fig.add_subplot(3, 1, 1)
    bars1 = ax1.bar(x - width/2, table_data_ton[0], width, label='Top 20', color='#1f77b4')
    bars2 = ax1.bar(x + width/2, table_data_ton[1], width, label='Outros', color='#ff7f0e')
    ax1.set_title('Tonelagem (kg) - Top 20 vs Outros')
    ax1.set_xticks(x)
    ax1.set_xticklabels(meses_validas)
    ax1.legend()
    ax1.grid(True, linestyle='--', alpha=0.7)
    
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:,.3f}'.replace(",", "X").replace(".", ",").replace("X", "."),
                    ha='center', va='bottom', fontsize=8)
    
    # Gráfico faturamento (Top 20 vs Outros)
    ax2 = fig.add_subplot(3, 1, 2)
    bars1 = ax2.bar(x - width/2, table_data_fat[0], width, label='Top 20', color='#1f77b4')
    bars2 = ax2.bar(x + width/2, table_data_fat[1], width, label='Outros', color='#ff7f0e')
    ax2.set_title('Faturamento (R$) - Top 20 vs Outros')
    ax2.set_xticks(x)
    ax2.set_xticklabels(meses_validas)
    ax2.legend()
    ax2.grid(True, linestyle='--', alpha=0.7)
    
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height,
                    f'R$ {height:,.2f}'.replace(",", "X").replace(".", ",").replace("X", "."),
                    ha='center', va='bottom', fontsize=8)
    
    # Gráfico margem (Top 20 vs Outros)
    ax3 = fig.add_subplot(3, 1, 3)
    bars1 = ax3.bar(x - width/2, table_data_marg[0], width, label='Top 20', color='#1f77b4')
    bars2 = ax3.bar(x + width/2, table_data_marg[1], width, label='Outros', color='#ff7f0e')
    ax3.set_title('Margem (%) - Top 20 vs Outros')
    ax3.set_xticks(x)
    ax3.set_xticklabels(meses_validas)
    ax3.legend()
    ax3.grid(True, linestyle='--', alpha=0.7)
    
    for bars in [bars1, bars2]:
        for bar in bars:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{height:.2f}%',
                    ha='center', va='bottom', fontsize=8)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

print(f"Relatório gerado com sucesso em: {pdf_path}")