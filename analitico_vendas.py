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
base_path = r"C:\Users\win11\OneDrive\Documentos\Relação vendedores\Ranking de Vendas\2025"
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro"]

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
    
    try:
        # Extrair estatísticas gerais (valores inteiros)
        dados['qtde_clientes'] = int(re.search(r'Qtde Clientes\s+(\d+)', texto_tabela).group(1))
        dados['qtde_vendedores'] = int(re.search(r'Qtde Vendedores\s+(\d+)', texto_tabela).group(1))
        dados['qtde_produtos'] = int(re.search(r'Qtde Produtos\s+(\d+)', texto_tabela).group(1))
    except (AttributeError, ValueError):
        dados['qtde_clientes'] = 0
        dados['qtde_vendedores'] = 0
        dados['qtde_produtos'] = 0
    
    # Inicializar todos os campos com valores padrão
    dados['tonelagem'] = {'total': 0.0, 'top_20': 0.0, 'outros': 0.0}
    dados['faturamento'] = {'total': 0.0, 'top_20': 0.0, 'outros': 0.0}
    dados['margem'] = {'total': 0.0, 'top_20': 0.0, 'outros': 0.0}
    
    try:
        # Extrair tonelagem (3 decimais)
        tonelagem_total = re.search(r'Tonelagem.*?Total:\s*([\d.,]+)', texto_tabela, re.DOTALL)
        if tonelagem_total:
            # Encontrar todos os valores entre parênteses na seção de tonelagem
            valores_ton = re.findall(r'\(([\d.,]+)\)', texto_tabela.split('Tonelagem')[-1].split('Faturamento')[0])
            if len(valores_ton) >= 2:
                dados['tonelagem'] = {
                    'total': round(limpar_numero(tonelagem_total.group(1)), 3),
                    'top_20': round(limpar_numero(valores_ton[1]), 3),
                    'outros': round(limpar_numero(valores_ton[0]), 3)
                }
    except Exception as e:
        print(f"Erro ao extrair tonelagem: {e}")
    
    try:
        # Extrair faturamento (2 decimais)
        faturamento_total = re.search(r'Faturamento.*?Total:\s*([R\$ \d.,]+)', texto_tabela, re.DOTALL)
        if faturamento_total:
            # Encontrar todos os valores entre parênteses na seção de faturamento
            valores_fat = re.findall(r'\(([R\$ \d.,]+)\)', texto_tabela.split('Faturamento')[-1].split('Margem')[0])
            if len(valores_fat) >= 2:
                dados['faturamento'] = {
                    'total': round(limpar_numero(faturamento_total.group(1)), 2),
                    'top_20': round(limpar_numero(valores_fat[1]), 2),
                    'outros': round(limpar_numero(valores_fat[0]), 2)
                }
    except Exception as e:
        print(f"Erro ao extrair faturamento: {e}")
    
    try:
        # Extrair margem (valores já estão com ponto decimal)
        margem_total = re.search(r'Margem.*?Total:\s*([\d.,]+%)', texto_tabela, re.DOTALL)
        if margem_total:
            # Encontrar todos os valores de margem
            margens = re.findall(r'(\d+\.\d+%)', texto_tabela.split('Margem')[-1])
            
            if len(margens) >= 3:  # Total + Top20 + Outros
                dados['margem'] = {
                    'total': round(limpar_numero(margem_total.group(1), True), 2),
                    'top_20': round(limpar_numero(margens[1], True), 2),
                    'outros': round(limpar_numero(margens[2], True), 2)
                }
            elif len(margens) == 1:  # Apenas Total
                dados['margem']['total'] = round(limpar_numero(margem_total.group(1), True), 2)
    except Exception as e:
        print(f"Erro ao extrair margem: {e}")
    
    return dados

for mes in meses:
    pdf_path = os.path.join(base_path, mes, f"Ranking de Vendas - Geral - {mes} 2025.pdf")
    
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
pdf_path = os.path.join(downloads_path, f"Relatório Analítico de Vendas - {primeiro_mes} a {ultimo_mes} de {ano}.pdf")

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

        # Ajuste dinâmico do espaçamento baseado no número de itens
        if num_itens == 1:
            fig.subplots_adjust(hspace=0.1, top=0.92, bottom=0.08)
        elif num_itens == 2:
            fig.subplots_adjust(hspace=0.3, top=0.94, bottom=0.06)
        else:
            fig.subplots_adjust(hspace=0.4, top=0.96, bottom=0.04)

        # Adiciona subplots vazios para manter o mesmo tamanho de página
        for i in range(num_itens, 3):
            ax = fig.add_subplot(3, 1, i+1)
            ax.axis('off')

        return fig

    def add_tabela(fig, posicao, dados, titulo, row_labels, col_labels, is_money=False, is_percent=False, is_large_table=False, show_total=True, is_integer=False):
        ax = fig.add_subplot(3, 1, posicao)
        ax.axis('off')

        # Configurações de tamanho ajustadas dinamicamente
        num_meses = len(row_labels)

        # Ajustar tamanhos baseado no número de meses
        if num_meses <= 6:
            col_width = 0.18
            row_height = 0.12
            header_height = 0.15
            font_size = 9
        elif num_meses <= 9:
            col_width = 0.16
            row_height = 0.10
            header_height = 0.13
            font_size = 8
        else:
            col_width = 0.14
            row_height = 0.08
            header_height = 0.11
            font_size = 7

        # Ajustes adicionais para tabelas grandes
        if is_large_table:
            col_width *= 1.2
            row_height *= 1.1
            header_height *= 1.1

        # Formatação dos dados
        formatted_data = []
        for row in dados:
            formatted_row = []
            for val in row:
                if is_money:
                    formatted_row.append(f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                elif is_percent:
                    formatted_row.append(f"{val:.2f}%")
                elif is_integer:
                    # Para números inteiros, formata sem casas decimais
                    formatted_row.append(f"{int(val):,}".replace(",", "."))
                elif isinstance(val, float):
                    if 'Tonelagem' in titulo:
                        formatted_row.append(f"{val:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    else:
                        formatted_row.append(f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                else:
                    formatted_row.append(f"{val:,}".replace(",", "."))
            formatted_data.append(formatted_row)

        # Calcular somas das colunas se mostrar total
        if show_total and dados:
            num_cols = len(dados[0])
            somas = []
            for col_idx in range(num_cols):
                if is_percent:
                    # Para percentuais, calcular média
                    soma_col = sum(row[col_idx] for row in dados) / len(dados)
                    somas.append(f"{soma_col:.2f}%")
                elif is_integer:
                    # Para inteiros, soma normal
                    soma_col = sum(row[col_idx] for row in dados)
                    somas.append(f"{int(soma_col):,}".replace(",", "."))
                elif is_money:
                    soma_col = sum(row[col_idx] for row in dados)
                    somas.append(f"R$ {soma_col:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                else:
                    soma_col = sum(row[col_idx] for row in dados)
                    if 'Tonelagem' in titulo:
                        somas.append(f"{soma_col:,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    else:
                        somas.append(f"{soma_col:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            somas = []

        col_labels = ['Meses'] + col_labels
        formatted_data_with_months = []
        for month, row in zip(row_labels, formatted_data):
            formatted_data_with_months.append([month] + row)

        # Adicionar linha de soma se necessário
        total_row_index = None
        if show_total and somas:
            formatted_data_with_months.append(['TOTAL'] + somas)
            total_row_index = len(formatted_data_with_months)  # Índice da linha TOTAL

        num_cols = len(col_labels)
        num_rows = len(formatted_data_with_months)
        table_width = num_cols * col_width
        table_height = (num_rows * row_height) + header_height

        # Ajustar posição vertical baseado no número de linhas
        if num_rows <= 8:
            bottom_pos = 0.1
        elif num_rows <= 10:
            bottom_pos = 0.05
        else:
            bottom_pos = 0.02

        left_pos = (0.94 - table_width) / 2

        table = ax.table(cellText=formatted_data_with_months,
                       rowLabels=[''] * num_rows,
                       colLabels=col_labels,
                       loc='center',
                       cellLoc='center',
                       bbox=[left_pos, bottom_pos, table_width, table_height])

        # Formatação de células - CORRIGIDA
        table.auto_set_font_size(False)
        table.set_fontsize(font_size)

        for key, cell in table.get_celld().items():
            row, col = key

            # Configurações básicas para todas as células
            cell.set_width(col_width)
            if row == 0:
                cell.set_height(header_height)
            else:
                cell.set_height(row_height)
            cell.set_edgecolor('lightgray')
            cell.set_linewidth(0.3)

            # Cabeçalho (linha 0) - TODAS as células
            if row == 0:
                cell.set_text_props(weight='bold', fontsize=font_size)
                cell.set_facecolor('#f0f0f0')

            # Linha TOTAL - TODAS as células (incluindo a primeira coluna)
            elif total_row_index is not None and row == total_row_index:
                cell.set_text_props(weight='bold', fontsize=font_size)
                cell.set_facecolor('#f0f0f0')  # Mesma cor do cabeçalho

            # Primeira coluna (coluna 0) - exceto cabeçalho e linha TOTAL
            elif col == 0 and row > 0 and (total_row_index is None or row != total_row_index):
                cell.set_text_props(weight='bold', fontsize=font_size)
                cell.set_facecolor('#f5f5f5')

        return fig

    # Função para adicionar gráfico de linhas
    def add_grafico_linhas(fig, posicao, dados_clientes, dados_vendedores, dados_produtos, meses_validas):
        ax = fig.add_subplot(3, 1, posicao)
        ax.clear()
        
        x = np.arange(len(meses_validas))
        ax.plot(x, dados_clientes, marker='o', label='Clientes', color='skyblue')
        ax.plot(x, dados_vendedores, marker='s', label='Vendedores', color='lightgreen')
        ax.plot(x, dados_produtos, marker='^', label='Produtos', color='salmon')
        
        ax.set_xticks(x)
        ax.set_xticklabels(meses_validas)
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.legend()
        
        # Adicionar valores nos pontos
        for i, (c, v, p) in enumerate(zip(dados_clientes, dados_vendedores, dados_produtos)):
            ax.text(i, c, f"{c}", ha='center', va='bottom', fontsize=7)
            ax.text(i, v, f"{v}", ha='center', va='bottom', fontsize=7)
            ax.text(i, p, f"{p}", ha='center', va='bottom', fontsize=7)
        
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
    
    # Tabela de quantidades (formato atualizado) - COM TOTAL mas como números inteiros
    fig = add_tabela(fig, 1,
                [[dados[mes]['qtde_clientes'], 
                 dados[mes]['qtde_vendedores'], 
                 dados[mes]['qtde_produtos']] for mes in meses_validas],
                "Quantidades",
                meses_validas,
                ['Clientes', 'Vendedores', 'Produtos'],
                is_large_table=True,
                show_total=True,  # COM TOTAL para quantidades
                is_integer=True)  # Especifica que são números inteiros
    
    # Gráfico de linhas combinado
    fig = add_grafico_linhas(fig, 2, 
                            [dados[mes]['qtde_clientes'] for mes in meses_validas],
                            [dados[mes]['qtde_vendedores'] for mes in meses_validas],
                            [dados[mes]['qtde_produtos'] for mes in meses_validas],
                            meses_validas)
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 3 ==========
    fig = nova_pagina("2. Totais")
    
    # Tabela de totais (formato atualizado) - COM TOTAL
    fig = add_tabela(fig, 1,
                [[dados[mes]['tonelagem']['total'],
                 dados[mes]['faturamento']['total'],
                 dados[mes]['margem']['total']] for mes in meses_validas],
                "Totais",
                meses_validas,
                ['Tonelagem (kg)', 'Faturamento (R$)', 'Margem (%)'],
                is_large_table=True,
                show_total=True)  # COM TOTAL para totais
    
    # Gráfico de tonelagem
    fig = add_grafico(fig, 2, 
                     [dados[mes]['tonelagem']['total'] for mes in meses_validas], 
                     "Tonelagem Total (kg)", 'skyblue')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # ========== PÁGINA 4 ==========
    fig = nova_pagina("2. Totais (continuação)", num_itens=2)

    # Gráfico de faturamento
    fig = add_grafico(fig, 1, 
                     [dados[mes]['faturamento']['total'] for mes in meses_validas], 
                     "Faturamento Total (R$)", 'lightgreen')
    
    # Gráfico de margem
    fig = add_grafico(fig, 2, 
                     [dados[mes]['margem']['total'] for mes in meses_validas], 
                     "Margem Total (%)", 'salmon', is_percent=True)

    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
   # ========== PÁGINA 5 ==========
    fig = nova_pagina("3. Top 20 vs Outros", num_itens=2)

    # Tabela consolidada - SEM TOTAL
    fig = add_tabela(fig, 1,
                [[dados[mes]['tonelagem']['top_20'],
                  dados[mes]['tonelagem']['outros'],
                  dados[mes]['faturamento']['top_20'],
                  dados[mes]['faturamento']['outros'],
                  dados[mes]['margem']['top_20'],
                  dados[mes]['margem']['outros']] for mes in meses_validas],
                "Top 20 vs Outros - Resumo",
                meses_validas,
                ['Top20 Ton.', 'Outros Ton.', 'Top20 Fat.', 'Outros Fat.', 'Top20 Marg.', 'Outros Marg.'],
                show_total=False)  # SEM TOTAL para Top 20 vs Outros

    # Gráfico de faturamento (Top 20 vs Outros) - Primeiro gráfico
    ax1 = fig.add_subplot(3, 1, 2)
    width = 0.35
    x = np.arange(len(meses_validas))

    bars1 = ax1.bar(x - width/2, [dados[mes]['faturamento']['top_20'] for mes in meses_validas],
                   width, label='Top 20', color='#1f77b4')
    bars2 = ax1.bar(x + width/2, [dados[mes]['faturamento']['outros'] for mes in meses_validas], 
                   width, label='Outros', color='#ff7f0e')

    ax1.set_title('Faturamento (R$) - Top 20 vs Outros')
    ax1.set_xticks(x)
    ax1.set_xticklabels(meses_validas)
    ax1.legend()
    ax1.grid(True, linestyle='--', alpha=0.7)
    
    # MELHORIA: Adicionar mais valores no eixo Y
    ax1.yaxis.set_major_locator(plt.MaxNLocator(10))  # Aumenta para 10 valores no eixo Y

    # REMOVIDO: Código que adicionava valores nas barras

    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

    # ========== PÁGINA 6 ==========
    fig = nova_pagina("3. Top 20 vs Outros (continuação)", num_itens=2)

    # Gráfico de tonelagem (Top 20 vs Outros) - Segundo gráfico
    ax2 = fig.add_subplot(3, 1, 1)
    bars1 = ax2.bar(x - width/2, [dados[mes]['tonelagem']['top_20'] for mes in meses_validas], 
                   width, label='Top 20', color='#1f77b4')
    bars2 = ax2.bar(x + width/2, [dados[mes]['tonelagem']['outros'] for mes in meses_validas], 
                   width, label='Outros', color='#ff7f0e')

    ax2.set_title('Tonelagem (kg) - Top 20 vs Outros')
    ax2.set_xticks(x)
    ax2.set_xticklabels(meses_validas)
    ax2.legend()
    ax2.grid(True, linestyle='--', alpha=0.7)
    
    # MELHORIA: Adicionar mais valores no eixo Y
    ax2.yaxis.set_major_locator(plt.MaxNLocator(10))  # Aumenta para 10 valores no eixo Y

    # REMOVIDO: Código que adicionava valores nas barras

    # Gráfico de margem (Top 20 vs Outros) - Terceiro gráfico
    ax3 = fig.add_subplot(3, 1, 2)
    bars1 = ax3.bar(x - width/2, [dados[mes]['margem']['top_20'] for mes in meses_validas], 
                  width, label='Top 20', color='#1f77b4')
    bars2 = ax3.bar(x + width/2, [dados[mes]['margem']['outros'] for mes in meses_validas], 
                  width, label='Outros', color='#ff7f0e')

    ax3.set_title('Margem (%) - Top 20 vs Outros')
    ax3.set_xticks(x)
    ax3.set_xticklabels(meses_validas)
    ax3.legend()
    ax3.grid(True, linestyle='--', alpha=0.7)
    
    # MELHORIA: Adicionar mais valores no eixo Y
    ax3.yaxis.set_major_locator(plt.MaxNLocator(10))  # Aumenta para 10 valores no eixo Y

    # REMOVIDO: Código que adicionava valores nas barras

    pdf.savefig(fig, bbox_inches='tight')
    plt.close()

print(f"Relatório gerado com sucesso em: {pdf_path}")