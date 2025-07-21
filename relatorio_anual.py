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
meses = ['Maio', 'Junho', 'Julho']
ano = '2025'
data_geracao = datetime.now().strftime('%Y-%m-%d')
output_pdf = os.path.join(os.path.expanduser('~'), 'Downloads', f'Relatorio_Consolidado_Vendas_{data_geracao}.pdf')

# Dicionário para armazenar os dados
dados_consolidados = {}

# Função para extrair texto de PDF
def extrair_texto_pdf(caminho_pdf):
    texto = ""
    try:
        with open(caminho_pdf, 'rb') as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            for pagina in leitor.pages:
                texto += pagina.extract_text()
    except Exception as e:
        print(f"Erro ao ler o arquivo {caminho_pdf}: {str(e)}")
    return texto

def extrair_dados(texto):
    dados = {}
    
    # Extração básica (mantida como estava)
    padroes_basicos = {
        'Qtde Clientes': r"Qtde Clientes\s*([\d,.]+)",
        'Qtde Vendedores': r"Qtde Vendedores\s*([\d,.]+)",
        'Qtde Produtos': r"Qtde Produtos\s*([\d,.]+)"
    }
    
    for campo, padrao in padroes_basicos.items():
        if match := re.search(padrao, texto):
            try:
                dados[campo] = float(match.group(1).replace('.', '').replace(',', '.'))
            except (ValueError, AttributeError):
                dados[campo] = 0
    
    # Extração Top 20 vs Resto (mantida como estava)
    padroes_top20 = {
        'top20_ton': r"\((\d+\.\d+,\d+)\)\d+\.\d+%",
        'resto_ton': r"\((\d+\.\d+,\d+)\)T otal:",
        'top20_fat': r"\(R\$([\d\.]+,\d+)\)\d+\.\d+%",
        'resto_fat': r"\(R\$([\d\.]+,\d+)\)T otal:",
        'top20_mar': r"\((\d+\.\d+)%\)\d+\.\d+%",
        'resto_mar': r"\((\d+\.\d+)%\)T otal:"
    }
    
    for campo, padrao in padroes_top20.items():
        if match := re.search(padrao, texto):
            try:
                valor = match.group(1)
                if 'mar' in campo:
                    dados[campo] = float(valor.replace('%', ''))
                else:
                    dados[campo] = float(valor.replace('.', '').replace(',', '.'))
            except (ValueError, AttributeError):
                dados[campo] = 0
    
    # CÁLCULO DOS TOTAIS BASEADO NO TOP20 + RESTO
    # Tonelagem Total = top20_ton + resto_ton
    if 'top20_ton' in dados and 'resto_ton' in dados:
        dados['total_tonelagem'] = dados['top20_ton'] + dados['resto_ton']
    
    # Faturamento Total = top20_fat + resto_fat
    if 'top20_fat' in dados and 'resto_fat' in dados:
        dados['total_faturamento'] = dados['top20_fat'] + dados['resto_fat']
    
    # Margem Total = média ponderada pelas quantidades
    if 'top20_fat' in dados and 'resto_fat' in dados:
        total_fat = dados['top20_fat'] + dados['resto_fat']
        if total_fat > 0:
            margem_top = dados.get('top20_mar', 0)
            margem_resto = dados.get('resto_mar', 0)
            dados['total_margem'] = ((dados['top20_fat'] * margem_top) + 
                                   (dados['resto_fat'] * margem_resto)) / total_fat
    
    return dados

# Processar cada arquivo PDF
for mes in meses:
    nome_arquivo = f"Relatório Analítico de Vendas - Geral - {mes} {ano}.pdf"
    caminho_pdf = os.path.join(caminho_base, mes, nome_arquivo)
    
    if os.path.exists(caminho_pdf):
        print(f"Processando arquivo: {caminho_pdf}")
        texto = extrair_texto_pdf(caminho_pdf)
        dados = extrair_dados(texto)
        dados_consolidados[mes] = dados
    else:
        print(f"Arquivo não encontrado: {caminho_pdf}")

# Verificar se existem dados para processar
if not dados_consolidados:
    print("Nenhum dado foi extraído dos arquivos PDF. Verifique os caminhos e os arquivos.")
    exit()

# Configurações de layout
PAGE_WIDTH = 11.69
PAGE_HEIGHT = 8.27
FONT_SIZE = 10
DPI = 300

def formatar_valor(valor, tipo):
    if pd.isna(valor) or valor == 0:
        return "N/A"
    try:
        if tipo == 'tonelagem':
            return f"{float(valor):,.3f} kg".replace(',', 'X').replace('.', ',').replace('X', '.')
        elif tipo == 'faturamento':
            return f"R${float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        elif tipo == 'margem':
            return f"{float(valor):.2f}%".replace('.', ',')
        return str(valor)
    except (ValueError, TypeError):
        return "N/A"

# Criar PDF
with PdfPages(output_pdf) as pdf:
    plt.rcParams.update({
        'font.size': FONT_SIZE,
        'figure.dpi': DPI,
        'figure.constrained_layout.use': True
    })
    
    # Página de título
    data_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.6, 'Relatório de Vendas', 
             ha='center', va='center', fontsize=16, fontweight='bold')
    plt.text(0.5, 0.5, f'Período: {meses[0]} a {meses[-1]} de {ano}', 
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
    
    # Tabela de quantidades
    ax_table = fig.add_axes([0.1, 0.55, 0.8, 0.35])
    ax_table.axis('off')
    
    try:
        df_quantidades = pd.DataFrame.from_dict(dados_consolidados, orient='index')
        colunas_quant = ['Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos']
        
        # Verificar quais colunas existem nos dados
        colunas_disponiveis = [col for col in colunas_quant if col in df_quantidades.columns]
        
        if colunas_disponiveis:
            df_quantidades = df_quantidades[colunas_disponiveis].astype(int)
            
            tabela = ax_table.table(
                cellText=df_quantidades.reset_index().values,
                colLabels=['Mês'] + colunas_disponiveis,
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
            
            # Gráfico de quantidades
            ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
            lines = df_quantidades.plot(ax=ax_graph, marker='o')
            ax_graph.set_title('Evolução Mensal das Quantidades', fontweight='bold')
            ax_graph.grid(True, linestyle='--', alpha=0.7)
            ax_graph.legend()
            
            for line in lines.get_lines():
                x = line.get_xdata()
                y = line.get_ydata()
                for xi, yi in zip(x, y):
                    ax_graph.text(xi, yi, f'{int(yi)}', 
                                 ha='center', va='bottom', fontsize=FONT_SIZE)
            
            ylim = ax_graph.get_ylim()
            ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
        else:
            ax_table.text(0.5, 0.5, 'Dados de quantidades não disponíveis', 
                         ha='center', va='center', fontsize=12)
            
    except Exception as e:
        ax_table.text(0.5, 0.5, f'Erro ao processar quantidades: {str(e)}', 
                     ha='center', va='center', fontsize=12, color='red')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção de Totais
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '2. Análise de Totais', 
             ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
    plt.axis('off')
    
    # Tabela de totais
    ax_table = fig.add_axes([0.1, 0.55, 0.8, 0.35])
    ax_table.axis('off')
    
    try:
        df_totais = pd.DataFrame.from_dict(dados_consolidados, orient='index')
        colunas_totais = ['total_tonelagem', 'total_faturamento', 'total_margem']
        
        # Verificar quais colunas existem nos dados
        colunas_disponiveis = [col for col in colunas_totais if col in df_totais.columns]
        
        if colunas_disponiveis:
            df_totais = df_totais[colunas_disponiveis]
            
            tabela_dados = []
            for mes, row in df_totais.iterrows():
                linha = [mes]
                if 'total_tonelagem' in colunas_disponiveis:
                    linha.append(formatar_valor(row.get('total_tonelagem'), 'tonelagem'))
                if 'total_faturamento' in colunas_disponiveis:
                    linha.append(formatar_valor(row.get('total_faturamento'), 'faturamento'))
                if 'total_margem' in colunas_disponiveis:
                    linha.append(formatar_valor(row.get('total_margem'), 'margem'))
                tabela_dados.append(linha)
            
            headers = ['Mês']
            if 'total_tonelagem' in colunas_disponiveis:
                headers.append('Tonelagem')
            if 'total_faturamento' in colunas_disponiveis:
                headers.append('Faturamento')
            if 'total_margem' in colunas_disponiveis:
                headers.append('Margem')
            
            tabela = ax_table.table(
                cellText=tabela_dados,
                colLabels=headers,
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
            
            # Gráfico de faturamento se disponível
            if 'total_faturamento' in colunas_disponiveis:
                ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
                bars = df_totais['total_faturamento'].plot(kind='bar', ax=ax_graph, color='#2ca02c')
                ax_graph.set_title('Faturamento Total', fontweight='bold')
                ax_graph.set_ylabel('R$')
                ax_graph.grid(True, linestyle='--', alpha=0.7, axis='y')
                
                for i, valor in enumerate(df_totais['total_faturamento']):
                    ax_graph.text(i, valor, formatar_valor(valor, 'faturamento'), 
                                 ha='center', va='bottom')
                
                ylim = ax_graph.get_ylim()
                ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
            else:
                ax_table.text(0.5, 0.1, 'Dados de faturamento não disponíveis', 
                             ha='center', va='center', fontsize=12)
        else:
            ax_table.text(0.5, 0.5, 'Dados de totais não disponíveis', 
                         ha='center', va='center', fontsize=12)
            
    except Exception as e:
        ax_table.text(0.5, 0.5, f'Erro ao processar totais: {str(e)}', 
                     ha='center', va='center', fontsize=12, color='red')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção de Totais - Página 2 (tonelagem e margem)
    fig = plt.figure(figsize=(PAGE_WIDTH, PAGE_HEIGHT))
    plt.text(0.5, 0.95, '2. Análise de Totais (continuação)', 
             ha='center', va='center', fontsize=14, fontweight='bold', transform=fig.transFigure)
    plt.axis('off')
    
    try:
        df_totais = pd.DataFrame.from_dict(dados_consolidados, orient='index')
        
        # Gráfico de tonelagem se disponível
        if 'total_tonelagem' in df_totais.columns:
            ax_graph1 = fig.add_axes([0.1, 0.55, 0.8, 0.35])
            df_totais['total_tonelagem'].plot(kind='bar', ax=ax_graph1, color='#1f77b4')
            ax_graph1.set_title('Tonelagem Total', fontweight='bold')
            ax_graph1.set_ylabel('kg')
            ax_graph1.grid(True, linestyle='--', alpha=0.7, axis='y')
            
            for i, valor in enumerate(df_totais['total_tonelagem']):
                ax_graph1.text(i, valor, formatar_valor(valor, 'tonelagem'), 
                              ha='center', va='bottom')
            
            ylim = ax_graph1.get_ylim()
            ax_graph1.set_ylim(ylim[0], ylim[1] * 1.1)
        else:
            ax_graph1 = fig.add_axes([0.1, 0.55, 0.8, 0.35])
            ax_graph1.text(0.5, 0.5, 'Dados de tonelagem não disponíveis', 
                           ha='center', va='center', fontsize=12)
            ax_graph1.axis('off')
        
        # Gráfico de margem se disponível
        if 'total_margem' in df_totais.columns:
            ax_graph2 = fig.add_axes([0.1, 0.1, 0.8, 0.35])
            df_totais['total_margem'].plot(kind='bar', ax=ax_graph2, color='#ff7f0e')
            ax_graph2.set_title('Margem Total', fontweight='bold')
            ax_graph2.set_ylabel('%')
            ax_graph2.grid(True, linestyle='--', alpha=0.7, axis='y')
            
            for i, valor in enumerate(df_totais['total_margem']):
                ax_graph2.text(i, valor, formatar_valor(valor, 'margem'), 
                              ha='center', va='bottom')
            
            ylim = ax_graph2.get_ylim()
            ax_graph2.set_ylim(ylim[0], ylim[1] * 1.1)
        else:
            ax_graph2 = fig.add_axes([0.1, 0.1, 0.8, 0.35])
            ax_graph2.text(0.5, 0.5, 'Dados de margem não disponíveis', 
                           ha='center', va='center', fontsize=12)
            ax_graph2.axis('off')
            
    except Exception as e:
        ax_error = fig.add_axes([0.1, 0.1, 0.8, 0.8])
        ax_error.text(0.5, 0.5, f'Erro ao processar gráficos: {str(e)}', 
                     ha='center', va='center', fontsize=12, color='red')
        ax_error.axis('off')
    
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()
    
    # Seção Top 20 vs Resto
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
        
        try:
            tabela_dados = []
            colunas_validas = True
            
            for mes in meses:
                if mes in dados_consolidados:
                    dados = dados_consolidados[mes]
                    linha = [mes]
                    linha.append(formatar_valor(dados.get(f'top20_{tipo}'), fmt))
                    linha.append(formatar_valor(dados.get(f'resto_{tipo}'), fmt))
                    tabela_dados.append(linha)
                else:
                    colunas_validas = False
                    break
            
            if colunas_validas and tabela_dados:
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
                
                # Gráfico
                ax_graph = fig.add_axes([0.1, 0.1, 0.8, 0.35])
                
                top20 = [dados_consolidados[mes].get(f'top20_{tipo}', 0) for mes in meses if mes in dados_consolidados]
                resto = [dados_consolidados[mes].get(f'resto_{tipo}', 0) for mes in meses if mes in dados_consolidados]
                
                if top20 and resto:
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
                    
                    for i in x:
                        ax_graph.text(i, top20[i], formatar_valor(top20[i], fmt), 
                                     ha='center', va='bottom', fontsize=FONT_SIZE-1)
                        ax_graph.text(i + bar_width, resto[i], formatar_valor(resto[i], fmt), 
                                     ha='center', va='bottom', fontsize=FONT_SIZE-1)
                    
                    ylim = ax_graph.get_ylim()
                    ax_graph.set_ylim(ylim[0], ylim[1] * 1.1)
                else:
                    ax_graph.text(0.5, 0.5, f'Dados insuficientes para {titulo}', 
                                 ha='center', va='center', fontsize=12)
                    ax_graph.axis('off')
            else:
                ax_table.text(0.5, 0.5, f'Dados não disponíveis para {titulo}', 
                             ha='center', va='center', fontsize=12)
                
        except Exception as e:
            ax_table.text(0.5, 0.5, f'Erro ao processar {titulo}: {str(e)}', 
                         ha='center', va='center', fontsize=12, color='red')
        
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

# Exibir resultado
print(json.dumps(dados_consolidados, indent=4, ensure_ascii=False))
print(f"\nRelatório gerado com sucesso em: {output_pdf}")