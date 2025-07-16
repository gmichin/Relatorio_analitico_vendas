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

# Função para extrair os dados do PDF
def extrair_dados(texto):
    dados = {}
    
    # Padrões para extração
    padrao_clientes = r"Qtde Clientes\s*([\d,.]+)"
    padrao_vendedores = r"Qtde Vendedores\s*([\d,.]+)"
    padrao_produtos = r"Qtde Produtos\s*([\d,.]+)"
    
    # Extraindo os valores
    match_clientes = re.search(padrao_clientes, texto)
    match_vendedores = re.search(padrao_vendedores, texto)
    match_produtos = re.search(padrao_produtos, texto)
    
    if match_clientes:
        dados['Qtde Clientes'] = float(match_clientes.group(1).replace('.', '').replace(',', '.'))
    if match_vendedores:
        dados['Qtde Vendedores'] = float(match_vendedores.group(1).replace('.', '').replace(',', '.'))
    if match_produtos:
        dados['Qtde Produtos'] = float(match_produtos.group(1).replace('.', '').replace(',', '.'))
    
    return dados

# Processar cada arquivo PDF
for mes in meses:
    nome_arquivo = f"Relatório Analítico de Vendas - Geral - {mes} {ano}.pdf"
    caminho_pdf = os.path.join(caminho_base, mes, nome_arquivo)
    
    if os.path.exists(caminho_pdf):
        texto = extrair_texto_pdf(caminho_pdf)
        dados = extrair_dados(texto)
        dados_consolidados[mes] = dados
    else:
        print(f"Arquivo não encontrado: {caminho_pdf}")

# Exibir dados no console em formato JSON
print("Dados extraídos dos relatórios:")
print(json.dumps(dados_consolidados, indent=4, ensure_ascii=False))

# Criar DataFrame para análise
df = pd.DataFrame.from_dict(dados_consolidados, orient='index')
df.index.name = 'Mês'
df = df.reset_index()

# Criar PDF com os resultados
with PdfPages(output_pdf) as pdf:
    # Configuração geral do estilo
    plt.rcParams['font.size'] = 10
    plt.figure(figsize=(11.69, 8.27))  # Tamanho A4
    
    # Título do relatório
    plt.text(0.5, 0.95, 'Relatório Consolidado de Vendas', 
             ha='center', va='center', fontsize=16, fontweight='bold')
    plt.text(0.5, 0.90, f'Período: Maio a Junho de {ano}', 
             ha='center', va='center', fontsize=12)
    plt.axis('off')
    pdf.savefig()
    plt.close()
    
    # Tabela com os dados consolidados
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.axis('tight')
    ax.axis('off')
    
    # Preparar dados para tabela
    tabela_dados = df[['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos']].values
    
    # Criar tabela
    tabela = ax.table(cellText=tabela_dados,
                      colLabels=['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos'],
                      loc='center',
                      cellLoc='center')
    
    # Formatar tabela
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(10)
    tabela.scale(1, 1.5)
    
    # Estilo da tabela
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#f2f2f2')
    
    plt.title('Dados Consolidados por Mês', y=1.08, fontweight='bold')
    pdf.savefig()
    plt.close()
    
    # Gráfico de linhas
    plt.figure(figsize=(11, 6))
    
    # Converter valores para inteiros (assumindo que são contagens)
    df['Qtde Clientes'] = df['Qtde Clientes'].astype(int)
    df['Qtde Vendedores'] = df['Qtde Vendedores'].astype(int)
    df['Qtde Produtos'] = df['Qtde Produtos'].astype(int)
    
    # Plotar linhas
    plt.plot(df['Mês'], df['Qtde Clientes'], marker='o', label='Qtde Clientes')
    plt.plot(df['Mês'], df['Qtde Vendedores'], marker='s', label='Qtde Vendedores')
    plt.plot(df['Mês'], df['Qtde Produtos'], marker='^', label='Qtde Produtos')
    
    # Configurações do gráfico
    plt.title('Evolução Mensal das Quantidades', fontweight='bold')
    plt.xlabel('Mês')
    plt.ylabel('Quantidade')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend()
    
    # Adicionar valores nos pontos
    for i, row in df.iterrows():
        plt.text(row['Mês'], row['Qtde Clientes'], str(row['Qtde Clientes']), 
                 ha='center', va='bottom')
        plt.text(row['Mês'], row['Qtde Vendedores'], str(row['Qtde Vendedores']), 
                 ha='center', va='bottom')
        plt.text(row['Mês'], row['Qtde Produtos'], str(row['Qtde Produtos']), 
                 ha='center', va='bottom')
    
    pdf.savefig()
    plt.close()

print(f"\nRelatório gerado com sucesso e salvo em: {output_pdf}")