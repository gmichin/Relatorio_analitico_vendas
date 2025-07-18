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
    
    # Seção 1: Quantidades
    plt.figure(figsize=(11.69, 8.27))
    plt.text(0.5, 0.95, '1. Análise de Quantidades', 
             ha='center', va='center', fontsize=14, fontweight='bold')
    plt.axis('off')
    pdf.savefig()
    plt.close()
    
    # Tabela de quantidades
    fig, ax = plt.subplots(figsize=(11, 4))
    ax.axis('tight')
    ax.axis('off')
    
    tabela_dados = df_quantidades[['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos']].values
    
    tabela = ax.table(cellText=tabela_dados,
                      colLabels=['Mês', 'Qtde Clientes', 'Qtde Vendedores', 'Qtde Produtos'],
                      loc='center',
                      cellLoc='center')
    
    tabela.auto_set_font_size(False)
    tabela.set_fontsize(10)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#f2f2f2')
    
    plt.title('Quantidades por Mês', y=1.08, fontweight='bold')
    pdf.savefig()
    plt.close()
    
    # Gráfico de quantidades
    plt.figure(figsize=(11, 6))
    
    df_quantidades['Qtde Clientes'] = df_quantidades['Qtde Clientes'].astype(int)
    df_quantidades['Qtde Vendedores'] = df_quantidades['Qtde Vendedores'].astype(int)
    df_quantidades['Qtde Produtos'] = df_quantidades['Qtde Produtos'].astype(int)
    
    plt.plot(df_quantidades['Mês'], df_quantidades['Qtde Clientes'], marker='o', label='Clientes')
    plt.plot(df_quantidades['Mês'], df_quantidades['Qtde Vendedores'], marker='s', label='Vendedores')
    plt.plot(df_quantidades['Mês'], df_quantidades['Qtde Produtos'], marker='^', label='Produtos')
    
    plt.title('Evolução Mensal das Quantidades', fontweight='bold')
    plt.xlabel('Mês')
    plt.ylabel('Quantidade')
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.legend()
    
    for i, row in df_quantidades.iterrows():
        plt.text(row['Mês'], row['Qtde Clientes'], str(row['Qtde Clientes']), 
                 ha='center', va='bottom')
        plt.text(row['Mês'], row['Qtde Vendedores'], str(row['Qtde Vendedores']), 
                 ha='center', va='bottom')
        plt.text(row['Mês'], row['Qtde Produtos'], str(row['Qtde Produtos']), 
                 ha='center', va='bottom')
    
    pdf.savefig()
    plt.close()
    
    # Seção 2: Totais com gráficos de barras
    plt.figure(figsize=(11.69, 8.27))
    plt.text(0.5, 0.95, '2. Análise de Totais', 
             ha='center', va='center', fontsize=14, fontweight='bold')
    plt.axis('off')
    pdf.savefig()
    plt.close()

    # Preparar dados para os gráficos
    meses_graf = df_totais['Mês']
    
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

    # Tabela de totais com formatação personalizada
    fig, ax = plt.subplots(figsize=(11, 4))
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
    tabela.set_fontsize(10)
    tabela.scale(1, 1.5)
    
    for (i, j), cell in tabela.get_celld().items():
        if i == 0:
            cell.set_text_props(fontweight='bold')
            cell.set_facecolor('#f2f2f2')
    
    plt.title('Totais por Mês', y=1.08, fontweight='bold')
    pdf.savefig()
    plt.close()
    
    # Gráfico de Tonelagem (kg)
    if 'total_tonelagem' in df_totais.columns:
        plt.figure(figsize=(11, 6))
        valores = df_totais['total_tonelagem']
        barras = plt.bar(meses_graf, valores, color='#1f77b4')
        
        plt.title('Tonelagem Total (kg)', fontweight='bold')
        plt.xlabel('Mês')
        plt.ylabel('Kg')
        plt.grid(True, linestyle='--', alpha=0.7, axis='y')
        
        # Adicionar valores nas barras
        for barra in barras:
            height = barra.get_height()
            plt.text(barra.get_x() + barra.get_width()/2., height,
                    formatar_valor(height, 'tonelagem'),
                    ha='center', va='bottom')
        
        pdf.savefig()
        plt.close()

    # Gráfico de Faturamento (R$)
    if 'total_faturamento' in df_totais.columns:
        plt.figure(figsize=(11, 6))
        valores = df_totais['total_faturamento']
        barras = plt.bar(meses_graf, valores, color='#2ca02c')
        
        plt.title('Faturamento Total', fontweight='bold')
        plt.xlabel('Mês')
        plt.ylabel('Valor (R$)')
        plt.grid(True, linestyle='--', alpha=0.7, axis='y')
        
        # Adicionar valores nas barras
        for barra in barras:
            height = barra.get_height()
            plt.text(barra.get_x() + barra.get_width()/2., height,
                    formatar_valor(height, 'faturamento'),
                    ha='center', va='bottom')
        
        pdf.savefig()
        plt.close()

    # Gráfico de Margem (%)
    if 'total_margem' in df_totais.columns:
        plt.figure(figsize=(11, 6))
        valores = df_totais['total_margem']
        barras = plt.bar(meses_graf, valores, color='#ff7f0e')
        
        plt.title('Margem Total', fontweight='bold')
        plt.xlabel('Mês')
        plt.ylabel('Percentual (%)')
        plt.grid(True, linestyle='--', alpha=0.7, axis='y')
        
        # Adicionar valores nas barras
        for barra in barras:
            height = barra.get_height()
            plt.text(barra.get_x() + barra.get_width()/2., height,
                    formatar_valor(height, 'margem'),
                    ha='center', va='bottom')
        
        pdf.savefig()
        plt.close()

print(f"\nRelatório gerado com sucesso e salvo em: {output_pdf}")