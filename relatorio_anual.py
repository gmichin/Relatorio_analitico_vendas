import os
import camelot
import warnings
import json
import re

# Ignorar warnings (opcional)
warnings.filterwarnings("ignore")

# Dicionário para armazenar os resultados
resultados = {}

# Caminho base
base_path = r"C:\Users\win11\OneDrive\Documentos\Relatórios Analítico de Vendas\2025"
meses = ["Maio", "Junho", "Julho"]

def extrair_valores(tabela_df):
    """Extrai os valores da tabela formatada com base na estrutura específica"""
    dados = {}
    texto_tabela = tabela_df.to_string(index=False, header=False)
    
    # Extrair estatísticas gerais
    dados['qtde_clientes'] = re.search(r'Qtde Clientes\s+(\d+)', texto_tabela).group(1)
    dados['qtde_vendedores'] = re.search(r'Qtde Vendedores\s+(\d+)', texto_tabela).group(1)
    dados['qtde_produtos'] = re.search(r'Qtde Produtos\s+(\d+)', texto_tabela).group(1)
    
    # Função auxiliar para extrair valores entre parênteses
    def extrair_parenteses(texto, padrao):
        matches = re.findall(padrao + r'.*?\(([^)]+)\)', texto, re.DOTALL)
        return matches[-1] if matches else None
    
    # Extrair tonelagem
    tonelagem_total = re.search(r'Tonelagem.*?Total:\s*([\d.,]+)', texto_tabela, re.DOTALL)
    if tonelagem_total:
        # Primeiro valor entre parênteses após porcentagem é Outros
        outros_ton = extrair_parenteses(texto_tabela, r'Tonelagem.*?\d+\.\d+%')
        # Segundo valor entre parênteses após porcentagem é Top 20
        top20_ton = extrair_parenteses(texto_tabela, r'Tonelagem.*?\d+\.\d+%.*?\d+\.\d+%')
        
        dados['tonelagem'] = {
            'total': f"Total: {tonelagem_total.group(1)}",
            'top_20': top20_ton if top20_ton else "",
            'outros': outros_ton if outros_ton else ""
        }
    
    # Extrair faturamento
    faturamento_total = re.search(r'Faturamento.*?Total:\s*([R\$ \d.,]+)', texto_tabela, re.DOTALL)
    if faturamento_total:
        # Primeiro valor entre parênteses após porcentagem é Outros
        outros_fat = extrair_parenteses(texto_tabela, r'Faturamento.*?\d+\.\d+%')
        # Segundo valor entre parênteses após porcentagem é Top 20
        top20_fat = extrair_parenteses(texto_tabela, r'Faturamento.*?\d+\.\d+%.*?\d+\.\d+%')
        
        dados['faturamento'] = {
            'total': f"Total: {faturamento_total.group(1).strip()}",
            'top_20': top20_fat.strip() if top20_fat else "",
            'outros': outros_fat.strip() if outros_fat else ""
        }
    
    # Extrair margem - nova abordagem baseada no alinhamento
    margem_total = re.search(r'Margem.*?Total:\s*([\d.,]+%)', texto_tabela, re.DOTALL)
    if margem_total:
        # Encontrar todos os valores de margem com seus espaçamentos
        margens = re.findall(r'(\s+)\(([\d.,]+%)\)', texto_tabela.split('Margem')[-1])
        
        if len(margens) >= 2:
            # Ordenar pelo espaçamento - menos espaços = top_20, mais espaços = outros
            margens.sort(key=lambda x: len(x[0]))
            
            dados['margem'] = {
                'total': f"Total: {margem_total.group(1)}",
                'top_20': margens[0][1],  # Menos espaços
                'outros': margens[1][1]   # Mais espaços
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