import pandas as pd
import os
from pathlib import Path
from datetime import datetime

def juntar_planilhas():
    """
    Junta todas as planilhas Excel da pasta 'Planilhas' e exporta 
    o resultado consolidado na pasta 'Resultados'.
    
    Recursos:
    - Força CPF e Número do Processo como texto (preserva zeros à esquerda)
    - Adiciona rastreamento de origem (arquivo fonte) no terminal
    - Remove duplicatas exatas
    """
    
    # Define os diretórios
    pasta_planilhas = Path("Planilhas")
    pasta_resultados = Path("Resultados")
    
    # Cria a pasta de resultados se não existir
    pasta_resultados.mkdir(exist_ok=True)
    
    # Lista todos os arquivos Excel na pasta Planilhas
    arquivos_excel = list(pasta_planilhas.glob("*.xlsx")) + list(pasta_planilhas.glob("*.xls"))
    
    if not arquivos_excel:
        print("❌ Nenhuma planilha encontrada na pasta 'Planilhas'")
        return
    
    print(f"📂 Encontradas {len(arquivos_excel)} planilha(s):")
    for arquivo in arquivos_excel:
        print(f"   - {arquivo.name}")
    
    # Lista para armazenar os DataFrames
    dataframes = []
    
    # Dicionário para rastrear origem das linhas
    rastreamento = []
    
    # Lê cada planilha
    for arquivo in arquivos_excel:
        try:
            print(f"\n📖 Lendo: {arquivo.name}")
            
            # Lê o arquivo forçando CPF e Número do Processo como texto
            # Primeiro, lê a primeira linha para identificar as colunas
            df_temp = pd.read_excel(arquivo, nrows=0)
            colunas = df_temp.columns.tolist()
            
            # Identifica colunas que devem ser texto (CPF, processo, etc)
            colunas_texto = {}
            for col in colunas:
                col_lower = col.lower()
                # Verifica se a coluna contém CPF, CNPJ ou Processo
                if any(palavra in col_lower for palavra in ['cpf', 'cnpj', 'processo', 'protocolo']):
                    colunas_texto[col] = str
            
            # Lê o arquivo com as colunas específicas como texto
            if colunas_texto:
                df = pd.read_excel(arquivo, dtype=colunas_texto)
                print(f"   🔒 Colunas travadas como texto: {list(colunas_texto.keys())}")
            else:
                df = pd.read_excel(arquivo)
            
            print(f"   ✓ {len(df)} linhas carregadas")
            
            # Adiciona rastreamento de origem (apenas para log interno)
            for idx in range(len(df)):
                rastreamento.append({
                    'linha_original': idx + 2,  # +2 porque Excel começa em 1 e tem cabeçalho
                    'arquivo_origem': arquivo.name
                })
            
            dataframes.append(df)
            
        except Exception as e:
            print(f"   ❌ Erro ao ler {arquivo.name}: {e}")
    
    if not dataframes:
        print("\n❌ Nenhuma planilha foi carregada com sucesso")
        return
    
    # Junta todas as planilhas
    print("\n🔄 Juntando planilhas...")
    df_consolidado = pd.concat(dataframes, ignore_index=True)
    
    print(f"   ✓ Total antes da remoção de duplicatas: {len(df_consolidado)} linhas")
    
    # Remove duplicatas exatas e identifica quais eram
    linhas_antes = len(df_consolidado)
    
    # Identifica duplicatas antes de remover
    duplicadas = df_consolidado[df_consolidado.duplicated(keep=False)]
    
    if len(duplicadas) > 0:
        print(f"\n🔍 ANÁLISE DE DUPLICATAS:")
        print("=" * 70)
        
        # Agrupa duplicatas idênticas
        grupos_duplicados = duplicadas.groupby(list(duplicadas.columns), dropna=False)
        
        print(f"   Total de linhas duplicadas: {len(duplicadas)}")
        print(f"   Grupos de duplicatas encontrados: {len(grupos_duplicados)}")
        print()
        
        # Mostra detalhes dos grupos duplicados
        for i, (valores, grupo) in enumerate(grupos_duplicados, 1):
            if i <= 5:  # Mostra apenas os primeiros 5 grupos para não poluir o terminal
                print(f"   Grupo {i}: {len(grupo)} ocorrências")
                
                # Pega a primeira linha do grupo para mostrar como amostra
                linha_amostra = grupo.iloc[0]
                
                # Mostra uma amostra dos dados duplicados (primeiras 3 colunas)
                colunas_amostra = df_consolidado.columns[:3].tolist()
                amostra_dict = {}
                for col in colunas_amostra:
                    valor = linha_amostra[col]
                    amostra_dict[col] = '(vazio)' if pd.isna(valor) else str(valor)[:50]  # Limita a 50 caracteres
                
                print(f"   Amostra: {amostra_dict}")
                print()
        
        if len(grupos_duplicados) > 5:
            print(f"   ... e mais {len(grupos_duplicados) - 5} grupo(s) de duplicatas")
            print()
        
        print("=" * 70)
    
    # Remove as duplicatas
    df_consolidado = df_consolidado.drop_duplicates()
    linhas_depois = len(df_consolidado)
    duplicatas_removidas = linhas_antes - linhas_depois
    
    if duplicatas_removidas > 0:
        print(f"\n   🗑️  Removidas {duplicatas_removidas} linha(s) duplicada(s)")
    else:
        print(f"\n   ✓ Nenhuma duplicata encontrada")
    
    print(f"   ✓ Total final: {len(df_consolidado)} linhas no arquivo consolidado")
    
    # Sanitização de dados (limpeza de texto)
    print("\n🧹 Sanitizando dados...")
    colunas_sanitizadas = 0
    
    for coluna in df_consolidado.columns:
        # Verifica se a coluna contém texto
        if df_consolidado[coluna].dtype == 'object':
            try:
                # Substitui NaN por string vazia ANTES de converter para string
                df_consolidado[coluna] = df_consolidado[coluna].fillna('')
                
                # Remove quebras de linha (\n, \r) e as substitui por espaço
                df_consolidado[coluna] = df_consolidado[coluna].astype(str).str.replace(r'[\n\r]+', ' ', regex=True)
                
                # Remove espaços nas pontas (strip)
                df_consolidado[coluna] = df_consolidado[coluna].str.strip()
                
                # Remove espaços duplos (ou múltiplos) e deixa apenas um
                df_consolidado[coluna] = df_consolidado[coluna].str.replace(r'\s+', ' ', regex=True)
                
                # Substitui células que ficaram vazias por NaN novamente (para o Excel entender como vazio)
                df_consolidado[coluna] = df_consolidado[coluna].replace('', pd.NA)
                
                colunas_sanitizadas += 1
            except Exception as e:
                print(f"   ⚠️  Aviso: Não foi possível sanitizar a coluna '{coluna}': {e}")
    
    print(f"   ✓ {colunas_sanitizadas} coluna(s) de texto sanitizada(s)")
    print(f"   ✓ Removidas: quebras de linha, espaços nas pontas e espaços duplos")
    
    # Mostra rastreamento detalhado
    print("\n📍 RASTREAMENTO DE ORIGEM:")
    print("=" * 70)
    
    # Agrupa por arquivo de origem
    idx_global = 0
    for i, df in enumerate(dataframes):
        arquivo_nome = arquivos_excel[i].name
        linhas_deste_arquivo = len(df)
        print(f"\n📄 {arquivo_nome}")
        print(f"   Linhas no consolidado: {idx_global + 1} até {idx_global + linhas_deste_arquivo}")
        print(f"   Total: {linhas_deste_arquivo} linhas")
        idx_global += linhas_deste_arquivo
    
    print("\n" + "=" * 70)
    
    # Gera nome do arquivo de saída com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = pasta_resultados / f"planilhas_consolidadas_{timestamp}.xlsx"
    
    # Exporta o resultado
    print(f"\n💾 Exportando para: {arquivo_saida}")
    df_consolidado.to_excel(arquivo_saida, index=False)
    
    print(f"\n✅ Processo concluído com sucesso!")
    
    # Resumo detalhado final
    print("\n" + "=" * 70)
    print("📊 RESUMO DETALHADO DO PROCESSAMENTO")
    print("=" * 70)
    
    print(f"\n📁 Arquivos processados:")
    print(f"   • Total de arquivos lidos: {len(arquivos_excel)}")
    
    print(f"\n📈 Estatísticas de linhas:")
    print(f"   • Linhas antes da deduplicação: {linhas_antes:,}")
    print(f"   • Linhas removidas (duplicatas): {duplicatas_removidas:,}")
    print(f"   • Linhas finais no arquivo: {len(df_consolidado):,}")
    print(f"   • Taxa de deduplicação: {(duplicatas_removidas/linhas_antes*100) if linhas_antes > 0 else 0:.2f}%")
    
    print(f"\n📋 Estrutura dos dados:")
    print(f"   • Total de colunas: {len(df_consolidado.columns)}")
    print(f"   • Colunas: {', '.join(df_consolidado.columns[:5].tolist())}")
    if len(df_consolidado.columns) > 5:
        print(f"     ... e mais {len(df_consolidado.columns) - 5} coluna(s)")
    
    print(f"\n🔒 Proteção de dados:")
    if colunas_texto:
        print(f"   • Colunas protegidas como texto: {len(colunas_texto)}")
        for col in colunas_texto.keys():
            print(f"     - {col}")
    else:
        print(f"   • Nenhuma coluna protegida (CPF/CNPJ/Processo não detectados)")
    
    print(f"\n🧹 Sanitização aplicada:")
    print(f"   • Colunas de texto sanitizadas: {colunas_sanitizadas}")
    print(f"   • Limpezas realizadas:")
    print(f"     - Quebras de linha removidas (\\n, \\r)")
    print(f"     - Espaços nas pontas removidos (strip)")
    print(f"     - Espaços múltiplos normalizados")
    
    # Informações sobre células vazias
    total_celulas = len(df_consolidado) * len(df_consolidado.columns)
    celulas_vazias = df_consolidado.isna().sum().sum()
    print(f"\n📊 Qualidade dos dados:")
    print(f"   • Total de células: {total_celulas:,}")
    print(f"   • Células vazias: {celulas_vazias:,}")
    print(f"   • Taxa de preenchimento: {((total_celulas - celulas_vazias)/total_celulas*100) if total_celulas > 0 else 0:.2f}%")
    
    print(f"\n💾 Arquivo de saída:")
    print(f"   • Nome: {arquivo_saida.name}")
    print(f"   • Caminho: {arquivo_saida}")
    print(f"   • Tamanho estimado: ~{os.path.getsize(arquivo_saida) / 1024 / 1024:.2f} MB")
    
    print("\n" + "=" * 70)

if __name__ == "__main__":
    try:
        juntar_planilhas()
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
