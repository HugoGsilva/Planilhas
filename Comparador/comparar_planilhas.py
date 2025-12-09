import pandas as pd
from pathlib import Path
from datetime import datetime

def comparar_e_remover_duplicatas():
    """
    Compara Planilha 1 (dados novos) com Planilha 2 (dados existentes).
    Remove da Planilha 1 todos os registros que já existem na Planilha 2.
    Salva o resultado (apenas dados novos únicos) na pasta 'resultado'.
    """
    
    # Define os diretórios
    pasta_novos = Path("planilha1_novos")
    pasta_existentes = Path("planilha2_existentes")
    pasta_resultado = Path("resultado")
    
    # Cria a pasta de resultado se não existir
    pasta_resultado.mkdir(exist_ok=True)
    
    print("=" * 70)
    print("🔍 COMPARADOR DE PLANILHAS - REMOVEDOR DE DUPLICATAS")
    print("=" * 70)
    
    # Lê planilha 1 (dados novos)
    arquivos_novos = list(pasta_novos.glob("*.xlsx")) + list(pasta_novos.glob("*.xls"))
    if not arquivos_novos:
        print("\n❌ Nenhuma planilha encontrada em 'planilha1_novos'")
        return
    
    print(f"\n📂 Planilha 1 (Dados Novos): {len(arquivos_novos)} arquivo(s)")
    
    # Junta todos os arquivos da planilha 1
    df_novos_list = []
    for arquivo in arquivos_novos:
        print(f"   📖 Lendo: {arquivo.name}")
        df = pd.read_excel(arquivo)
        print(f"      ✓ {len(df)} linhas")
        df_novos_list.append(df)
    
    df_novos = pd.concat(df_novos_list, ignore_index=True)
    print(f"   ✓ Total: {len(df_novos)} linhas na Planilha 1")
    
    # Lê planilha 2 (dados existentes)
    arquivos_existentes = list(pasta_existentes.glob("*.xlsx")) + list(pasta_existentes.glob("*.xls"))
    if not arquivos_existentes:
        print("\n❌ Nenhuma planilha encontrada em 'planilha2_existentes'")
        return
    
    print(f"\n📂 Planilha 2 (Dados Existentes): {len(arquivos_existentes)} arquivo(s)")
    
    # Junta todos os arquivos da planilha 2
    df_existentes_list = []
    for arquivo in arquivos_existentes:
        print(f"   📖 Lendo: {arquivo.name}")
        df = pd.read_excel(arquivo)
        print(f"      ✓ {len(df)} linhas")
        df_existentes_list.append(df)
    
    df_existentes = pd.concat(df_existentes_list, ignore_index=True)
    print(f"   ✓ Total: {len(df_existentes)} linhas na Planilha 2")
    
    # Verifica se as colunas são compatíveis
    print(f"\n🔍 Verificando compatibilidade...")
    colunas_novos = set(df_novos.columns)
    colunas_existentes = set(df_existentes.columns)
    
    if colunas_novos != colunas_existentes:
        print(f"   ⚠️  AVISO: As colunas não são idênticas")
        print(f"   Colunas apenas em Planilha 1: {colunas_novos - colunas_existentes}")
        print(f"   Colunas apenas em Planilha 2: {colunas_existentes - colunas_novos}")
        
        # Usa apenas as colunas em comum para comparação
        colunas_comuns = list(colunas_novos & colunas_existentes)
        if not colunas_comuns:
            print(f"   ❌ Nenhuma coluna em comum encontrada!")
            return
        print(f"   ✓ Usando {len(colunas_comuns)} coluna(s) em comum para comparação")
    else:
        colunas_comuns = list(df_novos.columns)
        print(f"   ✓ Colunas compatíveis ({len(colunas_comuns)} colunas)")
    
    # Remove duplicatas da Planilha 1 que existem na Planilha 2
    print(f"\n🔄 Comparando e removendo duplicatas...")
    
    linhas_antes = len(df_novos)
    
    # Cria uma cópia apenas com as colunas comuns para comparação
    df_novos_comparacao = df_novos[colunas_comuns].copy()
    df_existentes_comparacao = df_existentes[colunas_comuns].copy()
    
    # Marca as linhas da Planilha 1 que NÃO existem na Planilha 2
    # Converte para string para comparação precisa
    for col in colunas_comuns:
        df_novos_comparacao[col] = df_novos_comparacao[col].astype(str)
        df_existentes_comparacao[col] = df_existentes_comparacao[col].astype(str)
    
    # Cria um identificador único para cada linha
    df_novos_comparacao['_id'] = df_novos_comparacao.apply(lambda x: '|'.join(x.astype(str)), axis=1)
    df_existentes_comparacao['_id'] = df_existentes_comparacao.apply(lambda x: '|'.join(x.astype(str)), axis=1)
    
    # Identifica IDs que já existem
    ids_existentes = set(df_existentes_comparacao['_id'])
    
    # Filtra apenas as linhas que NÃO existem
    mask_nao_existe = ~df_novos_comparacao['_id'].isin(ids_existentes)
    df_resultado = df_novos[mask_nao_existe].copy()
    
    linhas_depois = len(df_resultado)
    linhas_removidas = linhas_antes - linhas_depois
    
    print(f"   ✓ Comparação concluída")
    print(f"   🗑️  Removidas: {linhas_removidas} linha(s) duplicada(s)")
    print(f"   ✅ Restantes: {linhas_depois} linha(s) única(s)")
    
    if linhas_removidas > 0:
        print(f"   📊 Taxa de duplicação: {(linhas_removidas/linhas_antes*100):.2f}%")
    
    # Salva o resultado
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = pasta_resultado / f"dados_unicos_{timestamp}.xlsx"
    
    print(f"\n💾 Salvando resultado...")
    df_resultado.to_excel(arquivo_saida, index=False)
    
    # Resumo final
    print("\n" + "=" * 70)
    print("📊 RESUMO FINAL")
    print("=" * 70)
    print(f"\n📥 Entrada:")
    print(f"   • Planilha 1 (Novos): {linhas_antes:,} linhas")
    print(f"   • Planilha 2 (Existentes): {len(df_existentes):,} linhas")
    
    print(f"\n🔄 Processamento:")
    print(f"   • Linhas removidas (duplicadas): {linhas_removidas:,}")
    print(f"   • Linhas mantidas (únicas): {linhas_depois:,}")
    
    print(f"\n💾 Saída:")
    print(f"   • Arquivo: {arquivo_saida.name}")
    print(f"   • Caminho: {arquivo_saida.absolute()}")
    print(f"   • Colunas: {len(df_resultado.columns)}")
    
    print("\n" + "=" * 70)
    print("✅ Processo concluído com sucesso!")
    print("=" * 70)

if __name__ == "__main__":
    try:
        comparar_e_remover_duplicatas()
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
