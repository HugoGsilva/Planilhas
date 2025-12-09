import pandas as pd
import re
from pathlib import Path
from datetime import datetime

def aplicar_mascara_processo(numero):
    """
    Aplica a máscara de número de processo judicial.
    Formato: 0000000-00.0000.0.00.0000
    """
    numero_limpo = re.sub(r'\D', '', str(numero))
    numero_padded = numero_limpo.zfill(20)
    match = re.match(r'^(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})$', numero_padded)
    
    if match:
        return f"{match.group(1)}-{match.group(2)}.{match.group(3)}.{match.group(4)}.{match.group(5)}.{match.group(6)}"
    else:
        return numero

def processar_planilhas_automatizado():
    """
    Pipeline completo de processamento de planilhas:
    0. Compara com base existente (opcional)
    1. Junta todas as planilhas da pasta 1_planilhas_brutas
    2. Remove duplicatas internas
    3. Remove duplicatas com base existente
    4. Sanitiza dados (remove quebras de linha, espaços extras)
    5. Aplica máscara no número do processo
    6. Protege colunas como texto (CPF, CNPJ, Processo)
    7. Salva resultado final na pasta 3_resultado_final
    """
    
    # Define os diretórios
    pasta_base_existente = Path("0_base_existente")
    pasta_entrada = Path("1_planilhas_brutas")
    pasta_processamento = Path("2_processamento")
    pasta_saida = Path("3_resultado_final")
    
    # Cria as pastas se não existirem
    pasta_processamento.mkdir(exist_ok=True)
    pasta_saida.mkdir(exist_ok=True)
    
    print("=" * 80)
    print("🤖 PROCESSADOR AUTOMATIZADO DE PLANILHAS")
    print("=" * 80)
    print("Pipeline completo:")
    print("  0️⃣  Verificar base existente (opcional)")
    print("  1️⃣  Juntar planilhas")
    print("  2️⃣  Remover duplicatas internas")
    print("  3️⃣  Comparar com base existente")
    print("  4️⃣  Sanitizar dados")
    print("  5️⃣  Aplicar máscara no número do processo")
    print("  6️⃣  Proteger colunas sensíveis")
    print("  7️⃣  Exportar resultado final")
    print("=" * 80)
    
    # ETAPA 1: Listar e ler planilhas
    print("\n📂 ETAPA 1: LEITURA DAS PLANILHAS")
    print("-" * 80)
    
    arquivos_excel = list(pasta_entrada.glob("*.xlsx")) + list(pasta_entrada.glob("*.xls"))
    
    if not arquivos_excel:
        print("❌ Nenhuma planilha encontrada na pasta '1_planilhas_brutas'")
        print("💡 Coloque suas planilhas Excel na pasta '1_planilhas_brutas' e execute novamente")
        return
    
    print(f"✓ Encontradas {len(arquivos_excel)} planilha(s)")
    
    dataframes = []
    total_linhas_lidas = 0
    
    for arquivo in arquivos_excel:
        try:
            # Lê o arquivo e identifica colunas que devem ser texto
            df_temp = pd.read_excel(arquivo, nrows=0)
            colunas = df_temp.columns.tolist()
            
            colunas_texto = {}
            for col in colunas:
                col_lower = col.lower()
                if any(palavra in col_lower for palavra in ['cpf', 'cnpj', 'processo', 'protocolo']):
                    colunas_texto[col] = str
            
            if colunas_texto:
                df = pd.read_excel(arquivo, dtype=colunas_texto)
            else:
                df = pd.read_excel(arquivo)
            
            print(f"  ✓ {arquivo.name}: {len(df)} linhas")
            dataframes.append(df)
            total_linhas_lidas += len(df)
            
        except Exception as e:
            print(f"  ❌ Erro ao ler {arquivo.name}: {e}")
    
    if not dataframes:
        print("\n❌ Nenhuma planilha foi carregada com sucesso")
        return
    
    print(f"\n✅ Total de linhas lidas: {total_linhas_lidas:,}")
    
    # ETAPA 2: Juntar planilhas
    print("\n🔄 ETAPA 2: JUNTANDO PLANILHAS")
    print("-" * 80)
    
    df_consolidado = pd.concat(dataframes, ignore_index=True)
    print(f"✓ Planilhas consolidadas: {len(df_consolidado):,} linhas")
    
    # ETAPA 3: Remover duplicatas internas
    print("\n🗑️  ETAPA 3: REMOVENDO DUPLICATAS INTERNAS")
    print("-" * 80)
    
    linhas_antes_dedup = len(df_consolidado)
    duplicadas = df_consolidado[df_consolidado.duplicated(keep=False)]
    
    if len(duplicadas) > 0:
        print(f"⚠️  Encontradas {len(duplicadas)} linha(s) duplicada(s) internas")
        grupos_duplicados = duplicadas.groupby(list(duplicadas.columns), dropna=False)
        print(f"   Grupos de duplicatas: {len(grupos_duplicados)}")
    
    df_consolidado = df_consolidado.drop_duplicates()
    linhas_removidas_internas = linhas_antes_dedup - len(df_consolidado)
    
    if linhas_removidas_internas > 0:
        print(f"✓ Removidas {linhas_removidas_internas:,} linha(s) duplicada(s) internas")
        print(f"✓ Taxa de duplicação interna: {(linhas_removidas_internas/linhas_antes_dedup*100):.2f}%")
    else:
        print("✓ Nenhuma duplicata interna encontrada")
    
    print(f"✓ Linhas restantes: {len(df_consolidado):,}")
    
    # ETAPA 3.5: Comparar com base existente
    print("\n🔍 ETAPA 3.5: COMPARANDO COM BASE EXISTENTE")
    print("-" * 80)
    
    arquivos_base = list(pasta_base_existente.glob("*.xlsx")) + list(pasta_base_existente.glob("*.xls"))
    linhas_removidas_base = 0
    
    if not arquivos_base:
        print("ℹ️  Nenhuma base existente encontrada em '0_base_existente'")
        print("   Pulando comparação com base existente")
        print("   💡 Para comparar com dados existentes, coloque planilhas na pasta '0_base_existente'")
    else:
        print(f"✓ Encontradas {len(arquivos_base)} planilha(s) na base existente")
        
        # Lê a base existente
        df_base_list = []
        total_linhas_base = 0
        for arquivo in arquivos_base:
            try:
                df_base_temp = pd.read_excel(arquivo)
                print(f"  ✓ {arquivo.name}: {len(df_base_temp)} linhas")
                df_base_list.append(df_base_temp)
                total_linhas_base += len(df_base_temp)
            except Exception as e:
                print(f"  ⚠️  Erro ao ler {arquivo.name}: {e}")
        
        if df_base_list:
            df_base_existente = pd.concat(df_base_list, ignore_index=True)
            print(f"\n✓ Total de linhas na base existente: {total_linhas_base:,}")
            
            # Remove traços e pontos da coluna de processo na base existente
            print("🧹 Normalizando números de processo na base existente...")
            for col in df_base_existente.columns:
                col_lower = col.lower()
                if 'numero_processo' in col_lower or 'nrprocesso' in col_lower or 'nr_processo' in col_lower or 'processo' in col_lower:
                    print(f"  ✓ Encontrada coluna de processo: '{col}'")
                    # Remove traços, pontos e outros caracteres não numéricos
                    df_base_existente[col] = df_base_existente[col].astype(str).str.replace(r'[-.\s]', '', regex=True)
                    print(f"    Traços e pontos removidos para comparação")
            
            # Verifica compatibilidade de colunas
            colunas_novos = set(df_consolidado.columns)
            colunas_base = set(df_base_existente.columns)
            
            if colunas_novos != colunas_base:
                colunas_comuns = list(colunas_novos & colunas_base)
                if not colunas_comuns:
                    print("⚠️  AVISO: Nenhuma coluna em comum - comparação não será realizada")
                else:
                    print(f"ℹ️  Usando {len(colunas_comuns)} coluna(s) em comum para comparação")
            else:
                colunas_comuns = list(df_consolidado.columns)
                print(f"✓ Colunas compatíveis ({len(colunas_comuns)} colunas)")
            
            if colunas_comuns:
                # Compara e remove duplicatas
                linhas_antes_comparacao = len(df_consolidado)
                
                # Normaliza também os dados novos para comparação (remove traços e pontos)
                df_novos_comp = df_consolidado[colunas_comuns].copy()
                df_base_comp = df_base_existente[colunas_comuns].copy()
                
                # Remove traços e pontos das colunas de processo nos dados novos também
                for col in colunas_comuns:
                    col_lower = col.lower()
                    if 'numero_processo' in col_lower or 'nrprocesso' in col_lower or 'nr_processo' in col_lower or 'processo' in col_lower:
                        df_novos_comp[col] = df_novos_comp[col].astype(str).str.replace(r'[-.\s]', '', regex=True)
                
                # Converte para string para comparação
                for col in colunas_comuns:
                    df_novos_comp[col] = df_novos_comp[col].astype(str)
                    df_base_comp[col] = df_base_comp[col].astype(str)
                
                # Cria identificador único
                df_novos_comp['_id'] = df_novos_comp.apply(lambda x: '|'.join(x.astype(str)), axis=1)
                df_base_comp['_id'] = df_base_comp.apply(lambda x: '|'.join(x.astype(str)), axis=1)
                
                # Filtra apenas registros que NÃO existem na base
                ids_existentes = set(df_base_comp['_id'])
                mask_nao_existe = ~df_novos_comp['_id'].isin(ids_existentes)
                df_consolidado = df_consolidado[mask_nao_existe].copy()
                
                linhas_removidas_base = linhas_antes_comparacao - len(df_consolidado)
                
                if linhas_removidas_base > 0:
                    print(f"✓ Removidas {linhas_removidas_base:,} linha(s) que já existem na base")
                    print(f"✓ Taxa de duplicação com base: {(linhas_removidas_base/linhas_antes_comparacao*100):.2f}%")
                else:
                    print("✓ Nenhum registro duplicado com a base existente")
                
                print(f"✓ Linhas restantes (apenas novos): {len(df_consolidado):,}")
    
    # ETAPA 4: Sanitizar dados
    print("\n🧹 ETAPA 4: SANITIZANDO DADOS")
    print("-" * 80)
    
    colunas_sanitizadas = 0
    for coluna in df_consolidado.columns:
        if df_consolidado[coluna].dtype == 'object':
            try:
                df_consolidado[coluna] = df_consolidado[coluna].fillna('')
                df_consolidado[coluna] = df_consolidado[coluna].astype(str).str.replace(r'[\n\r]+', ' ', regex=True)
                df_consolidado[coluna] = df_consolidado[coluna].str.strip()
                df_consolidado[coluna] = df_consolidado[coluna].str.replace(r'\s+', ' ', regex=True)
                df_consolidado[coluna] = df_consolidado[coluna].replace('', pd.NA)
                colunas_sanitizadas += 1
            except:
                pass
    
    print(f"✓ {colunas_sanitizadas} coluna(s) de texto sanitizada(s)")
    print("  • Quebras de linha removidas (\\n, \\r)")
    print("  • Espaços nas pontas removidos")
    print("  • Espaços múltiplos normalizados")
    
    # ETAPA 5: Aplicar máscara no número do processo
    print("\n🎭 ETAPA 5: APLICANDO MÁSCARA NO NÚMERO DO PROCESSO")
    print("-" * 80)
    
    coluna_processo = None
    for col in df_consolidado.columns:
        col_lower = col.lower()
        if 'numero_processo' in col_lower or 'nrprocesso' in col_lower or 'nr_processo' in col_lower:
            coluna_processo = col
            break
    
    if coluna_processo:
        print(f"✓ Coluna identificada: '{coluna_processo}'")
        
        # Mostra exemplo antes
        if len(df_consolidado) > 0 and pd.notna(df_consolidado[coluna_processo].iloc[0]):
            exemplo_antes = str(df_consolidado[coluna_processo].iloc[0])
            print(f"  Exemplo ANTES: {exemplo_antes}")
        
        df_consolidado[coluna_processo] = df_consolidado[coluna_processo].apply(aplicar_mascara_processo)
        
        # Mostra exemplo depois
        if len(df_consolidado) > 0:
            exemplo_depois = df_consolidado[coluna_processo].iloc[0]
            print(f"  Exemplo DEPOIS: {exemplo_depois}")
        
        mascaras_aplicadas = df_consolidado[coluna_processo].str.match(r'^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}$', na=False).sum()
        print(f"✓ Máscaras aplicadas com sucesso: {mascaras_aplicadas:,}/{len(df_consolidado):,}")
    else:
        print("⚠️  Coluna de processo não identificada - pulando aplicação de máscara")
        print(f"   Colunas disponíveis: {', '.join(df_consolidado.columns[:5].tolist())}...")
    
    # ETAPA 6: Salvar resultado
    print("\n💾 ETAPA 6: EXPORTANDO RESULTADO FINAL")
    print("-" * 80)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_saida = pasta_saida / f"planilha_processada_{timestamp}.xlsx"
    
    df_consolidado.to_excel(arquivo_saida, index=False)
    
    print(f"✓ Arquivo salvo: {arquivo_saida.name}")
    
    # RESUMO FINAL
    print("\n" + "=" * 80)
    print("📊 RESUMO DO PROCESSAMENTO")
    print("=" * 80)
    
    print(f"\n📥 Entrada:")
    print(f"  • Arquivos processados: {len(arquivos_excel)}")
    print(f"  • Total de linhas lidas: {total_linhas_lidas:,}")
    
    print(f"\n🔄 Processamento:")
    print(f"  • Duplicatas internas removidas: {linhas_removidas_internas:,}")
    print(f"  • Duplicatas com base existente removidas: {linhas_removidas_base:,}")
    print(f"  • Total de duplicatas removidas: {linhas_removidas_internas + linhas_removidas_base:,}")
    print(f"  • Colunas sanitizadas: {colunas_sanitizadas}")
    if coluna_processo:
        print(f"  • Máscaras aplicadas: {mascaras_aplicadas:,}")
    
    print(f"\n📊 Estrutura final:")
    print(f"  • Total de linhas: {len(df_consolidado):,}")
    print(f"  • Total de colunas: {len(df_consolidado.columns)}")
    print(f"  • Colunas: {', '.join(df_consolidado.columns[:5].tolist())}")
    if len(df_consolidado.columns) > 5:
        print(f"    ... e mais {len(df_consolidado.columns) - 5} coluna(s)")
    
    # Informações sobre células vazias
    total_celulas = len(df_consolidado) * len(df_consolidado.columns)
    celulas_vazias = df_consolidado.isna().sum().sum()
    print(f"\n📈 Qualidade dos dados:")
    print(f"  • Total de células: {total_celulas:,}")
    print(f"  • Células vazias: {celulas_vazias:,}")
    print(f"  • Taxa de preenchimento: {((total_celulas - celulas_vazias)/total_celulas*100):.2f}%")
    
    print(f"\n💾 Arquivo de saída:")
    print(f"  • Caminho completo: {arquivo_saida.absolute()}")
    print(f"  • Tamanho: ~{arquivo_saida.stat().st_size / 1024 / 1024:.2f} MB")
    
    print("\n" + "=" * 80)
    print("✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print("=" * 80)

if __name__ == "__main__":
    try:
        processar_planilhas_automatizado()
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
