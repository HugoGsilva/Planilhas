import pandas as pd
from pathlib import Path
from datetime import datetime

def remover_tracos():
    """
    Remove traços e pontos dos números de processo na coluna '04 - NrProcesso (short text)'.
    Lê planilhas da pasta 'planilha' e salva o resultado na pasta 'resultado'.
    """
    
    # Define os diretórios
    pasta_entrada = Path("planilha")
    pasta_saida = Path("resultado")
    
    # Cria a pasta de resultado se não existir
    pasta_saida.mkdir(exist_ok=True)
    
    # Lista todos os arquivos Excel na pasta planilha
    arquivos_excel = list(pasta_entrada.glob("*.xlsx")) + list(pasta_entrada.glob("*.xls"))
    
    if not arquivos_excel:
        print("❌ Nenhuma planilha encontrada na pasta 'planilha'")
        return
    
    print(f"📂 Encontradas {len(arquivos_excel)} planilha(s) para processar\n")
    
    # Processa cada planilha
    for arquivo in arquivos_excel:
        try:
            print(f"📖 Processando: {arquivo.name}")
            
            # Lê o arquivo Excel
            df = pd.read_excel(arquivo)
            
            print(f"   ✓ {len(df)} linhas carregadas")
            print(f"   ✓ {len(df.columns)} colunas encontradas")
            
            # Identifica a coluna de número do processo
            coluna_processo = None
            for col in df.columns:
                if 'nrprocesso' in col.lower() or 'processo' in col.lower():
                    coluna_processo = col
                    break
            
            if coluna_processo:
                print(f"   🔍 Coluna identificada: '{coluna_processo}'")
                
                # Remove traços e pontos, mantendo apenas números
                df[coluna_processo] = df[coluna_processo].astype(str).str.replace(r'[-.]', '', regex=True)
                
                print(f"   ✓ Traços e pontos removidos")
                
                # Exemplo de transformação
                if len(df) > 0:
                    exemplo_antes = "0082162-14.2016.8.09.0051"
                    exemplo_depois = df[coluna_processo].iloc[0]
                    print(f"   📝 Exemplo: {exemplo_antes} → {exemplo_depois}")
            else:
                print(f"   ⚠️  Nenhuma coluna de processo identificada")
                print(f"   Colunas disponíveis: {', '.join(df.columns.tolist())}")
            
            # Gera nome do arquivo de saída
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_saida = arquivo.stem + f"_sem_tracos_{timestamp}.xlsx"
            arquivo_saida = pasta_saida / nome_saida
            
            # Salva o arquivo processado
            print(f"   💾 Salvando: {nome_saida}")
            df.to_excel(arquivo_saida, index=False)
            
            print(f"   ✅ Concluído!\n")
            
        except Exception as e:
            print(f"   ❌ Erro ao processar {arquivo.name}: {e}\n")
    
    print("=" * 70)
    print("✅ Processamento finalizado!")
    print(f"📊 Arquivo(s) salvo(s) em: {pasta_saida.absolute()}")
    print("=" * 70)

if __name__ == "__main__":
    try:
        remover_tracos()
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
