import pandas as pd
import re
from pathlib import Path
from datetime import datetime

def aplicar_mascara_processo(numero):
    """
    Aplica a máscara de número de processo judicial.
    Formato: 0000000-00.0000.0.00.0000
    Exemplo: 0082162-14.2016.8.09.0051
    """
    
    # Remove qualquer caractere que não seja número
    numero_limpo = re.sub(r'\D', '', str(numero))
    
    # Preenche com zeros à esquerda até completar 20 dígitos
    numero_padded = numero_limpo.zfill(20)
    
    # Aplica a máscara usando regex
    # Padrão: ^(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})$
    # Formato: $1-$2.$3.$4.$5.$6
    match = re.match(r'^(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})$', numero_padded)
    
    if match:
        return f"{match.group(1)}-{match.group(2)}.{match.group(3)}.{match.group(4)}.{match.group(5)}.{match.group(6)}"
    else:
        # Se não conseguir aplicar a máscara, retorna o número original
        return numero

def aplicar_mascara_planilhas():
    """
    Aplica máscara de número de processo em planilhas.
    Lê arquivos da pasta 'input' e salva na pasta 'output'.
    """
    
    # Define os diretórios
    pasta_input = Path("input")
    pasta_output = Path("output")
    
    # Cria a pasta de output se não existir
    pasta_output.mkdir(exist_ok=True)
    
    print("=" * 70)
    print("🎭 APLICADOR DE MÁSCARA - NÚMERO DE PROCESSO JUDICIAL")
    print("=" * 70)
    print("Formato: 0000000-00.0000.0.00.0000")
    print("Exemplo: 0082162-14.2016.8.09.0051")
    print("=" * 70)
    
    # Lista todos os arquivos Excel na pasta input
    arquivos_excel = list(pasta_input.glob("*.xlsx")) + list(pasta_input.glob("*.xls"))
    
    if not arquivos_excel:
        print("\n❌ Nenhuma planilha encontrada na pasta 'input'")
        return
    
    print(f"\n📂 Encontradas {len(arquivos_excel)} planilha(s) para processar\n")
    
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
                col_lower = col.lower()
                if 'numero_processo' in col_lower or 'nrprocesso' in col_lower or 'processo' in col_lower or 'nr_processo' in col_lower:
                    coluna_processo = col
                    break
            
            if coluna_processo:
                print(f"   🔍 Coluna identificada: '{coluna_processo}'")
                
                # Mostra exemplo antes da transformação
                if len(df) > 0 and pd.notna(df[coluna_processo].iloc[0]):
                    exemplo_antes = str(df[coluna_processo].iloc[0])
                    print(f"   📝 Exemplo ANTES: {exemplo_antes}")
                
                # Aplica a máscara
                print(f"   🎭 Aplicando máscara...")
                df[coluna_processo] = df[coluna_processo].apply(aplicar_mascara_processo)
                
                # Mostra exemplo depois da transformação
                if len(df) > 0:
                    exemplo_depois = df[coluna_processo].iloc[0]
                    print(f"   ✅ Exemplo DEPOIS: {exemplo_depois}")
                
                # Conta quantas máscaras foram aplicadas com sucesso
                mascaras_aplicadas = df[coluna_processo].str.match(r'^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}$').sum()
                print(f"   ✓ Máscaras aplicadas: {mascaras_aplicadas}/{len(df)}")
                
            else:
                print(f"   ⚠️  Nenhuma coluna de processo identificada")
                print(f"   Colunas disponíveis: {', '.join(df.columns.tolist())}")
                print(f"   💡 Renomeie a coluna para 'numero_processo' ou similar")
            
            # Gera nome do arquivo de saída
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_saida = arquivo.stem + f"_com_mascara_{timestamp}.xlsx"
            arquivo_saida = pasta_output / nome_saida
            
            # Salva o arquivo processado
            print(f"   💾 Salvando: {nome_saida}")
            df.to_excel(arquivo_saida, index=False)
            
            print(f"   ✅ Concluído!\n")
            
        except Exception as e:
            print(f"   ❌ Erro ao processar {arquivo.name}: {e}\n")
    
    print("=" * 70)
    print("✅ Processamento finalizado!")
    print(f"📊 Arquivo(s) salvo(s) em: {pasta_output.absolute()}")
    print("=" * 70)

if __name__ == "__main__":
    try:
        aplicar_mascara_planilhas()
    except Exception as e:
        print(f"\n❌ Erro durante a execução: {e}")
        import traceback
        traceback.print_exc()
