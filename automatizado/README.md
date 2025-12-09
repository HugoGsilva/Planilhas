# 🤖 Processador Automatizado de Planilhas

Sistema completo e automatizado para processar planilhas Excel com múltiplas etapas de limpeza e formatação.

## 📁 Estrutura de Pastas

```
automatizado/
├── processar_automatico.py    (Script principal - EXECUTE ESTE!)
├── README.md                   (Este arquivo)
├── 0_base_existente/          (📋 COLOQUE PLANILHAS JÁ NO DB - OPCIONAL)
├── 1_planilhas_brutas/        (📥 COLOQUE SUAS PLANILHAS AQUI)
├── 2_processamento/           (Pasta de trabalho - uso interno)
└── 3_resultado_final/         (📤 RESULTADO FINAL SAI AQUI)
```

## 🚀 Como Usar

### 0️⃣ (OPCIONAL) Base existente
Se você já tem dados no banco de dados e quer evitar duplicatas, coloque essas planilhas na pasta **`0_base_existente/`**. O script irá comparar e remover registros que já existem.

### 1️⃣ Preparar os dados
Coloque todas as suas planilhas Excel (`.xlsx` ou `.xls`) novas na pasta **`1_planilhas_brutas/`**

### 2️⃣ Executar o script
```powershell
cd automatizado
python processar_automatico.py
```

### 3️⃣ Pegar o resultado
## ⚙️ O Que o Script Faz (Pipeline Completo)

### 0. 🔍 Comparar com Base Existente (OPCIONAL)
- Se houver planilhas em `0_base_existente/`, compara os dados novos
- Remove da planilha nova todos os registros que já existem na base
- Garante que apenas dados **inéditos** sejam processados

### 1. 📂 Juntar Planilhas
- Lê todas as planilhas da pasta `1_planilhas_brutas/`
- Consolida tudo em uma única planilha
- Protege colunas com CPF, CNPJ e Processo como **TEXTO** (evita perder zeros à esquerda)

### 2. 🗑️ Remover Duplicatas Internas
- Identifica linhas duplicadas dentro das planilhas novas
- Remove duplicatas automaticamente
- Mostra estatísticas de quantas foram removidas

### 3. 🧹 Sanitizar Dadosaticamente
- Mostra estatísticas de quantas foram removidas

### 3. 🧹 Sanitizar Dados
- Remove quebras de linha (`\n`, `\r`) dentro das células
- Remove espaços extras no início e fim
- Normaliza espaços múltiplos para espaço único
- Exemplo: `"  São Paulo  "` → `"São Paulo"`

### 4. 🎭 Aplicar Máscara no Número do Processo
- Identifica automaticamente a coluna de número do processo
- Remove traços e pontos existentes
- Preenche com zeros à esquerda (20 dígitos)
- Aplica máscara padrão: `0000000-00.0000.0.00.0000`
- Exemplo: `82162142016809051` → `0082162-14.2016.8.09.0051`

### 5. 🔒 Proteger Colunas Sensíveis
- Força CPF, CNPJ e Processo como texto
- Preserva zeros à esquerda
- Evita notação científica

### 6. 💾 Exportar Resultado
- Gera arquivo final na pasta `3_resultado_final/`
- Nome com timestamp para evitar sobrescrever
- Formato Excel (.xlsx)

## 📊 Informações Exibidas

Durante o processamento, o script mostra:
- ✅ Quantos arquivos foram lidos
- ✅ Total de linhas processadas
- ✅ Quantas duplicatas internas foram removidas
- ✅ Quantas duplicatas com a base existente foram removidas
- ✅ Quantas colunas foram sanitizadas
- ✅ Quantas máscaras foram aplicadas
- ✅ Taxa de preenchimento dos dados
- ✅ Tamanho final do arquivo

## 🔧 Requisitos

### Python 3.x
```powershell
python --version
```

### Bibliotecas necessárias
```powershell
pip install pandas openpyxl
```

## 💡 Dicas

### Base Existente (Evitar Duplicatas)
Se você já processou dados anteriormente e quer adicionar apenas registros novos:
1. Coloque a planilha com dados já existentes em `0_base_existente/`
2. Coloque os dados novos em `1_planilhas_brutas/`
3. Execute o script
4. O resultado terá apenas os registros que **NÃO** existem na base

### Nome da Coluna de Processo
Para que a máscara seja aplicada automaticamente, nomeie a coluna como:
- `numero_processo`
- `nrprocesso`
- `nr_processo`
- Qualquer nome contendo "processo"

### Múltiplas Planilhas
Você pode colocar quantas planilhas quiser na pasta `1_planilhas_brutas/`. O script processa todas automaticamente.

### Arquivos Grandes
Para planilhas muito grandes (>100MB), o processamento pode demorar alguns minutos. Aguarde a conclusão.

## ⚠️ Observações

### "Nenhuma planilha encontrada"
→ Certifique-se de colocar arquivos `.xlsx` ou `.xls` na pasta `1_planilhas_brutas/`

### "Nenhuma base existente encontrada"
→ Isso é normal se você não colocou nada em `0_base_existente/`. É uma etapa opcional.

### "Coluna de processo não identificada"s

## 🆘 Problemas Comuns

### "Nenhuma planilha encontrada"
→ Certifique-se de colocar arquivos `.xlsx` ou `.xls` na pasta `1_planilhas_brutas/`

### "Coluna de processo não identificada"
→ Renomeie a coluna para conter a palavra "processo" no nome

### Erro de importação
→ Instale as dependências: `pip install pandas openpyxl`

## 📝 Exemplo de Uso Completo

```powershell
# 1. Navegar até a pasta
cd C:\Users\seu-usuario\planilhas\Planilhas\automatizado

# 2. Colocar planilhas na pasta 1_planilhas_brutas/

# 3. Executar o script
python processar_automatico.py

# 4. Pegar resultado em 3_resultado_final/
```
Você terá uma planilha limpa, consolidada e formatada:
- ✅ Sem duplicatas internas
- ✅ Sem duplicatas com base existente (se fornecida)
- ✅ Sem espaços extras ou quebras de linha
- ✅ Números de processo formatados corretamente
- ✅ CPF/CNPJ preservados como texto
- ✅ Pronta para usoras ou quebras de linha
- ✅ Números de processo formatados corretamente
- ✅ CPF/CNPJ preservados como texto
- ✅ Pronta para uso

---

**Desenvolvido para processamento automatizado de planilhas jurídicas** 📊⚖️
