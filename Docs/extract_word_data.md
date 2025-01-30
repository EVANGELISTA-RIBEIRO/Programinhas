# Documentação do Script de Extração de Dados de Arquivos Word

## Visão Geral
Este script processa arquivos Word (`.docx`) contendo informações sobre FSAs (Formulários de Solicitação de Acesso) e extrai dados relevantes, como o número da FSA, datas de acesso e informações das pessoas listadas no documento. Os dados extraídos são armazenados em um arquivo Excel, com cada FSA representada em uma planilha separada.

## Dependências
Este script requer as seguintes bibliotecas Python:

```sh
pip install python-docx pandas openpyxl
```

## Estrutura do Código

### 1. `extract_word_data(file_path)`
Esta função lê um arquivo Word específico e extrai os seguintes dados:
- **Número da FSA**: Identificado no início do documento.
- **Data de Início e Término do Acesso**: Localizadas em uma linha que contém "Início:" e "Término:".
- **Horário do Acesso**: Extraído da mesma linha das datas.
- **Informações das Pessoas**: Extraídas das tabelas do documento, incluindo nome, filiação, data de nascimento, cidadania, nacionalidade, endereço, empresa, profissão e cargo.

### 2. `sanitize_sheet_name(sheet_name)`
- Remove caracteres inválidos dos nomes das planilhas para garantir compatibilidade com o Excel.

### 3. `process_multiple_files(folder_path, output_excel_path)`
- Percorre todos os arquivos `.docx` em uma pasta.
- Para cada arquivo, extrai os dados usando `extract_word_data()`.
- Cria uma planilha para cada FSA, contendo:
  - Metadados (Número da FSA, datas e horário de acesso)
  - Dados das pessoas listadas no documento
- Salva os dados extraídos em um arquivo Excel.

## Exemplo de Uso

```python
folder_path = r"FSAs"  # Diretório contendo os arquivos .docx
output_excel_path = r"output.xlsx"  # Caminho do arquivo Excel de saída
process_multiple_files(folder_path, output_excel_path)
```

## Observações
- O script ignora arquivos temporários (`~$`) gerados pelo Word.
- Caso o número da FSA não seja encontrado, um nome de planilha genérico é usado.
- O nome da planilha é sanitizado para evitar erros no Excel.

## Conclusão
Este script automatiza a extração de informações relevantes de documentos Word, estruturando-as em planilhas Excel para fácil análise e organização.

## Autor

Criado por João Evangelista.