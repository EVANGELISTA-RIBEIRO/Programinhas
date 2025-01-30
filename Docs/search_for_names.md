# Documentação do Script de Busca de Nomes em Arquivos Word

## Visão Geral
Este script percorre arquivos Word (`.docx`) dentro de um diretório especificado e busca por uma lista de nomes dentro do conteúdo do documento, incluindo textos dentro e fora de tabelas. Ele retorna os arquivos em que os nomes foram encontrados.

## Dependências
Este script requer as seguintes bibliotecas Python:

```sh
pip install python-docx
```

## Estrutura do Código

### 1. `read_docx_with_tables(file_path)`
Esta função lê um arquivo Word e extrai:
- **Texto normal**: Parágrafos do documento.
- **Texto em tabelas**: Cada célula das tabelas contidas no documento.

Os textos extraídos são retornados como uma única string separada por quebras de linha.

### 2. `find_names_in_word_files(directory, names)`
Esta função:
- Percorre todos os arquivos `.docx` no diretório especificado.
- Usa `read_docx_with_tables()` para obter o texto de cada arquivo.
- Verifica se algum dos nomes fornecidos está presente no texto do arquivo.
- Retorna um dicionário onde as chaves são os nomes dos arquivos e os valores são as listas de nomes encontrados dentro deles.

### 3. Configuração e Execução
- Define o diretório que contém os arquivos Word.
- Define a lista de nomes a serem procurados.
- Chama `find_names_in_word_files()` e exibe os resultados no console.

## Exemplo de Uso

```python
directory = r'C:\Users\joao.ribeiro\Documents\GPT'
names_to_find = [
    "Henrique",
    "Rodney",
    "Sebastião"
]
found_names = find_names_in_word_files(directory, names_to_find)

if found_names:
    for filename, found in found_names.items():
        print(f"Nomes encontrados no arquivo {filename}: {', '.join(found)}")
else:
    print("Nenhum nome foi encontrado nos arquivos.")
```

## Observações
- O script ignora arquivos que não possuem a extensão `.docx`.
- A busca por nomes é **case insensitive**.
- Se ocorrer um erro ao ler um arquivo, ele será registrado no console.

## Conclusão
Este script automatiza a busca de nomes dentro de documentos Word, facilitando a verificação de presença de indivíduos em um conjunto de arquivos.

## Autor

Criado por João Evangelista.