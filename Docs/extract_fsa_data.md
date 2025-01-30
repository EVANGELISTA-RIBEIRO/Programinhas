# Documentação do Script extract\_fsa\_data.py

## Descrição

Este script lê um arquivo Excel contendo múltiplas abas e extrai informações específicas sobre FSAs (Formulários de Solicitação de Acesso). O objetivo é estruturar esses dados e salvá-los em um novo arquivo Excel formatado corretamente.

## Funcionalidades

- Lê todas as abas de um arquivo Excel.
- Extrai o **Número da FSA**.
- Captura **datas de início e término**, que aparecem no formato `Início: data` e `Término: data`.
- Coleta informações das pessoas listadas entre as colunas `E` e `N`, começando geralmente na linha `6` ou `7`.
- Garante que todas as informações extraídas estejam bem formatadas e tenham o mesmo número de colunas.
- Salva os dados em um novo arquivo Excel.

## Dependências

- **Python 3**
- **Pandas** (`pip install pandas`)
- **OpenPyXL** (`pip install openpyxl`)

## Como Usar

1. Substitua o nome do arquivo Excel no código pelo caminho correto do arquivo que deseja processar.
2. Execute o script com Python.
3. O arquivo processado será salvo como `FSA_Dados_Extraidos.xlsx` no mesmo diretório do script.

## Estrutura do Código

### 1. Carregamento da Planilha

```python
xls = pd.ExcelFile(file_path)
```

Lê o arquivo Excel e obtém a lista de abas.

### 2. Extração do Número da FSA

```python
fsa_number = df.iloc[1, 0] if len(df) > 1 else None
```

Obtém o valor da célula `A2`.

### 3. Extração das Datas

```python
for row in df.itertuples():
    for cell in row:
        if isinstance(cell, str):
            if "Início:" in cell:
                start_date = cell.split("Início:")[-1].strip().split(" ")[0]
            if "Término:" in cell:
                end_date = cell.split("Término:")[-1].strip().split(" ")[0]
```

Procura pelas strings "Início:" e "Término:" e extrai as datas corretamente.

### 4. Identificação da Linha Inicial dos Dados

```python
start_row = 5 if len(df) > 5 and isinstance(df.iloc[5, 4], str) else 6
```

Garante que os nomes das pessoas sejam identificados corretamente.

### 5. Coleta de Dados das Pessoas

```python
for _, row in df.iloc[start_row:, 4:14].dropna(how="all").iterrows():
    person_data = [fsa_number, start_date, end_date] + row.tolist()
    while len(person_data) < 13:
        person_data.append(None)
    extracted_data.append(person_data)
```

Coleta os dados pessoais e assegura que todas as linhas tenham 13 colunas.

### 6. Salvamento do Arquivo

```python
fsa_df.to_excel(output_file, index=False)
```

Salva os dados extraídos em um novo arquivo Excel.

## Autor

Criado por João Evangelista.

