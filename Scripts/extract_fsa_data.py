import pandas as pd

def extract_fsa_data(file_path):
    xls = pd.ExcelFile(file_path)
    extracted_data = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        # Número da FSA (Sempre na célula A2)
        fsa_number = df.iloc[1, 0] if len(df) > 1 else None

        # Procurar a linha com datas
        start_date, end_date = None, None
        for row in df.itertuples():
            for cell in row:
                if isinstance(cell, str):
                    if "Início:" in cell:
                        start_date = cell.split("Início:")[-1].strip().split(" ")[0]
                    if "Término:" in cell:
                        end_date = cell.split("Término:")[-1].strip().split(" ")[0]

        # Identificar a linha onde começam os nomes (E6 ou E7 geralmente)
        start_row = 5 if len(df) > 5 and isinstance(df.iloc[5, 4], str) else 6

        # Coletar dados das pessoas (colunas de E a N)
        for _, row in df.iloc[start_row:, 4:14].dropna(how="all").iterrows():
            person_data = [fsa_number, start_date, end_date] + row.tolist()
            while len(person_data) < 13:
                person_data.append(None)  # Garantir que todas as linhas tenham 13 colunas
            extracted_data.append(person_data)

    # Criar DataFrame final
    columns = ["Número da FSA", "Início", "Término", "Nome Completo", "Filiação", "Data de Nascimento",
               "Cidadania", "Nacionalidade", "Endereço Residencial", "Empresa", "Profissão", "Atividade",
               "Função"]
    
    return pd.DataFrame(extracted_data, columns=columns)

# Caminho do arquivo Excel de entrada
file_path = "output.xlsx"  # Substitua pelo caminho real

# Extrair os dados
fsa_df = extract_fsa_data(file_path)

# Salvar os dados extraídos em um novo arquivo Excel
output_file = "FSA_Dados_Extraidos.xlsx"
fsa_df.to_excel(output_file, index=False)

print(f"Arquivo gerado com sucesso: {output_file}")
