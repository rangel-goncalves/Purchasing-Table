import pandas as pd

# --- Configuração dos Nomes das Colunas ---
# É crucial que estes nomes correspondam exatamente aos nomes das colunas nos seus arquivos Excel.
COLUMN_NAME = 'Name'
COLUMN_DESCRIPTION = "Description"
COLUMN_DESIGNATOR = 'Designator'
COLUMN_QUANTITY = 'Quantity'
# COLUMN_MOUNTING = '.Mounting' # Descomente se esta coluna for necessária e existir nos arquivos
COLUMN_MANUFACTURER_1 = 'Manufacturer 1'
COLUMN_MANUFACTURER_PART_NUMBER_1 = 'Manufacturer Part Number 1'

# --- Configuração dos Nomes dos Arquivos ---
# Substitua pelos nomes dos seus arquivos.
# Certifique-se de que os arquivos estão no mesmo diretório que o script, ou forneça o caminho completo.
# O script assume arquivos .xlsx. Se forem .csv, mude pd.read_excel para pd.read_csv.

# Arquivo "NOVO" ou "RAW" (lista de materiais mais recente)
new_bom_file = "BOM RAW  - Remote Access Module-canon-carambola2.xlsx"
#new_bom_file = "Compras_P1_-_LE910C1-LA.xlsx"  # Exemplo do script original, ajuste conforme necessário

# Arquivo "ANTIGO" ou "Inventário" (lista de materiais de referência)
old_bom_file = "BOM CANON CARAMBOLA 2(PROTOTYPE).xlsx"
#old_bom_file = "Bill of Materials-landis-gyr-cellular-module(MCU_LE910C1-LA).xlsx"  # Exemplo do script original, ajuste conforme necessário

output_file_name = "Relatorio_Comparacao_BOM.xlsx"


def compare_boms(new_file_path, old_file_path, output_path):
    """
    Compara duas listas de materiais (BOMs) e gera um relatório das diferenças.

    Args:
        new_file_path (str): Caminho para o arquivo Excel da BOM nova/RAW.
        old_file_path (str): Caminho para o arquivo Excel da BOM antiga/inventário.
        output_path (str): Caminho para salvar o relatório Excel gerado.
    """
    try:
        # Tenta ler os arquivos Excel. Adicione error_bad_lines=False ou use openpyxl se houver problemas com xlsx.
        # Para arquivos CSV, use pd.read_csv(file_path)
        df_new = pd.read_excel(new_file_path)
        df_old = pd.read_excel(old_file_path)
        print("Arquivos Excel carregados com sucesso.")
    except FileNotFoundError:
        print(
            f"Erro: Um ou ambos os arquivos não foram encontrados. Verifique os caminhos:\n- Novo BOM: {new_file_path}\n- Antigo BOM: {old_file_path}")
        return
    except Exception as e:
        print(f"Erro ao ler os arquivos Excel: {e}")
        return

    # --- Pré-processamento e Criação de Chave de Comparação ---
    # Usar 'Manufacturer Part Number 1' e 'Description' como chave.
    # Normalizar para string e remover espaços extras para uma comparação mais robusta.
    key_cols = [COLUMN_MANUFACTURER_PART_NUMBER_1, COLUMN_DESCRIPTION]

    for col in key_cols:
        if col not in df_new.columns:
            print(f"Erro: Coluna '{col}' não encontrada na BOM nova. Verifique os nomes das colunas.")
            return
        if col not in df_old.columns:
            print(f"Erro: Coluna '{col}' não encontrada na BOM antiga. Verifique os nomes das colunas.")
            return

        df_new[col] = df_new[col].astype(str).str.strip()
        df_old[col] = df_old[col].astype(str).str.strip()

    # Criar uma chave temporária para o merge
    df_new['temp_key'] = df_new[COLUMN_MANUFACTURER_PART_NUMBER_1] + " | " + df_new[COLUMN_DESCRIPTION]
    df_old['temp_key'] = df_old[COLUMN_MANUFACTURER_PART_NUMBER_1] + " | " + df_old[COLUMN_DESCRIPTION]

    # Verificar se há chaves duplicadas em cada DataFrame antes do merge, o que pode indicar problemas nos dados de entrada
    if df_new['temp_key'].duplicated().any():
        print("Aviso: Chaves duplicadas encontradas na BOM nova. A primeira ocorrência será usada para comparação.")
        # Opcional: decidir como lidar com duplicatas (ex: agregar quantidades ou remover)
        # df_new = df_new.drop_duplicates(subset=['temp_key'], keep='first')
    if df_old['temp_key'].duplicated().any():
        print("Aviso: Chaves duplicadas encontradas na BOM antiga. A primeira ocorrência será usada para comparação.")
        # df_old = df_old.drop_duplicates(subset=['temp_key'], keep='first')

    # --- Merge das Duas BOMs ---
    # 'outer' merge para manter todos os itens de ambas as listas.
    # Suffixes são adicionados a colunas com o mesmo nome (exceto as chaves de merge).
    merged_df = pd.merge(df_new, df_old, on='temp_key', how='outer', suffixes=('_new', '_old'))
    print(f"Merge realizado. Total de linhas no DataFrame mesclado: {len(merged_df)}")

    # --- Identificação das Mudanças ---
    results = []

    for index, row in merged_df.iterrows():
        # Identificar o Part Number e Descrição principais (podem vir da parte _new ou _old)
        # Se temp_key está presente, significa que o item existe em pelo menos uma das BOMs.

        # Obter valores das colunas com sufixo, tratando ausência (NaN)
        mpn_new = row.get(COLUMN_MANUFACTURER_PART_NUMBER_1 + '_new')
        desc_new = row.get(COLUMN_DESCRIPTION + '_new')
        qty_new_val = row.get(COLUMN_QUANTITY + '_new')
        designator_new = row.get(COLUMN_DESIGNATOR + '_new')
        name_new = row.get(COLUMN_NAME + '_new')
        manufacturer_1_new = row.get(COLUMN_MANUFACTURER_1 + '_new')
        # mounting_new = row.get(COLUMN_MOUNTING + '_new') # Se for usar

        mpn_old = row.get(COLUMN_MANUFACTURER_PART_NUMBER_1 + '_old')
        desc_old = row.get(COLUMN_DESCRIPTION + '_old')
        qty_old_val = row.get(COLUMN_QUANTITY + '_old')
        designator_old = row.get(COLUMN_DESIGNATOR + '_old')
        name_old = row.get(COLUMN_NAME + '_old')
        manufacturer_1_old = row.get(COLUMN_MANUFACTURER_1 + '_old')
        # mounting_old = row.get(COLUMN_MOUNTING + '_old') # Se for usar

        # Converter quantidades para numérico, tratando erros e NaN
        # Se a quantidade não for um número válido ou estiver ausente, será tratada como 0 para cálculo da diferença.
        qty_new_numeric = pd.to_numeric(qty_new_val, errors='coerce')
        qty_old_numeric = pd.to_numeric(qty_old_val, errors='coerce')

        # Verificar se o item é novo, removido ou comum
        is_new_present = pd.notna(mpn_new) and pd.notna(desc_new)  # Item existe na BOM nova
        is_old_present = pd.notna(mpn_old) and pd.notna(desc_old)  # Item existe na BOM antiga

        status = ""
        final_mpn = ""
        final_desc = ""
        qty_diff = 0

        # Lógica para determinar o status
        if is_new_present and not is_old_present:
            status = "Adicionado"
            final_mpn = mpn_new
            final_desc = desc_new
            qty_new_numeric = 0 if pd.isna(qty_new_numeric) else qty_new_numeric
            qty_old_numeric = 0  # Não existia antes
            qty_diff = qty_new_numeric
        elif not is_new_present and is_old_present:
            status = "Removido"
            final_mpn = mpn_old
            final_desc = desc_old
            qty_new_numeric = 0  # Não existe mais
            qty_old_numeric = 0 if pd.isna(qty_old_numeric) else qty_old_numeric
            qty_diff = -qty_old_numeric
        elif is_new_present and is_old_present:  # Item comum a ambas as BOMs
            final_mpn = mpn_new  # Ou mpn_old, são os mesmos pela chave
            final_desc = desc_new  # Ou desc_old

            qty_new_numeric = 0 if pd.isna(qty_new_numeric) else qty_new_numeric
            qty_old_numeric = 0 if pd.isna(qty_old_numeric) else qty_old_numeric

            if qty_new_numeric != qty_old_numeric:
                status = "Quantidade Alterada"
                qty_diff = qty_new_numeric - qty_old_numeric
            else:
                status = "Sem Alteração"  # Ou pode optar por não incluir estes no relatório
                qty_diff = 0
        else:
            # Caso raro ou erro nos dados, onde um item não tem MPN/Descrição nem no novo nem no velho
            print(f"Aviso: Item na linha {index} do DataFrame mesclado não pôde ser classificado.")
            continue

        # Adicionar ao resultado apenas se houver uma mudança relevante (Adicionado, Removido, Quantidade Alterada)
        # Ou, se quiser incluir "Sem Alteração", remova esta condição.
        if status in ["Adicionado", "Removido", "Quantidade Alterada"
                      ]:  # Adicionado "Sem Alteração" para exemplo se quiser adicionar esses  colocar no if ", Sem Alteração"
            result_item = {
                'Status': status,
                COLUMN_MANUFACTURER_PART_NUMBER_1: final_mpn,
                COLUMN_DESCRIPTION: final_desc,
                'Quantidade_Nova': qty_new_numeric if is_new_present else 0,  # Mostrar 0 se não estiver na nova BOM
                'Quantidade_Antiga': qty_old_numeric if is_old_present else 0,  # Mostrar 0 se não estiver na antiga BOM
                'Diferenca_Quantidade': qty_diff,
                f'{COLUMN_DESIGNATOR}_Novo': designator_new if is_new_present else pd.NA,
                f'{COLUMN_DESIGNATOR}_Antigo': designator_old if is_old_present else pd.NA,
                f'{COLUMN_NAME}_Novo': name_new if is_new_present else pd.NA,
                f'{COLUMN_NAME}_Antigo': name_old if is_old_present else pd.NA,
                f'{COLUMN_MANUFACTURER_1}_Novo': manufacturer_1_new if is_new_present else pd.NA,
                f'{COLUMN_MANUFACTURER_1}_Antigo': manufacturer_1_old if is_old_present else pd.NA,
                # f'{COLUMN_MOUNTING}_Novo': mounting_new if is_new_present else pd.NA, # Se for usar
                # f'{COLUMN_MOUNTING}_Antigo': mounting_old if is_old_present else pd.NA, # Se for usar
            }
            results.append(result_item)

    if not results:
        print("Nenhuma diferença encontrada ou nenhum item processado.")
        return

    # --- Criação e Salvamento do DataFrame de Resultados ---
    df_results = pd.DataFrame(results)

    # Definir a ordem das colunas para o arquivo de saída
    output_columns_ordered = [
        'Status', COLUMN_MANUFACTURER_PART_NUMBER_1, COLUMN_DESCRIPTION,
        'Quantidade_Nova', 'Quantidade_Antiga', 'Diferenca_Quantidade',
        f'{COLUMN_DESIGNATOR}_Novo', f'{COLUMN_DESIGNATOR}_Antigo',
        f'{COLUMN_NAME}_Novo', f'{COLUMN_NAME}_Antigo',
        f'{COLUMN_MANUFACTURER_1}_Novo', f'{COLUMN_MANUFACTURER_1}_Antigo',
        # f'{COLUMN_MOUNTING}_Novo', f'{COLUMN_MOUNTING}_Antigo' # Se for usar
    ]
    # Filtrar para manter apenas as colunas que realmente existem no df_results
    # (caso alguma coluna opcional como .Mounting não tenha sido usada)
    final_output_columns = [col for col in output_columns_ordered if col in df_results.columns]

    df_results = df_results[final_output_columns]

    try:
        df_results.to_excel(output_path, index=False)
        print(f"Relatório de comparação salvo com sucesso em: {output_path}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")


# --- Execução do Script ---
if __name__ == "__main__":
    # Chamar a função de comparação
    compare_boms(new_bom_file, old_bom_file, output_file_name)

