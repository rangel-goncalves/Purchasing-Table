import pandas as pd

# inventario peças que ja temos
# boas relaçao de peças para produzir uma placa
inventario = "Compras_P1_-_LE910C1-LA.xlsx"
boas = "Bill of Materials-landis-gyr-cellular-module(MCU_LE910C1-LA).xlsx"

COLUMN_NAME = 'Name'
COLUMN_DESCRIPTION = "Description"
COLUMN_DESIGNATOR = 'Designator'
COLUMN_QUANTITY = 'Quantity'
COLUMN_MOUNTING = '.Mounting'
COLUMN_MANUFACTURER_1 = 'Manufacturer 1'
COLUMN_MANUFACTURER_PART_NUMBER_1 = 'Manufacturer Part Number 1'

COLUMN_INVENTARIO = 'Quantity'
COLUMN_VALUE = 'Value'

# numero de placa que queremos produzir
NUM_PLACAS = 1

# abre arquivos excel
old_file = pd.read_excel(inventario)
new_file = pd.read_excel(boas)

# Criar uma lista para armazenar os resultados
resultados = []
tabela = []


for index, new_element in new_file.iterrows():

    # procura os itens com o mesmo part number
    old_element = old_file[
        old_file[COLUMN_MANUFACTURER_1] == new_element[COLUMN_MANUFACTURER_1]
    ]

    # quando econtra adiciona a lista com as informacoes importantes e a quantidade necessaria para se fazer NUM_PLACAS
    if not old_element.empty:
        #debug
        print(old_element[COLUMN_INVENTARIO].values[0])
        # if abaixo filtra os resultados deixando passar sometne os que estão em falta
        if new_element[COLUMN_QUANTITY]*NUM_PLACAS - old_element[COLUMN_INVENTARIO].values[0] > 0:
            resultados.append(
                [
                    new_element[COLUMN_NAME],
                    new_element[COLUMN_DESCRIPTION],
                    new_element[COLUMN_DESIGNATOR],
                    new_element[COLUMN_QUANTITY],
                    #new_element[COLUMN_MOUNTING],
                    new_element[COLUMN_MANUFACTURER_1],
                    new_element[COLUMN_MANUFACTURER_PART_NUMBER_1],
                    old_element[COLUMN_INVENTARIO].values[0],
                    max(0,new_element[COLUMN_QUANTITY]*NUM_PLACAS - old_element[COLUMN_INVENTARIO].values[0])
                ]
            )
        # cria uma tabela de BOM, não é mais necessario
        tabela.append(
            [
                new_element[COLUMN_MANUFACTURER_PART_NUMBER_1],
                new_element[COLUMN_DESCRIPTION],
                old_element[COLUMN_INVENTARIO].values[0],
            ]
        )
    # so entra aqui se encontar um componente que não temos em inventario, ou seja faltam todos
    else:
        resultados.append(
            [
                new_element[COLUMN_NAME],
                new_element[COLUMN_DESCRIPTION],
                new_element[COLUMN_DESIGNATOR],
                new_element[COLUMN_QUANTITY],
                #new_element[COLUMN_MOUNTING],
                new_element[COLUMN_MANUFACTURER_1],
                new_element[COLUMN_MANUFACTURER_PART_NUMBER_1],
                0,
                max(0, new_element[COLUMN_QUANTITY] * NUM_PLACAS)
            ]
        )
        tabela.append(
            [
                new_element[COLUMN_MANUFACTURER_PART_NUMBER_1],
                new_element[COLUMN_DESCRIPTION],
                0,
            ]
        )

# # Criar um DataFrame a partir da lista de resultados
#df_resultados = pd.DataFrame(resultados, columns=['Name','Description','Designator', 'Quantity','Mounting','Manufacturer 1',
  #                                                'Manufacturer Part Number 1', 'Inventario', 'Comprar'])
df_resultados = pd.DataFrame(resultados, columns=['Name','Description','Designator', 'Quantity','Manufacturer 1',
                                                  'Manufacturer Part Number 1', 'Inventario', 'Comprar'])
# # Salvar o DataFrame em um arquivo Excel
res_file_name = boas.split("/")[-1].replace(".xlsx", "")
df_resultados.to_excel(res_file_name + f" Para {NUM_PLACAS} Placas vitao.xlsx", index=False)

df_table = pd.DataFrame(tabela, columns=['Manufacturer Part Number 1','Description','Quantity'])

# # Salvar o DataFrame em um arquivo Excel
res_file_name = inventario.split("/")[-1].replace(".xlsx", "")
df_table.to_excel(res_file_name + f" tipo BOM.xlsx", index=False)