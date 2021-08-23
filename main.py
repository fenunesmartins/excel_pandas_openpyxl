import pandas as pd
tabela = pd.read_excel('Produtos.xlsx')
display(tabela)

#Exemplos utilizando o Jupyter
#Pandas mais utilizado para ler e tratar as tabelas
#Openpyxl mais utilizado para editar planilhas semelhante ao VBA
#Tabela utilizada está em anexo, Produtos.xlsx

#Atualizar o multiplicador utilizando o pandas
#tabela.loc[linha, coluna] = value 
tabela.loc[tabela["Tipo"]=="Serviço", "Multiplicador Imposto"] = 1.5

#fazer a conta do Preço Base Reais
tabela["Preço Base Reais"] = tabela["Multiplicador Imposto"] * tabela["Preço Base Original"]

#salvar no excel, removendo o índice
tabela.to_excel("ProdutosPandas.xlsx", index=False)


#Usando o openpyxl

from openpyxl import Workbook, load_workbook

planilha = load_workbook("Produtos.xlsx")
aba_ativa = planilha.active

for celula in aba_ativa["C"]:
    if celula.value == "Serviço":
        linha = celula.row
        aba_ativa[f"D{linha}"] = 1.5
    
planilha.save("ProdutosOpenPy.xlsx")
