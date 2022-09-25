import pandas as pd, string
x = pd.read_excel("debate.xlsx", usecols='F')
dicionario = {'Palavra': [], 'Qtd': []}
for i in range(len(x['Texto'])):
    x['Texto'][i] = x['Texto'][i].translate(str.maketrans({key: None for key in string.punctuation}))
    x['Texto'][i] = x['Texto'][i].split()
    for j in x['Texto'][i]:
        if not j.upper() in dicionario['Palavra']:
            dicionario['Palavra'].append(j.upper())
            dicionario['Qtd'].append(0)
        k = dicionario['Palavra'].index(j.upper())
        dicionario['Qtd'][k] = dicionario['Qtd'][k] + 1
dados = pd.DataFrame(data = dicionario)
dados = dados.sort_values(by='Qtd', ascending=False)
dados.to_excel('relatório.xlsx', index=False)