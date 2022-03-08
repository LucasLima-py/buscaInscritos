import requests
from bs4 import BeautifulSoup
import pandas as pd

base = pd.read_excel(r'C:\Users\L5874789\Desktop\Arquivos - Lucas\Códigos\Brenda\Youtube.xlsx')
print(base.head())

listaBusca = base['Link tratado'].values.tolist()

listaLink = []
listaInscritos = []
i = 0

for l in listaBusca:
    try:
        url = l
        link = 'https://www.speakrj.com/audit/report/{}/youtube'.format(url)

        req = requests.get(link)
        html = req.content

        soup = BeautifulSoup(html, 'html.parser')

        util = soup.prettify()
        html_numero = soup.find('div',attrs={'class':'col d-flex flex-column justify-content-center'})

        valorFinal = html_numero.find('a',attrs={"data-toggle":"tooltip"})
        escreverExcel = valorFinal.text
        listaLink.append(l)
        listaInscritos.append(escreverExcel)
        i = i + 1
        print(i)

    except:
        print("valor não encontrado")
        continue

excelFinal = pd.DataFrame({'Link Cortado' : listaLink,
                                'Inscritos' : listaInscritos})

print(excelFinal.head())
writer = pd.ExcelWriter("C:\\Users\\L5874789\\Desktop\\Arquivos - Lucas\\Códigos\\Brenda\\resultadoFinal.xlsx", engine='xlsxwriter')
excelFinal.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()
print('excel gerado com sucesso')
