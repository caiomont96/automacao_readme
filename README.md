### Sistema de geração de cotação automatizado.
.
.

“ Este projeto é um reflexo de desafios que já tive em ambiente de trabalho real, entretanto, o formato do desafio e todos os dados, incluindo o nome da empresa, clientes e produtos, são fictícios. “

.
.

Em uma empresa do setor de agronegócio, o time de marketing e vendas está experimentando um aumento significativo na demanda, resultando em uma enxurrada de solicitações de orçamento para produtos agrícolas.
Contudo, esse aumento nas oportunidades também trouxe consigo um novo problema. 

Anteriormente, quando as solicitações eram menos numerosas, os vendedores conseguiam lidar eficientemente, processando a planilha do fornecedor, realizando as adaptações necessárias no Excel e enviando os orçamentos em PDF para cada potencial comprador.

No entanto, com a crescente demanda, os vendedores agora enfrentam o desafio de dedicar mais tempo à geração de orçamentos do que à negociação em si. Mesmo com esses esforços, muitas vezes não conseguem cumprir os prazos estipulados, resultando na perda de clientes devido a atrasos no retorno.

Diante desse cenário, a equipe de dados recebeu o desafio de desenvolver uma solução de automação para esse processo. 

O objetivo é transformar a planilha do fornecedor em um documento final em PDF, incorporando dados do cliente e margens de lucro, de maneira rápida e eficiente, permitindo que os vendedores foquem no que fazem de melhor: negociar e atender aos clientes, como no exemplo abaixo:

# Dividindo por passos:

* O fornecedor costuma enviar a planilha em um formato horizontal, onde as características estão na horizontal e os produtos na vertical, essa planilha deve ser transposta.

* Das colunas originais, apenas as de Produtos e Quantidade devem ser mantidas na versão final.

* Será necessário criar as colunas ‘Valor Unitário’ e ‘Valor Total’, ambas com a margem de lucro inserida.

* A coluna Unidade de Medida precisa ser alterada com a sua medida multiplicada por sua quantidade.

* Precisará ser criada na última linha o ‘Total’ do valor final que o cliente irá pagar.

* O PDF final precisa ter a logo da empresa no canto superior esquerdo e os dados da empresa e do cliente em cima da tabela.

* Após processar a tabela e criar o pdf, criar um executável pela biblioteca Tkinter para ser usado pelo vendedores.

# Primeiro passo

```bash
!pip install reportlab
!pip install openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from reportlab.platypus import Spacer

df = pd.read_excel('fornecedor_agro.xlsx')

```
# formato inicial da planilha
```bash

df = df.T
df = df.reset_index()
df.columns = df.iloc[0]
df = df[1:]
df

```
# formato da planilha


| Produtos                        | Un. Medida                       | Quantidade | Valor Unitário |
|---------------------------------|----------------------------------|------------|----------------|
| Herbicidas Glyphosate           | Embalagem de 1 litro(s)          | 90         | 67             |
| Herbicidas Paraquat             | Embalagem de 1 litro(s)          | 77         | 78             |
| Herbicidas Atrazine             | Embalagem de 1 litro(s)          | 11         | 34             |
| Fungicidas Mancozeb              | Embalagem de 500g                | 21         | 45.67          |
| Fungicidas Azoxystrobin          | Embalagem de 500g                | 13         | 45.99          |
| Fungicidas Tebuconazole          | Embalagem de 500g                | 5          | 95.98          |
| Inseticidas Imidacloprid         | Embalagem de 250ml               | 32         | 46.63          |
| Inseticidas Lambda-cyhalothrin   | Embalagem de 250ml               | 5          | 74.67          |
| Inseticidas Chlorpyrifos         | Embalagem de 250ml               | 7          | 91.34          |
| Fertilizantes Nitrato de amônio  | Saco de 25 kg                    | 17         | 84.84          |
| Fertilizantes Fosfato diamônico (DAP) | Saco de 25 kg                | 9          | 105.12         |
| Fertilizantes Cloreto de potássio | Saco de 25 kg                   | 45         | 115.89         |
| Reguladores de crescimento: Ácido giberélico | Embalagem de 10g    | 23         | 17.56          |
| Reguladores de crescimento: Paclobutrazol | Embalagem de 10g        | 43         | 23.45          |
| Reguladores de crescimento: Ethephon | Embalagem de 10g            | 18         | 35.47          |
| Adjuvantes Óleo mineral          | Frasco de 500ml                  | 8          | 33.23          |
| Adjuvantes Surfactantes          | Frasco de 500ml                  | 32         | 42.66          |
| Adjuvantes Espalhantes adesivos   | Frasco de 500ml                  | 19         | 55.33          |



```bash
df['Un. Medida'] = df['Un. Medida'].str.replace('Embalagem de|Saco de|Frasco de', '', regex=True)
```

| Produtos                                      | Un. Medida                         | Quantidade | Valor Unitário |
|-----------------------------------------------|------------------------------------|------------|----------------|
| Herbicidas Glyphosate                         | 1 litro(s)                         | 90         | 67             |
| Herbicidas Paraquat                           | 1 litro(s)                         | 77         | 78             |
| Herbicidas Atrazine                           | 1 litro(s)                         | 11         | 34             |
| Fungicidas Mancozeb                           | 500g                               | 21         | 45.67          |
| Fungicidas Azoxystrobin                       | 500g                               | 13         | 45.99          |
| Fungicidas Tebuconazole                       | 500g                               | 5          | 95.98          |
| Inseticidas Imidacloprid                      | 250ml                              | 32         | 46.63          |
| Inseticidas Lambda-cyhalothrin                | 250ml                              | 5          | 74.67          |
| Inseticidas Chlorpyrifos                      | 250ml                              | 7          | 91.34          |
| Fertilizantes Nitrato de amônio               | 25 kg                              | 17         | 84.84          |
| Fertilizantes Fosfato diamônico (DAP)        | 25 kg                              | 9          | 105.12         |
| Fertilizantes Cloreto de potássio             | 25 kg                              | 45         | 115.89         |
| Reguladores de crescimento: Ácido giberélico  | 10g                                | 23         | 17.56          |
| Reguladores de crescimento: Paclobutrazol     | 10g                                | 43         | 23.45          |
| Reguladores de crescimento: Ethephon          | 10g                                | 18         | 35.47          |
| Adjuvantes Óleo mineral                       | 500ml                              | 8          | 33.23          |
| Adjuvantes Surfactantes                       | 500ml                              | 32         | 42.66          |
| Adjuvantes Espalhantes adesivos               | 500ml                              | 19         | 55.33          |

```bash

df[['Medida', 'Un.']] = df['Un. Medida'].str.extract('(\d+\.?\d*)\s*([a-zA-Z\(s\)]*)')
df = df.drop('Un. Medida', axis=1)

df['Medida'] = pd.to_numeric(df['Medida'], errors='coerce')

df['Valor Total'] = df['Quantidade'] * df['Valor Unitário']
df['Quantidade Total'] = df['Quantidade'] * df['Medida']

df['Un. Medida'] = df['Quantidade Total'].astype(str) + ' ' + df['Un.']
df = df.drop(['Quantidade Total', 'Medida', 'Un.'], axis=1)


# A margem de lucro de Herbicidas é de 22%
# A margem de lucro de Fungicidas é de 14%
# A margem de lcuro de Inseticidas é de 31%
# Demais margens são 10%

def determinar_margem_lucro(produto):
    if 'Herbicidas' in produto:
        return 0.22
    elif 'Fungicidas' in produto:
        return 0.14
    elif 'Inseticidas' in produto:
        return 0.31
    else:
        return 0.10  

df['Margem Lucro'] = df['Produtos'].apply(determinar_margem_lucro)

df['Valor Unitário Ajustado'] = df['Valor Unitário'] * (1 + df['Margem Lucro'])
df['Valor Total Ajustado'] = df['Quantidade'] * df['Valor Unitário Ajustado']

df = df.drop(['Valor Unitário', 'Valor Total', 'Margem Lucro'], axis=1)

df = df.rename(columns={'Valor Unitário Ajustado': 'Valor Unitário', 'Valor Total Ajustado': 'Valor Total'})

nova_ordem_colunas = ['Produtos', 'Quantidade', 'Valor Unitário', 'Valor Total','Un. Medida']

df = df[nova_ordem_colunas]

df['Valor Unitário'] = df['Valor Unitário'].astype(float).round(2)
df['Valor Total'] = df['Valor Total'].astype(float).round(2)


df[['Un. Medida_Numero', 'Un. Medida_Medida']] = df['Un. Medida'].str.extract(r'(\d+\.?\d*)\s?(\D+)')

# Converta 'Un. Medida_Numero' para float
df['Un. Medida_Numero'] = df['Un. Medida_Numero'].astype(float)

# Aplicando as condições
condicao = (df['Un. Medida_Numero'] > 999) & ((df['Un. Medida_Medida'] == 'g') | (df['Un. Medida_Medida'] == 'ml'))

# Aplicando as conversões corrigidas
df.loc[condicao, 'Un. Medida'] = df.loc[condicao].apply(lambda row: f'{row["Un. Medida_Numero"]/1000} kg' if row["Un. Medida_Medida"] == 'g' else f'{row["Un. Medida_Numero"]/1000} litro(s)', axis=1)

# Dropando colunas auxiliares
df.drop(['Un. Medida_Numero', 'Un. Medida_Medida'], axis=1, inplace=True)


# Exibindo o DataFrame atualizado
original_ordem_colunas = df.columns.tolist()

total_linha = pd.DataFrame({'Produtos': 'Total', 'Un.': '', 'Valor Unitário': '', 'Unidades': '',
                          'Valor Total': df['Valor Total'].sum()}, index=[len(df)])

total_linha.loc[:, df.columns.difference(['Produtos', 'Valor Total'])] = ''

df = pd.concat([df, total_linha])

df = df[original_ordem_colunas]


```

# df pronto! 
| Produtos                                      | Quantidade | Valor Unitário | Valor Total | Un. Medida   |
|-----------------------------------------------|------------|----------------|-------------|--------------|
| Herbicidas Glyphosate                         | 90         | 81.74          | 7356.60     | 90 litro(s)  |
| Herbicidas Paraquat                           | 77         | 95.16          | 7327.32     | 77 litro(s)  |
| Herbicidas Atrazine                           | 11         | 41.48          | 456.28      | 11 litro(s)  |
| Fungicidas Mancozeb                           | 21         | 52.06          | 1093.34     | 10.5 kg      |
| Fungicidas Azoxystrobin                       | 13         | 52.43          | 681.57      | 6.5 kg       |
| Fungicidas Tebuconazole                       | 5          | 109.42         | 547.09      | 2.5 kg       |
| Inseticidas Imidacloprid                      | 32         | 61.09          | 1954.73     | 8.0 litro(s) |
| Inseticidas Lambda-cyhalothrin                | 5          | 97.82          | 489.09      | 1.25 litro(s)|
| Inseticidas Chlorpyrifos                      | 7          | 119.66         | 837.59      | 1.75 litro(s)|
| Fertilizantes Nitrato de amônio               | 17         | 93.32          | 1586.51     | 425 kg       |
| Fertilizantes Fosfato diamônico (DAP)        | 9          | 115.63         | 1040.69     | 225 kg       |
| Fertilizantes Cloreto de potássio             | 45         | 127.48         | 5736.56     | 1125 kg      |
| Reguladores de crescimento: Ácido giberélico  | 23         | 19.32          | 444.27      | 230 g        |
| Reguladores de crescimento: Paclobutrazol     | 43         | 25.8           | 1109.19     | 430 g        |
| Reguladores de crescimento: Ethephon          | 18         | 39.02          | 702.31      | 180 g        |
| Adjuvantes Óleo mineral                       | 8          | 36.55          | 292.42      | 4.0 litro(s) |
| Adjuvantes Surfactantes                       | 32         | 46.93          | 1501.63     | 16.0 litro(s)|
| Adjuvantes Espalhantes adesivos               | 19         | 60.86          | 1156.40     | 9.5 litro(s) |
| Total                                         |            |                | 34313.59    |              |

# Momento Criação do pdf.


```bash
nome_cliente = 'Priscila Motta'
cnpj_cliente = '11111111/0001-11'
endereco_cliente = 'Presidente Prudente'

nome_vendedor = 'Pedro Martinez'
cnpj_empresa = '11121112/0001-29'
endereco_empresa = 'Sorocaba'

caminho_logotipo = 'agro_logo.png'

nome_arquivo_pdf = 'tabela_orcamento.pdf'
doc = SimpleDocTemplate(nome_arquivo_pdf, pagesize=letter)
story = []

tabela_cabecalho = Table([
    [Image(caminho_logotipo, width=70, height=55, hAlign='LEFT'),
     f'Vendedor: Jones Carlos Manoel\nCNPJ Empresa: 11121112/0001-29\nEndereço Empresa: Sorocaba',
     f'Cliente: Itaburi Agro\nCNPJ: 11111111/0001-11\nEndereço: Presidente Prudente']
])

style_cabecalho = TableStyle([
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
])

tabela_cabecalho.setStyle(style_cabecalho)
story.append(tabela_cabecalho)
story.append(Spacer(1, 12))
data_precos = [df.columns.tolist()] + df.values.tolist()
table_precos = Table(data_precos)

style_precos = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
])

table_precos.setStyle(style_precos)
story.append(table_precos)
doc.build(story)

print(f'Tabela exportada para {nome_arquivo_pdf} com sucesso!')
```

Até o momento, o código está operacional, mas sua aplicação é restrita ao meu ambiente de compilação, carecendo de uma interface mais abrangente para ser utilizada como uma ferramenta por um vendedor.

Abaixo, irei resumir as ações em funções e incluir e criar uma aplicação pela biblioteca Tkinter

# Código separado em funções

```bash
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from reportlab.platypus import Spacer
from tkinter import StringVar

caminho_logotipo = 'agro_logo.png'
df = None
tabela_cabecalho = None  # Adicione esta linha para definir a tabela_cabecalho como uma variável global

# Parte 1: Leitura do arquivo Excel e criação do DataFrame
def processar_planilha():
    global df
    vendedor_escolhido = vendedor_var.get()
    cliente_escolhido = cliente_var.get()
    # Abre a janela para selecionar o arquivo Excel
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecionar planilha",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if caminho_arquivo:
        try:
            df_original = pd.read_excel(caminho_arquivo)
            df_transformado = transformar_dataframe(df_original)
            print(df_transformado)
            processar_e_gerar_pdf(df_transformado, vendedor_escolhido, cliente_escolhido)

        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar a planilha: {e}")



def transformar_dataframe(df_original):

    df = df_original.T

    #df = df.T
    
    df = df.reset_index()
    
    df.columns = df.iloc[0]
    
    df = df[1:]
    
    df['Un. Medida'] = df['Un. Medida'].str.replace('Embalagem de|Saco de|Frasco de', '', regex=True)
    
    df[['Medida', 'Un.']] = df['Un. Medida'].str.extract('(\d+\.?\d*)\s*([a-zA-Z\(s\)]*)')
    df = df.drop('Un. Medida', axis=1)
    
    df['Medida'] = pd.to_numeric(df['Medida'], errors='coerce')
    
    df['Valor Total'] = df['Quantidade'] * df['Valor Unitário']
    df['Quantidade Total'] = df['Quantidade'] * df['Medida']
    
    df['Un. Medida'] = df['Quantidade Total'].astype(str) + ' ' + df['Un.']
    df = df.drop(['Quantidade Total', 'Medida', 'Un.'], axis=1)

    # A margem de lucro de Herbicidas é de 22%
    # A margem de lucro de Fungicidas é de 14%
    # A margem de lcuro de Inseticidas é de 31%
    # Demais margens são 10%
    
    def determinar_margem_lucro(produto):
        if 'Herbicidas' in produto:
            return 0.22
        elif 'Fungicidas' in produto:
            return 0.14
        elif 'Inseticidas' in produto:
            return 0.31
        else:
            return 0.10  
    
    df['Margem Lucro'] = df['Produtos'].apply(determinar_margem_lucro)
    
    df['Valor Unitário Ajustado'] = df['Valor Unitário'] * (1 + df['Margem Lucro'])
    
    df['Valor Total Ajustado'] = df['Quantidade'] * df['Valor Unitário Ajustado']
    
    df = df.drop(['Valor Unitário', 'Valor Total', 'Margem Lucro'], axis=1)
    
    df = df.rename(columns={'Valor Unitário Ajustado': 'Valor Unitário', 'Valor Total Ajustado': 'Valor Total'})
    
    nova_ordem_colunas = ['Produtos', 'Quantidade', 'Valor Unitário', 'Valor Total','Un. Medida']
    
    df = df[nova_ordem_colunas]
    
    df['Valor Unitário'] = df['Valor Unitário'].astype(float).round(2)
    df['Valor Total'] = df['Valor Total'].astype(float).round(2)
    
    df[['Un. Medida_Numero', 'Un. Medida_Medida']] = df['Un. Medida'].str.extract(r'(\d+\.?\d*)\s?(\D+)')
    
    # Converta 'Un. Medida_Numero' para float
    df['Un. Medida_Numero'] = df['Un. Medida_Numero'].astype(float)
    
    # Aplicando as condições
    condicao = (df['Un. Medida_Numero'] > 999) & ((df['Un. Medida_Medida'] == 'g') | (df['Un. Medida_Medida'] == 'ml'))
    
    # Aplicando as conversões corrigidas
    df.loc[condicao, 'Un. Medida'] = df.loc[condicao].apply(lambda row: f'{row["Un. Medida_Numero"]/1000} kg' if row["Un. Medida_Medida"] == 'g' else f'{row["Un. Medida_Numero"]/1000} litro(s)', axis=1)
    
    # Dropando colunas auxiliares
    df.drop(['Un. Medida_Numero', 'Un. Medida_Medida'], axis=1, inplace=True)
    
    original_ordem_colunas = df.columns.tolist()
    
    total_linha = pd.DataFrame({'Produtos': 'Total', 'Un.': '', 'Valor Unitário': '', 'Unidades': '',
                              'Valor Total': df['Valor Total'].sum()}, index=[len(df)])
    
    total_linha.loc[:, df.columns.difference(['Produtos', 'Valor Total'])] = ''
    
    df = pd.concat([df, total_linha])
    
    df = df[original_ordem_colunas]
    
    return df 


def inicializar_tabela_cabecalho(vendedor_escolhido, cliente_escolhido):
    global tabela_cabecalho
    tabela_cabecalho = Table([
        [Image(caminho_logotipo, width=70, height=55, hAlign='LEFT'),
         f'Vendedor: {vendedor_escolhido}\nCNPJ Empresa: 11121112/0001-29\nEndereço Empresa: Sorocaba',
         f'Cliente: {cliente_escolhido}\nCNPJ: 11111111/0001-11\nEndereço: {get_endereco(cliente_escolhido)}']
    ])

# Parte 2: Criação do PDF com base no DataFrame
def processar_e_gerar_pdf(df, vendedor_escolhido, cliente_escolhido):
    try:
        # Inicializar a tabela de cabeçalho antes de utilizá-la
        inicializar_tabela_cabecalho(vendedor_escolhido, cliente_escolhido)

        caminho_salvar = filedialog.asksaveasfilename(
            title="Salvar orçamento",
            filetypes=[("Arquivos PDF", "*.pdf")],
            defaultextension=".pdf"
        )

        if caminho_salvar:
            # Continuar com a geração do PDF
            gerar_pdf(df, caminho_salvar, vendedor_escolhido, cliente_escolhido)
            messagebox.showinfo("Sucesso", f"Tabela exportada para {caminho_salvar} com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar o PDF: {e}")
        messagebox.showerror("Erro", f"Erro ao gerar o PDF: {e}")

#| ==========================================================================================

def get_endereco(cliente_nome):
    for cliente in clientes:
        if cliente["nome"] == cliente_nome:
            return cliente["endereco"]
    return ""

def gerar_pdf(df, caminho_pdf, vendedor_escolhido, cliente_escolhido):
    print(f"Gerando PDF em: {caminho_pdf}")

    doc = SimpleDocTemplate(caminho_pdf, pagesize=letter)
    story = []

    inicializar_tabela_cabecalho(vendedor_escolhido, cliente_escolhido)

    style_cabecalho = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ])

    tabela_cabecalho.setStyle(style_cabecalho)
    story.append(tabela_cabecalho)
    story.append(Spacer(1, 12))
    data_precos = [df.columns.tolist()] + df.values.tolist()
    table_precos = Table(data_precos)

    style_precos = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    table_precos.setStyle(style_precos)
    story.append(table_precos)
    doc.build(story)
    
#gerar_pdf(df, caminho_salvar_pdf, caminho_logotipo)


```

# Tkinter 

```bash

# Interface Gráfica (Tkinter)
janela = tk.Tk()
janela.title("Processador de Planilha")

frame_botoes = tk.Frame(janela)

vendedores = [" ", "Jones Karlos Manoel", "Adriana Motta", "Viviane Fitipaldi"]

clientes = [
    {"nome": " ", "cnpj": " ", "endereco": " "},
    {"nome": "Ytaburi Agro", "cnpj": "11111111/0001-11", "endereco": "Presidente Prudente"},
    {"nome": "Fazenda Moinho de vento", "cnpj": "11111111/0001-12", "endereco": "Itu"},
    {"nome": "Boiadero do Dudu", "cnpj": "11111111/0001-13", "endereco": "Betim"}
]

vendedor_var = StringVar(janela)
vendedor_var.set(vendedores[0])

cliente_var = StringVar(janela)
cliente_var.set(clientes[0]["nome"])

# Adicione rótulos acima dos Comboboxes
rotulo_vendedor = tk.Label(frame_botoes, text="Vendedor(a):")
rotulo_cliente = tk.Label(frame_botoes, text="Cliente:")

menu_vendedor = ttk.Combobox(frame_botoes, textvariable=vendedor_var, values=vendedores, state="readonly", width=20)
menu_cliente = ttk.Combobox(frame_botoes, textvariable=cliente_var, values=[cliente["nome"] for cliente in clientes], state="readonly", width=20)

botao_processar = tk.Button(frame_botoes, text="Processar", command=processar_planilha)

frame_botoes.pack(padx=5, pady=5)

# Rótulo e Menu para o Vendedor
rotulo_vendedor = tk.Label(frame_botoes, text="Selecione o(a) vendedor(a):")
rotulo_vendedor.pack(side=tk.LEFT, padx=5)
menu_vendedor.pack(side=tk.LEFT, padx=5)

# Rótulo e Menu para o Cliente
rotulo_cliente = tk.Label(frame_botoes, text="Selecione o cliente:")
rotulo_cliente.pack(side=tk.LEFT, padx=5)
menu_cliente.pack(side=tk.LEFT, padx=5)

botao_processar.pack(side=tk.LEFT, padx=5)

janela.mainloop()
```

yyy

```bash
select * from xyz
```

## Próximos passos:

* Em um cenário real, uma equipe comercial frequentemente não se restringe a criar cotações de apenas um fornecedor; eles alternam entre fornecedores ou, em alguns casos, consolidam um orçamento único proveniente de diversas fontes. Considerando a evolução deste projeto, planejo a inclusão da capacidade de transformar a planilha a partir de múltiplas fontes de fornecimento para gerar um orçamento final, seja ele uma combinação de vários fornecedores ou não.
  
* Criar uma calculadora de impostos sobre os valores dos itens.

* Criar uma calculadora de frete com base na cidade de coleta e entrega, com critérios de precificação baseados em peso e km.

* Criar a possibilidade do usuário enviar o e-mail para o cliente pelo executável
