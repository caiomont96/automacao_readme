### Sistema de geração de cotação automatizado.
.
.

“ Este projeto é um reflexo de desafios que já tive em ambiente de trabalho real, entretanto, o formato do desafio e todos os dados, incluindo o nome da empresa, clientes, margem de lucro e produtos, são fictícios. “

.
.

Em uma empresa do setor de agronegócio, o time de marketing e vendas está experimentando um aumento significativo na demanda, resultando em uma enxurrada de solicitações de orçamento para produtos agrícolas.
Contudo, esse aumento nas oportunidades também trouxe consigo um novo problema. 

Anteriormente, quando as solicitações eram menos numerosas, os vendedores conseguiam lidar eficientemente, processando a planilha do fornecedor, realizando as adaptações necessárias no Excel e enviando os orçamentos em PDF para cada potencial comprador.

No entanto, com a crescente demanda, os vendedores agora enfrentam o desafio de dedicar mais tempo à geração de orçamentos do que à negociação em si. Mesmo com esses esforços, muitas vezes não conseguem cumprir os prazos estipulados, resultando na perda de clientes devido a atrasos no retorno.

Diante desse cenário, a equipe de dados recebeu o desafio de desenvolver uma solução de automação para esse processo. 

O objetivo é transformar a planilha do fornecedor em um documento final em PDF, incorporando dados do cliente, conversões de unidade de medida e margens de lucro, de maneira rápida e eficiente, permitindo que os vendedores foquem no que fazem de melhor: negociar e atender aos clientes.

O vídeo abaixo mostra como o oçamento era feito manualmente antes da solução:

## VIDEO

# Dividindo a solução em passos:

* O fornecedor costuma enviar a planilha em um formato horizontal, onde as características estão na horizontal e os produtos na vertical, essa planilha deve ser transposta.
  
* A coluna [Un. Medida] precisa ser alterada com a sua medida multiplicada por sua quantidade. O resultado disso por sua vez precisa ser convertido (ex: 2000g = 2kg)

* Precisará ser inclusa a coluna [Valor Total] multiplicando a coluna [Quantidade] e [Valor Unitário]

* Será necessário incluir a margem de lucro nas colunas [Valor Unitário] e consequentemente [Valor Total]

* Precisará ser criada na última linha o ‘Total’ do valor final que o cliente irá pagar.

* O PDF final precisa ter a logo da empresa no canto superior esquerdo e os dados da empresa e do cliente em cima da tabela.

* Após processar a tabela e criar o pdf, criar um executável (uma janela de interação) para ser usado pelo vendedores.

  ---


# Primeiro passo

O primeiro passo é inserir as bibliotecas:

* Pandas para manipulação da planilha xlsx
  
* Reportlab para conversão em pdf

* Tkinter para criar uma interface gráfica para o usuário inserir a planilha e exportar em pdf em uma janela

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
```

Inserção da planilha

```bash

df = pd.read_excel('fornecedor_agro.xlsx')

```
# formato inicial da planilha

A planilha chega assim:

|      | Produtos                               | Herbicidas Glyphosate | Herbicidas Paraquat | Herbicidas Atrazine | ... |
|------|-------------------------------------- | ---------------------- | ------------------- | ------------------- | --- |
| 0    | Descrição                              | Embalagem de 1 litro(s) | Embalagem de 1 litro(s) | Embalagem de 1 litro(s) | ... |
| 1    | Unidades                              | 90                     | 77                  | 11                  | ... |
| 2    | Valor Unitário                        | 67                     | 78                  | 34                  | ... |
| ...  | ...                                  | ...                    | ...                 | ...                 | ... |



precisaremos transpor.

```bash
# A linha de código "df = df.T" transpõe (inverte linhas e colunas) o DataFrame.
# Isso significa que as linhas do DataFrame original agora se tornam colunas e vice-versa.

df = df.T
```

Ao transpor, a planilha fica dessa forma

|                            | 0                         | 1          | 2              |
|----------------------------|---------------------------|------------|----------------|
| Produtos                   | Un. Medida                | Quantidade  | Valor Unitário |
| Herbicidas Glyphosate      | Embalagem de 1 litro(s)   | 90          | 67             |
| Herbicidas Paraquat         | Embalagem de 1 litro(s)   | 77         | 78             |
| Herbicidas Atrazine         | Embalagem de 1 litro(s)   | 11         | 34             |
| Fungicidas Mancozeb         | Embalagem de 500g         | 21         | 45.67          |
| Fungicidas Azoxystrobin     | Embalagem de 500g         | 13         | 45.99          |
| Fungicidas Tebuconazole     | Embalagem de 500g         | 5          | 95.98          |

os produtos viraram o índice e o cabeçalho não é interpretado como cabeçalho e sim como primeira linha.

O código abaixo ajusta:

```bash

# Redefinindo os índices do DataFrame.
df = df.reset_index()

# Redefinindo os nomes das colunas para serem iguais aos valores da primeira linha do DataFrame.
df.columns = df.iloc[0]

# Reatribuindo ao DataFrame todos os dados, excluindo a primeira linha, que agora contém os nomes das colunas.
df = df[1:]

df

```
# formato da planilha transposta


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

Agora temos mais clareza sobre como está a planilha ao verticalizá-la e podemos seguir com o passo a passo.

  ---

Un. Medida terá uma alteração forte da forma que está até sua transformação, vou usar essa linha como exemplo:

| Produtos                        | Un. Medida                       | Quantidade | Valor Unitário |
|---------------------------------|----------------------------------|------------|----------------|
| Fungicidas Mancozeb             | Embalagem de 500g                | 21         | 45.67          |

São 21 unidades com 500g cada embalagem, a coluna Un. Medida irá sumir e deverá ser totalizado os 10500g na nova coluna de Un. Medida.

Estes 10500g por sua vez serão convertidos em 10,5 kg

Primeiramente, vamos tirar as frases "Embalagem de", "Saco de", "Frasco de" e posteriormente separar o numero e sua unidade de medida (ex: 500 e g)
para podermos tratar essa coluna de forma numérica e não como string.

(A partir daqui, por questões práticas e estéticas, não mostrarei todas as linhas da tabela mas apenas as primeiras ou mais importantes para a interpretação do que está acontecendo)

```bash

# Os padrões são 'Embalagem de', 'Saco de', e 'Frasco de', e eles são substituídos por uma string vazia ('') pelo método str.replace().

df['Un. Medida'] = df['Un. Medida'].str.replace('Embalagem de|Saco de|Frasco de', '', regex=True)
```

| Produtos                                      | Un. Medida                         | Quantidade | Valor Unitário |
|-----------------------------------------------|------------------------------------|------------|----------------|
| Herbicidas Glyphosate                         | 1 litro(s)                         | 90         | 67             |
| Herbicidas Paraquat                           | 1 litro(s)                         | 77         | 78             |
| Fungicidas Tebuconazole                       | 500g                               | 5          | 95.98          |
| Inseticidas Imidacloprid                      | 250ml                              | 32         | 46.63          |

O código abaixo vai separar

```bash

# Extraindo informações de medidas numéricas e unidades da coluna 'Un. Medida' e as atribui a novas colunas 'Medida' e 'Un.'.
# A expressão regular (\d+\.?\d*)\s*([a-zA-Z\(s\)]*) captura um ou mais dígitos (com ou sem ponto decimal) como 'Medida' e letras (com parênteses) como 'Un.'.
df[['Medida', 'Un.']] = df['Un. Medida'].str.extract('(\d+\.?\d*)\s*([a-zA-Z\(s\)]*)')

# Removendo a coluna original 'Un. Medida' do DataFrame df.
df = df.drop('Un. Medida', axis=1)

# convertendo a coluna 'Medida' para valores numéricos, 1ualquer valor que não possa ser convertido é definido como NaN (Not a Number) devido ao parâmetro errors='coerce'.
df['Medida'] = pd.to_numeric(df['Medida'], errors='coerce')

```

A partir daqui, a planilha Vai ter 'Medida' e 'Un.' separados, um como número int. e outro como string.


| Produtos                   | Quantidade | Valor Unitário | Medida | Un.     |
|----------------------------|------------|----------------|--------|---------|
| Herbicidas Glyphosate      | 90         | 67             | 1      | litro(s) |
| Herbicidas Paraquat         | 77         | 78             | 1      | litro(s) |
| Herbicidas Atrazine         | 11         | 34             | 1      | litro(s) |
| Fungicidas Mancozeb         | 21         | 45.67          | 500    | g       |
| Fungicidas Azoxystrobin     | 13         | 45.99          | 500    | g       |
| Fungicidas Tebuconazole     | 5          | 95.98          | 500    | g       |
| Inseticidas Imidacloprid    | 32         | 46.63          | 250    | ml      |

Agora, vamos criar uma coluna chamada Valor Total, que contém a Quantidade multiplicada pelo Valor Unitário e uma chamada Quantidade Total eu multiplica Quantidade e Medida

```bash

df['Valor Total'] = df['Quantidade'] * df['Valor Unitário']
df['Quantidade Total'] = df['Quantidade'] * df['Medida']

```

| Produtos                   | Quantidade | Valor Unitário | Medida | Un.     | Valor Total | Quantidade Total |
|----------------------------|------------|----------------|--------|---------|-------------|-------------------|
| Herbicidas Glyphosate      | 90         | 67             | 1      | litro(s) | 6030        | 90                |
| Herbicidas Paraquat         | 77         | 78             | 1      | litro(s) | 6006        | 77                |
| Herbicidas Atrazine         | 11         | 34             | 1      | litro(s) | 374         | 11                |
| Fungicidas Mancozeb         | 21         | 45.67          | 500    | g       | 959.07      | 10500             |
| Fungicidas Azoxystrobin     | 13         | 45.99          | 500    | g       | 597.87      | 6500              |
| Fungicidas Tebuconazole     | 5          | 95.98          | 500    | g       | 479.9       | 2500              |
| Inseticidas Imidacloprid    | 32         | 46.63          | 250    | ml      | 1492.16     | 8000              |


Agora, vamos re-criar a coluna 'Un. Medida' concatenando a coluna Quantidade Total e de Un.

Vamos aproveitar e excluir as colunas Quantidade Total, Medida e Un., suas funções eram auxiliares e não farão parte o produto final.

```bash

# Criando uma nova coluna 'Un. Medida' concatenando a coluna 'Quantidade Total' convertida para string com a coluna 'Un.'.
df['Un. Medida'] = df['Quantidade Total'].astype(str) + ' ' + df['Un.']

# Removendo as colunas 'Quantidade Total', 'Medida', e 'Un.' do DataFrame que estavam sendo usadas como auxiliares.
df = df.drop(['Quantidade Total', 'Medida', 'Un.'], axis=1)

```

| Produtos                   | Quantidade | Valor Unitário | Valor Total | Un. Medida    |
|----------------------------|------------|----------------|-------------|---------------|
| Herbicidas Glyphosate      | 90         | 67             | 6030        | 90 litro(s)   |
| Herbicidas Paraquat        | 77         | 78             | 6006        | 77 litro(s)   |
| Herbicidas Atrazine        | 11         | 34             | 374         | 11 litro(s)   |
| Fungicidas Mancozeb        | 21         | 45.67          | 959.07      | 10500 g       |
| Fungicidas Azoxystrobin    | 13         | 45.99          | 597.87      | 6500 g        |
| Fungicidas Tebuconazole    | 5          | 95.98          | 479.9       | 2500 g        |
| Inseticidas Imidacloprid   | 32         | 46.63          | 1492.16     | 8000 ml       |


```bash

# A função determinar_margem_lucro(produto) recebe o nome do produto como entrada e retorna a margem de lucro correspondente.
# A lógica é definida por categorias de produtos: Herbicidas têm margem de 8%, Fungicidas têm margem de 11%, Inseticidas têm margem de 13%, e os demais têm margem de 7%.

# A linha de código "df['Margem Lucro'] = df['Produtos'].apply(determinar_margem_lucro)"
# aplica a função determinar_margem_lucro a cada valor na coluna 'Produtos' do DataFrame df e cria uma nova coluna 'Margem Lucro' com os resultados.


# A margem de lucro de Herbicidas é de 8%
# A margem de lucro de Fungicidas é de 11%
# A margem de lcuro de Inseticidas é de 13%
# Demais margens são 7%

def determinar_margem_lucro(produto):
    if 'Herbicidas' in produto:
        return 0.08
    elif 'Fungicidas' in produto:
        return 0.11
    elif 'Inseticidas' in produto:
        return 0.13
    else:
        return 0.07  

df['Margem Lucro'] = df['Produtos'].apply(determinar_margem_lucro)

```

| Produtos                   | Quantidade | Valor Unitário | Medida  | Un.  | Valor Total | Quantidade Total | Margem Lucro |
|----------------------------|------------|----------------|---------|------|-------------|-------------------|--------------|
| Herbicidas Glyphosate      | 90         | 67             | 1       | litro(s) | 6030        | 90                | 0.08         |
| Herbicidas Paraquat        | 77         | 78             | 1       | litro(s) | 6006        | 77                | 0.08         |
| Herbicidas Atrazine        | 11         | 34             | 1       | litro(s) | 374         | 11                | 0.08         |
| Fungicidas Mancozeb        | 21         | 45.67          | 500     | g    | 959.07      | 10500             | 0.11         |
| Fungicidas Azoxystrobin    | 13         | 45.99          | 500     | g    | 597.87      | 6500              | 0.11         |
| Fungicidas Tebuconazole    | 5          | 95.98          | 500     | g    | 479.9       | 2500              | 0.11         |
| Inseticidas Imidacloprid   | 32         | 46.63          | 250     | ml   | 1492.16     | 8000              | 0.13         |

Agora aplicando:

```bash

# Calculando o valor unitário ajustado para cada linha multiplicando o valor unitário original pelo fator (1 + Margem de Lucro).
df['Valor Unitário Ajustado'] = df['Valor Unitário'] * (1 + df['Margem Lucro'])

# Calculando o valor total ajustado para cada linha multiplicando a quantidade pela coluna 'Valor Unitário Ajustado'.
df['Valor Total Ajustado'] = df['Quantidade'] * df['Valor Unitário Ajustado']
# Essas operações refletem o ajuste dos preços com base nas margens de lucro específicas de cada categoria de produto.

```

| Produtos                   | Quantidade | Valor Unitário | Medida | Un. | Valor Total | Quantidade Total | Margem Lucro | Valor Unitário Ajustado | Valor Total Ajustado |
|----------------------------|------------|----------------|--------|-----|-------------|-------------------|--------------|-------------------------|-----------------------|
| Herbicidas Glyphosate      | 90         | 67             | 1      | litro(s) | 6030        | 90                | 0.08         | 72.36                   | 6512.4                |
| Herbicidas Paraquat        | 77         | 78             | 1      | litro(s) | 6006        | 77                | 0.08         | 84.24                   | 6486.48               |
| Herbicidas Atrazine        | 11         | 34             | 1      | litro(s) | 374         | 11                | 0.08         | 36.72                   | 403.92                |
| Fungicidas Mancozeb        | 21         | 45.67          | 500    | g   | 959.07      | 10500             | 0.11         | 50.6937                 | 1064.5677             |
| Fungicidas Azoxystrobin    | 13         | 45.99          | 500    | g   | 597.87      | 6500              | 0.11         | 51.0489                 | 663.6357              |
| Fungicidas Tebuconazole    | 5          | 95.98          | 500    | g   | 479.9       | 2500              | 0.11         | 106.5378                | 532.689               |
| Inseticidas Imidacloprid   | 32         | 46.63          | 250    | ml  | 1492.16     | 8000              | 0.13         | 52.6919                 | 1686.1408             |


Arrumando as várias colunas

```bash

# Removendo as colunas 'Valor Unitário', 'Valor Total', e 'Margem Lucro' do DataFrame df.
df = df.drop(['Valor Unitário', 'Valor Total', 'Margem Lucro'], axis=1)

# Renomeando as colunas 'Valor Unitário Ajustado' para 'Valor Unitário' e 'Valor Total Ajustado' para 'Valor Total'.
df = df.rename(columns={'Valor Unitário Ajustado': 'Valor Unitário', 'Valor Total Ajustado': 'Valor Total'})

nova_ordem_colunas = ['Produtos', 'Quantidade', 'Valor Unitário', 'Valor Total','Un. Medida']

# Reorganizando as colunas do DataFrame df na ordem especificada por nova_ordem_colunas.
df = df[nova_ordem_colunas]

```
| Produtos                   | Quantidade | Valor Unitário | Valor Total | Un. Medida    |
|----------------------------|------------|----------------|-------------|---------------|
| Herbicidas Glyphosate      | 90         | 72.36          | 6512.4      | 90 litro(s)   |
| Herbicidas Paraquat        | 77         | 84.24          | 6486.48     | 77 litro(s)   |
| Herbicidas Atrazine        | 11         | 36.72          | 403.92      | 11 litro(s)   |
| Fungicidas Mancozeb        | 21         | 50.6937        | 1064.5677   | 10500 g       |
| Fungicidas Azoxystrobin    | 13         | 51.0489        | 663.6357    | 6500 g        |
| Fungicidas Tebuconazole    | 5          | 106.5378       | 532.689     | 2500 g        |
| Inseticidas Imidacloprid   | 32         | 52.6919        | 1686.1408   | 8000 ml       |

```bash

# As linhas "df['Valor Unitário'] = df['Valor Unitário'].astype(float).round(2)" e "df['Valor Total'] = df['Valor Total'].astype(float).round(2)"
# convertem as colunas 'Valor Unitário' e 'Valor Total' para o tipo de dado float e arredondam os valores para duas casas decimais.
# Isso garante que essas colunas tenham valores numéricos formatados corretamente.


df['Valor Unitário'] = df['Valor Unitário'].astype(float).round(2)
df['Valor Total'] = df['Valor Total'].astype(float).round(2)
```
| Produtos                   | Quantidade | Valor Unitário | Valor Total | Un. Medida    |
|----------------------------|------------|----------------|-------------|---------------|
| Herbicidas Glyphosate      | 90         | 72.36          | 6512.40     | 90 litro(s)   |
| Herbicidas Paraquat        | 77         | 84.24          | 6486.48     | 77 litro(s)   |
| Herbicidas Atrazine        | 11         | 36.72          | 403.92      | 11 litro(s)   |
| Fungicidas Mancozeb        | 21         | 50.69          | 1064.57     | 10500 g       |
| Fungicidas Azoxystrobin    | 13         | 51.05          | 663.64      | 6500 g        |
| Fungicidas Tebuconazole    | 5          | 106.54         | 532.69      | 2500 g        |
| Inseticidas Imidacloprid   | 32         | 52.69          | 1686.14     | 8000 ml       |

```bash

# extrai informações numéricas e alfabéticas da coluna 'Un. Medida' e atribui a esses valores as novas colunas 'Un. Medida_Numero' e 'Un. Medida_Medida'.

df[['Un. Medida_Numero', 'Un. Medida_Medida']] = df['Un. Medida'].str.extract(r'(\d+\.?\d*)\s?(\D+)')

#converte a coluna 'Un. Medida_Numero' para o tipo de dado float.
df['Un. Medida_Numero'] = df['Un. Medida_Numero'].astype(float)

# cria uma condição para identificar linhas que atendem a determinados critérios (número maior que 999 e medida em gramas ou mililitros).

condicao = (df['Un. Medida_Numero'] > 999) & ((df['Un. Medida_Medida'] == 'g') | (df['Un. Medida_Medida'] == 'ml'))

# aplica uma função lambda às linhas que atendem à condição. Essa função converte as medidas de gramas para quilogramas ou mililitros para litros, conforme apropriado.

df.loc[condicao, 'Un. Medida'] = df.loc[condicao].apply(lambda row: f'{row["Un. Medida_Numero"]/1000} kg' if row["Un. Medida_Medida"] == 'g' else f'{row["Un. Medida_Numero"]/1000} litro(s)', axis=1)

# remove as colunas auxiliares 'Un. Medida_Numero' e 'Un. Medida_Medida'.

df.drop(['Un. Medida_Numero', 'Un. Medida_Medida'], axis=1, inplace=True)

```

| Produtos                                     | Quantidade | Valor Unitário | Valor Total | Un. Medida         |
|----------------------------------------------|------------|----------------|-------------|--------------------|
| Herbicidas Glyphosate                        | 90         | 72.36          | 6512.40     | 90 litro(s)        |
| Herbicidas Paraquat                          | 77         | 84.24          | 6486.48     | 77 litro(s)        |
| Herbicidas Atrazine                          | 11         | 36.72          | 403.92      | 11 litro(s)        |
| Fungicidas Mancozeb                          | 21         | 50.69          | 1064.57     | 10.5 kg            |
| Fungicidas Azoxystrobin                      | 13         | 51.05          | 663.64      | 6.5 kg             |
| Fungicidas Tebuconazole                      | 5          | 106.54         | 532.69      | 2.5 kg             |
| Inseticidas Imidacloprid                     | 32         | 52.69          | 1686.14     | 8.0 litro(s)       |

```bash

# Exibindo o DataFrame atualizado
# Salva a ordem original das colunas do DataFrame.
original_ordem_colunas = df.columns.tolist()

# cria um novo DataFrame contendo uma linha total com a soma dos valores da coluna 'Valor Total'.

total_linha = pd.DataFrame({'Produtos': 'Total', 'Un.': '', 'Valor Unitário': '', 'Unidades': '',
                          'Valor Total': df['Valor Total'].sum()}, index=[len(df)])

# preenche todas as colunas, exceto 'Produtos' e 'Valor Total', da linha total com valores vazios.
total_linha.loc[:, df.columns.difference(['Produtos', 'Valor Total'])] = ''

# concatena o DataFrame original com a linha total.
df = pd.concat([df, total_linha])

# reorganiza as colunas do DataFrame df para a ordem original salva anteriormente.
df = df[original_ordem_colunas]
```
  ---

Abaixo, a tabela está pronta e dentro do formato que o cliente exige receber, com margens de lucro embutidas e suas respectivas unidades de Medida convertidas com base na Quantidade.


|  Produtos                                    | Quantidade | Valor Unitário | Valor Total | Un. Medida         |
|----------------------------------------------|------------|----------------|-------------|--------------------|
| Herbicidas Glyphosate                        | 90         | 72.36          | 6512.40     | 90 litro(s)        |
| Herbicidas Paraquat                          | 77         | 84.24          | 6486.48     | 77 litro(s)        |
| Herbicidas Atrazine                          | 11         | 36.72          | 403.92      | 11 litro(s)        |
| Fungicidas Mancozeb                          | 21         | 50.69          | 1064.57     | 10.5 kg            |
| Fungicidas Azoxystrobin                      | 13         | 51.05          | 663.64      | 6.5 kg             |
| Fungicidas Tebuconazole                      | 5          | 106.54         | 532.69      | 2.5 kg             |
| Inseticidas Imidacloprid                     | 32         | 52.69          | 1686.14     | 8.0 litro(s)       |
| Inseticidas Lambda-cyhalothrin               | 5          | 84.38          | 421.89      | 1.25 litro(s)      |
| Inseticidas Chlorpyrifos                     | 7          | 103.21         | 722.50      | 1.75 litro(s)      |
| Fertilizantes Nitrato de amônio              | 17         | 90.78          | 1543.24     | 425 kg             |
| Fertilizantes Fosfato diamônico (DAP)        | 9          | 112.48         | 1012.31     | 225 kg             |
| Fertilizantes Cloreto de potássio            | 45         | 124.0          | 5580.10     | 1125 kg            |
| Reguladores de crescimento: Ácido giberélico | 23         | 18.79          | 432.15      | 230 g              |
| Reguladores de crescimento: Paclobutrazol    | 43         | 25.09          | 1078.93     | 430 g              |
| Reguladores de crescimento: Ethephon         | 18         | 37.95          | 683.15      | 180 g              |
| Adjuvantes Óleo mineral                      | 8          | 35.56          | 284.45      | 4.0 litro(s)       |
| Adjuvantes Surfactantes                      | 32         | 45.65          | 1460.68     | 16.0 litro(s)      |
| Adjuvantes Espalhantes adesivos              | 19         | 59.2           | 1124.86     | 9.5 litro(s)       |
| Total                                        |            |                | 31694.10    |                    |

  ---

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

Até o momento, o código está funcional, mas sua aplicação é restrita ao meu ambiente de compilação, carecendo de uma interface mais abrangente para ser utilizada como uma ferramenta por um vendedor.

Abaixo, irei resumir as ações em funções e incluir e criar uma aplicação pela biblioteca Tkinter

  ---

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
  ---

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

  ---

# Vídeo conclusão

# VIDEO

## Próximos passos:

* Em um cenário real, uma equipe comercial frequentemente não se restringe a criar cotações de apenas um fornecedor; eles alternam entre fornecedores ou, em alguns casos, consolidam um orçamento único proveniente de diversas fontes de fornecedor. Considerando a evolução deste projeto, planejo a inclusão da capacidade de transformar a planilha a partir de múltiplas fontes de fornecimento para gerar um orçamento final, seja ele uma combinação de vários fornecedores ou não.
  
* Criar uma calculadora de impostos (ex: pis/cofins) sobre os valores dos itens.

* Criar uma calculadora de frete com base na cidade de coleta e entrega, com critérios de precificação baseados em peso e km.

* Criar a possibilidade do usuário enviar o e-mail para o cliente diretamente pelo executável
