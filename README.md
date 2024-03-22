# Sistema de geração de cotação automatizado.
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

O vídeo abaixo mostra como o oçamento é feito manualmente antes da solução ser implementada:

## VIDEO

[![Assista ao vídeo](https://img.youtube.com/watch?v=HnrVCK9Mu-M/0.jpg)](https://www.youtube.com/watch?v=HnrVCK9Mu-M)

 ---

## Estratégia de Solução em Etapas

* A solução será desenvolvida no Jupyter, uma aplicação de código aberto que permite analisar e transformar dados.

* O fornecedor costuma enviar a planilha em um formato horizontal, onde as características estão na horizontal e os produtos na vertical, essa planilha deve ser transposta.

* Ao ser transposta, temos 4 colunas: [Produto], [Un. Medida], [Quantidade] e [Valor Unitário]
  
* A coluna [Un. Medida] precisa ser alterada com a sua medida multiplicada por sua quantidade. O resultado disso por sua vez precisa ser convertido (ex: 2000g = 2kg)

* Precisará ser inclusa a coluna [Valor Total] multiplicando a coluna [Quantidade] e [Valor Unitário]

* Será necessário incluir a margem de lucro nas colunas [Valor Unitário] e consequentemente [Valor Total]

* Precisará ser criada na última linha o ‘Total’ do valor final que o cliente irá pagar, somando a coluna [Valor Total].

* Após estas trasnformações, as colunas precisarão ser  [Produtos], [Quantidade], [Valor Unitário], [Valor Total], [Un. Medida] nesta ordem. 

* O PDF final precisa ter a logo da empresa no canto superior esquerdo e os dados da empresa e do cliente em cima da tabela.

* Após processar a tabela e criar o pdf, criar um executável (uma janela de interação) para a automação ser executada pelo vendedores.

* Este executável será possível convertendo o código de _.ipkernel_ para _.py_ e gerado no repositório PyPI.

  ---


## Primeiro passo

No primeiro passo, configuramos as bibliotecas necessárias para o processo:

* Pandas: Utilizado para a manipulação de planilhas Excel (xlsx).

* Reportlab: Essencial para a conversão dos dados em formato PDF.

* Tkinter: Utilizado para criar uma interface gráfica intuitiva, permitindo ao usuário inserir a planilha e exportar para PDF em uma janela interativa.

Este conjunto de bibliotecas é fundamental para garantir a fu

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
## formato inicial da planilha

A leitura da planilha é conduzida desta maneira:


|      | Produtos                               | Herbicidas Glyphosate | Herbicidas Paraquat | Herbicidas Atrazine | ... |
|------|-------------------------------------- | ---------------------- | ------------------- | ------------------- | --- |
| 0    | Descrição                              | Embalagem de 1 litro(s) | Embalagem de 1 litro(s) | Embalagem de 1 litro(s) | ... |
| 1    | Unidades                              | 90                     | 77                  | 11                  | ... |
| 2    | Valor Unitário                        | 67                     | 78                  | 34                  | ... |
| ...  | ...                                  | ...                    | ...                 | ...                 | ... |



Precisaremos transpor de forma que as celulas da primeira linha se tornem os cabeçalhos.

```bash
# A linha de código "df = df.T" transpõe (inverte linhas e colunas) o DataFrame.
# Isso significa que as linhas do DataFrame original agora se tornam colunas e vice-versa.

df = df.T
```

Ao transpor, a planilha fica dessa forma:

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

_A partir daqui, por questões práticas e estéticas, não mostrarei todas as linhas da tabela mas apenas as primeiras ou mais importantes para a interpretação do que está acontecendo_

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

# Extraindo informações da coluna [Un. Medida] e as atribui a novas colunas 'Medida' e 'Un.'.
# Separando número e palavra com o str.extract('(\d+\.?\d*)\s*([a-zA-Z\(s\)]*)'), uma expressão regular projetada para extrair números decimais seguidos de unidades de medida de uma string.

df[['Medida', 'Un.']] = df['Un. Medida'].str.extract('(\d+\.?\d*)\s*([a-zA-Z\(s\)]*)')

# Removendo a coluna original 'Un. Medida' do DataFrame df.

df = df.drop('Un. Medida', axis=1)

# convertendo a coluna 'Medida' para valores numéricos, qualquer valor que não possa ser convertido é definido como NaN (Not a Number) devido ao parâmetro errors='coerce'.

df['Medida'] = pd.to_numeric(df['Medida'], errors='coerce')

```

A partir daqui, a planilha Vai ter 'Medida' e 'Un.' separados, um como número int e outro como string.


| Produtos                   | Quantidade | Valor Unitário | Medida | Un.     |
|----------------------------|------------|----------------|--------|---------|
| Herbicidas Glyphosate      | 90         | 67             | 1      | litro(s) |
| Herbicidas Paraquat         | 77         | 78             | 1      | litro(s) |
| Herbicidas Atrazine         | 11         | 34             | 1      | litro(s) |
| Fungicidas Mancozeb         | 21         | 45.67          | 500    | g       |
| Fungicidas Azoxystrobin     | 13         | 45.99          | 500    | g       |
| Fungicidas Tebuconazole     | 5          | 95.98          | 500    | g       |
| Inseticidas Imidacloprid    | 32         | 46.63          | 250    | ml      |

Agora, vamos criar uma coluna chamada [Valor Total], que contém a Quantidade multiplicada pelo [Valor Unitário] e uma chamada [Quantidade Total] eu multiplica [Quantidade] e [Medida]

```bash

# Calculando o valor total multiplicando a quantidade pelo valor unitário
df['Valor Total'] = df['Quantidade'] * df['Valor Unitário']

# Calculando a quantidade total multiplicando a quantidade pela medida
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

Inserindo as margens de lucro.

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


Ajustando colunas:

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

Agora precisaremos resumir os centavos duas casas após a vírgula:

```bash
# Convertendo as colunas 'Valor Unitário' e 'Valor Total' para o tipo de dado float e arredondam os valores para duas casas decimais.

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

Vamos garantir que unidades como 'gramas' e 'mililitros', que excedam 999, sejam  convertidas para 'Kg' e 'litro(s)', respectivamente.

```bash

# Extraindo informações numéricas e alfabéticas da coluna 'Un. Medida' e atribuindo a esses valores as novas colunas 'Un. Medida_Numero' e 'Un. Medida_Medida'.

df[['Un. Medida_Numero', 'Un. Medida_Medida']] = df['Un. Medida'].str.extract(r'(\d+\.?\d*)\s?(\D+)')

#converte a coluna 'Un. Medida_Numero' para o tipo de dado float.
df['Un. Medida_Numero'] = df['Un. Medida_Numero'].astype(float)

# Criando uma condição para identificar linhas que atendem a determinados critérios (número maior que 999 e medida em gramas ou mililitros).

condicao = (df['Un. Medida_Numero'] > 999) & ((df['Un. Medida_Medida'] == 'g') | (df['Un. Medida_Medida'] == 'ml'))

# Aplicando uma função lambda às linhas que atendem à condição. Essa função converte as medidas de gramas para quilogramas ou mililitros para litros, conforme apropriado.

df.loc[condicao, 'Un. Medida'] = df.loc[condicao].apply(lambda row: f'{row["Un. Medida_Numero"]/1000} kg' if row["Un. Medida_Medida"] == 'g' else f'{row["Un. Medida_Numero"]/1000} litro(s)', axis=1)

# Removendo as colunas auxiliares 'Un. Medida_Numero' e 'Un. Medida_Medida'.

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

Criando a linha com o Total da soma da coluna [Valor Total].

```bash

# Exibindo o DataFrame atualizado
# Salvando a ordem original das colunas do DataFrame.

original_ordem_colunas = df.columns.tolist()

# Criando um novo DataFrame contendo uma linha total com a soma dos valores da coluna 'Valor Total'.
total_linha = pd.DataFrame({'Produtos': 'Total', 'Un.': '', 'Valor Unitário': '', 'Unidades': '',
                          'Valor Total': df['Valor Total'].sum()}, index=[len(df)])

# Preenche todas as colunas, exceto 'Produtos' e 'Valor Total', da linha total com valores vazios.
total_linha.loc[:, df.columns.difference(['Produtos', 'Valor Total'])] = ''

# concatena o DataFrame original com a linha total.
df = pd.concat([df, total_linha])

# reorganiza as colunas do DataFrame df para a ordem original salva anteriormente.
df = df[original_ordem_colunas]
```
  ---

## Tabela transformada.

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

Agora é o momento de transformar o Data Frame em um orçamento formal, de forma que a logo fique no canto superior esquerdo e dados de vendedor e cliente em cima da tabela, usaremos a biblioteca ReportLab. 

Esta biblioteca será responsável pela geração e personalização de documentos PDF. 

```bash
# Definindo o nome do cliente como 'Priscila Motta'
nome_cliente = 'Priscila Motta'

# Definindo o CNPJ do cliente como '11111111/0001-11'
cnpj_cliente = '11111111/0001-11'

# Definindo o endereço do cliente como 'Presidente Prudente'
endereco_cliente = 'Presidente Prudente'

# Definindo o nome do vendedor como 'Pedro Martinez'
nome_vendedor = 'Pedro Martinez'

# Definindo o CNPJ da empresa como '11121112/0001-29'
cnpj_empresa = '11121112/0001-29'

# Definindo o endereço da empresa como 'Sorocaba'
endereco_empresa = 'Sorocaba'

# Definindo o caminho do logotipo como 'agro_logo.png'
caminho_logotipo = 'agro_logo.png'

# Definindo o nome do arquivo PDF como 'tabela_orcamento.pdf'
nome_arquivo_pdf = 'tabela_orcamento.pdf'

# Criando um documento PDF
doc = SimpleDocTemplate(nome_arquivo_pdf, pagesize=letter)
story = []

# Criando a tabela de cabeçalho
tabela_cabecalho = Table([
    [Image(caminho_logotipo, width=70, height=55, hAlign='LEFT'),
     f'Vendedor: Jones Carlos Manoel\nCNPJ Empresa: 11121112/0001-29\nEndereço Empresa: Sorocaba',
     f'Cliente: Itaburi Agro\nCNPJ: 11111111/0001-11\nEndereço: Presidente Prudente']
])

# Aplicando estilo à tabela de cabeçalho
style_cabecalho = TableStyle([
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
])

# Adicionando a tabela de cabeçalho à história
tabela_cabecalho.setStyle(style_cabecalho)
story.append(tabela_cabecalho)
story.append(Spacer(1, 12))

# Convertendo dados do DataFrame para uma lista e criando a tabela de preços
data_precos = [df.columns.tolist()] + df.values.tolist()
table_precos = Table(data_precos)

# Aplicando estilo à tabela de preços
style_precos = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
])

# Adicionando a tabela de preços à história
table_precos.setStyle(style_precos)
story.append(table_precos)

# Construindo o documento PDF
doc.build(story)

# Imprimindo mensagem de sucesso
print(f'Tabela exportada para {nome_arquivo_pdf} com sucesso!')

```


Abaixo, irei resumir todas as ações acima em funções. Em seguida, criarei uma aplicação pela biblioteca Tkinter.

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
    
    df['Valor Unitário Ajustado'] = df['Valor Unitário'] * (1 + df['Margem Lucro'])
    
    df['Valor Total Ajustado'] = df['Quantidade'] * df['Valor Unitário Ajustado']
    
    df = df.drop(['Valor Unitário', 'Valor Total', 'Margem Lucro'], axis=1)
    
    df = df.rename(columns={'Valor Unitário Ajustado': 'Valor Unitário', 'Valor Total Ajustado': 'Valor Total'})
    
    nova_ordem_colunas = ['Produtos', 'Quantidade', 'Valor Unitário', 'Valor Total','Un. Medida']
    
    df = df[nova_ordem_colunas]
    
    df['Valor Unitário'] = df['Valor Unitário'].astype(float).round(2)
    df['Valor Total'] = df['Valor Total'].astype(float).round(2)
    
    df[['Un. Medida_Numero', 'Un. Medida_Medida']] = df['Un. Medida'].str.extract(r'(\d+\.?\d*)\s?(\D+)')
    
    df['Un. Medida_Numero'] = df['Un. Medida_Numero'].astype(float)
    
    condicao = (df['Un. Medida_Numero'] > 999) & ((df['Un. Medida_Medida'] == 'g') | (df['Un. Medida_Medida'] == 'ml'))
    
    df.loc[condicao, 'Un. Medida'] = df.loc[condicao].apply(lambda row: f'{row["Un. Medida_Numero"]/1000} kg' if row["Un. Medida_Medida"] == 'g' else f'{row["Un. Medida_Numero"]/1000} litro(s)', axis=1)
    
    df.drop(['Un. Medida_Numero', 'Un. Medida_Medida'], axis=1, inplace=True)
    
    original_ordem_colunas = df.columns.tolist()
    
    total_linha = pd.DataFrame({'Produtos': 'Total', 'Un.': '', 'Valor Unitário': '', 'Unidades': '',
                              'Valor Total': df['Valor Total'].sum()}, index=[len(df)])
    
    total_linha.loc[:, df.columns.difference(['Produtos', 'Valor Total'])] = ''
    
    df = pd.concat([df, total_linha])
    
    df = df[original_ordem_colunas]
    
    return df 

# PDF ================================================================

def inicializar_tabela_cabecalho(vendedor_escolhido, cliente_escolhido):
    global tabela_cabecalho
    tabela_cabecalho = Table([
        [Image(caminho_logotipo, width=70, height=55, hAlign='LEFT'),
         f'Vendedor: {vendedor_escolhido}\nCNPJ Empresa: 11121112/0001-29\nEndereço Empresa: Sorocaba',
         f'Cliente: {cliente_escolhido}\nCNPJ: 11111111/0001-11\nEndereço: {get_endereco(cliente_escolhido)}']
    ])

def processar_e_gerar_pdf(df, vendedor_escolhido, cliente_escolhido):
    try:
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

#| =============================================================

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

Até o momento, o código está funcional, mas sua aplicação é restrita ao meu ambiente de compilação, para o vendedor não precisar ter aceso direto ao código, criarei uma interface gráfica simples usando a biblioteca Tkinter.

Quando abrir o executável, vai aparecer para o vendedor uma janela. Nela, ele pode se escolher como ponto focal (seu nome e dados) e os de um cliente de um menu.

Depois, ele irá clicar no botão "Processar". Isso vai permitir que o vendedor escolha um arquivo do Excel com as informações do fornecedor.

Nesse meio tempo a conversão será feita, e uma vez terminado, outra janela vai aparecer para que o vendedor escolha onde quer salvar o resultado.

```bash

# Criando a janela principal da interface gráfica
janela = tk.Tk()

# Definindo o título da janela
janela.title("Processador de Planilha")

# Criando um frame para conter os widgets (botões e menus)
frame_botoes = tk.Frame(janela)

# Definindo uma lista de vendedores disponíveis
vendedores = [" ", "Jones Karlos Manoel", "Adriana Motta", "Viviane Fitipaldi"]

# Definindo uma lista de clientes disponíveis, com informações de nome, CNPJ e endereço
clientes = [
    {"nome": " ", "cnpj": " ", "endereco": " "},
    {"nome": "Ytaburi Agro", "cnpj": "11111111/0001-11", "endereco": "Presidente Prudente"},
    {"nome": "Fazenda Moinho de vento", "cnpj": "11111111/0001-12", "endereco": "Itu"},
    {"nome": "Boiadero do Dudu", "cnpj": "11111111/0001-13", "endereco": "Betim"}
]

# Criando uma variável para armazenar o vendedor selecionado e definindo seu valor inicial
vendedor_var = StringVar(janela)
vendedor_var.set(vendedores[0])

# Criando uma variável para armazenar o cliente selecionado e definindo seu valor inicial
cliente_var = StringVar(janela)
cliente_var.set(clientes[0]["nome"])

# Criando um rótulo para o menu de seleção de vendedor
rotulo_vendedor = tk.Label(frame_botoes, text="Vendedor(a):")

# Criando um rótulo para o menu de seleção de cliente
rotulo_cliente = tk.Label(frame_botoes, text="Cliente:")

# Criando um menu suspenso (Combobox) para seleção do vendedor
menu_vendedor = ttk.Combobox(frame_botoes, textvariable=vendedor_var, values=vendedores, state="readonly", width=20)

# Criando um menu suspenso (Combobox) para seleção do cliente
menu_cliente = ttk.Combobox(frame_botoes, textvariable=cliente_var, values=[cliente["nome"] for cliente in clientes], state="readonly", width=20)

# Criando um botão para processar a planilha
botao_processar = tk.Button(frame_botoes, text="Processando", command=processar_planilha)

# Empacotando o frame de botões na janela, com margens de 5 pixels
frame_botoes.pack(padx=5, pady=5)

# Empacotando o rótulo e o menu de seleção de vendedor, lado a lado
rotulo_vendedor.pack(side=tk.LEFT, padx=5)
menu_vendedor.pack(side=tk.LEFT, padx=5)

# Empacotando o rótulo e o menu de seleção de cliente, lado a lado
rotulo_cliente.pack(side=tk.LEFT, padx=5)
menu_cliente.pack(side=tk.LEFT, padx=5)

# Empacotando o botão de processamento à esquerda dos menus
botao_processar.pack(side=tk.LEFT, padx=5)

# Iniciando o loop principal da interface gráfica
janela.mainloop()
```

  ---

# Vídeo conclusão

O vídeo abaixo demonstra o funcionamento da solução, onde a transformação da planilha para o pdf de orçamento final é feito em poucos cliques.

# VIDEO

[![Assista ao vídeo](https://img.youtube.com/vi/rLJLyROZpHE/0.jpg)](https://www.youtube.com/embed/rLJLyROZpHE)

---

# Próximos passos:

* Em um cenário real, uma equipe comercial frequentemente não se restringe a criar cotações de apenas um fornecedor; eles alternam entre fornecedores ou, em alguns casos, consolidam um orçamento único proveniente de diversas fontes de fornecedor. Considerando a evolução deste projeto, planejo a inclusão da capacidade de transformar a planilha a partir de múltiplas fontes de fornecimento para gerar um orçamento final, seja ele uma combinação de vários fornecedores ou não.
  
* Criar uma calculadora de impostos (exemplo: PIS/COFINS) sobre os valores dos itens.

* Criar uma calculadora de frete com base na cidade de coleta e entrega, com critérios de precificação baseados em peso total e km de distância.

* Criar a possibilidade do usuário enviar o e-mail para o cliente diretamente pelo executável.
