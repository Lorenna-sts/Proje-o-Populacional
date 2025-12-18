
# Processamento de Projeções Populacionais (ETL) - Goiás

Este repositório contém ferramentas de ETL (Extract, Transform, Load) desenvolvidas em Python para processar dados de projeções populacionais do IBGE (ou fonte similar) para o Estado de Goiás, gerando arquivos compatíveis para importação no Banco de Dados Estatísticos (BDE).

O projeto lê planilhas originais em Excel, calcula agregados demográficos específicos (faixas etárias, totais por sexo, etc.), mapeia para códigos de variáveis (`VAR_COD`) pré-definidos e exporta arquivos CSV anuais formatados.

## 📂 Estrutura dos Arquivos

### 1. Scripts Python

* **`Script var completo.py`** (Recomendado): Versão mais robusta e completa. Processa:
* Totais Gerais (Códigos 939, 940, 941).
* Faixas etárias masculinas específicas (Códigos 942, 943, 944).
* Grandes grupos etários agregados e idosos (Códigos 979 a 983).


* **`Script Claude.py`**: Uma versão focada especificamente no cálculo e extração dos grupos etários agregados quinquenais (Códigos 979 a 983) e padronização de chaves de mesclagem.

### 2. Banco de Dados (SQL)

* **`ALTER TABLE tb_dados_Inclusão anos_BDE.txt`**: Script DDL para adequar a tabela de destino (`tb_dados`), adicionando colunas dinâmicas para os anos projetados (ex: `d_2041`), permitindo a inserção dos dados gerados.

## ⚙️ Pré-requisitos

* **Python 3.x**
* Bibliotecas Python necessárias:
```bash
pip install pandas numpy openpyxl

```


* **Arquivos de Entrada** (Devem estar no mesmo diretório ou configurados no script):
* `projecoes_2024.xlsx`: Dados brutos das projeções (Aba: "2) POP_GRUPO QUINQUENAL").
* `Variáveis Projeção.xlsx`: Tabela de-para contendo a relação entre descrição textual e `VAR_COD`.



## 🚀 Funcionalidades do Script Principal (`Script var completo.py`)

1. **Carregamento e Limpeza**: Lê arquivos Excel, remove caracteres especiais de cabeçalhos e padroniza nomes de colunas.
2. **Filtragem**: Seleciona apenas dados referentes à sigla **GO** (Goiás).
3. **Cálculo de Agregados**:
* Soma faixas etárias quinquenais para criar grandes grupos (ex: 0-14, 15-29, 30-64, 65+).
* Isola grupos específicos (ex: Mulheres 90+).
* Calcula totais por sexo (Homens, Mulheres, Ambos).


4. **Mapeamento (Merge)**: Cruza os dados processados com a planilha de variáveis usando uma chave composta (`GRUPO_PADRONIZADO + SEXO_PADRONIZADO`) para atribuir o `VAR_COD` correto.
5. **Exportação**: Gera um arquivo CSV para cada ano (2000 a 2070).

## 📝 Formato de Saída (CSV)

Os arquivos são gerados no diretório configurado (ex: `Projeções 2070`) seguindo o padrão `GO_{ANO}.csv`.

**Especificações do arquivo:**

* **Separador**: Ponto e vírgula (`;`)
* **Encoding**: Latin-1
* **Formato Numérico**: Padrão brasileiro (milhar com ponto), sem casas decimais (ex: `1.500`).

**Colunas Geradas:**
| Coluna | Descrição | Exemplo |
| :--- | :--- | :--- |
| `LOC_NOME` | Nome do Local (Fixo) | Estado de Goiás |
| `LOC_COD` | Código do Local (Fixo) | 1000 |
| `VAR_COD` | Código da Variável | 939 |
| `d_{ANO}` | Valor da População | 1.250.000 |

## 🛠️ Como Utilizar

1. **Configuração de Caminhos**:
Abra o script `.py` e ajuste a variável `OUTPUT_DIR` para o caminho desejado na sua máquina:
```python
OUTPUT_DIR = r"C:\Caminho\Para\Seus\Documentos\Output"

```


2. **Execução**:
Execute o script via terminal ou IDE:
```bash
python "Script var completo.py"

```


3. **Atualização do Banco de Dados**:
Antes de importar os CSVs, execute o comando SQL contido em `ALTER TABLE...txt` no seu gerenciador de banco de dados para garantir que as colunas dos anos (ex: `d_2041`) existam na tabela `tb_dados`.

## 🔍 Códigos de Variáveis Processados

O script garante o mapeamento dos seguintes códigos (sujeito à existência na planilha de variáveis):

* **939**: Total Geral (Ambos)
* **940**: Total Homens
* **941**: Total Mulheres
* **942-944**: Faixas etárias jovens (Homens)
* **979**: Mulheres 90 anos ou mais
* **980**: Ambos 0 a 14 anos
* **981**: Ambos 15 a 29 anos
* **982**: Ambos 30 a 64 anos
* **983**: Ambos 65 anos ou mais

## ⚠️ Notas Importantes

* **Validação de MergeKeys**: O script possui logs de debug detalhados (prints) para verificar se as chaves de texto criadas a partir do Excel de projeção batem com as chaves do Excel de variáveis. Verifique o console se algum código aparecer como "não mapeado".
* **Formatação de Números**: O script converte os números para string para aplicar a formatação visual brasileira (pontos como separadores de milhar) antes de salvar o CSV. Certifique-se de que o sistema de destino espera este formato (VARCHAR/String) e não numérico puro.
