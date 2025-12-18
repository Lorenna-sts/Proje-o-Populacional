import pandas as pd
import numpy as np
import os
import re

# =============================================================================
# 1. CONFIGURAÇÕES
# =============================================================================

PROJECOES_FILE = "projecoes_2024.xlsx"
VARIAVEIS_FILE = "Variáveis Projeção.xlsx"
PROJECOES_SHEET = "2) POP_GRUPO QUINQUENAL"
VARIAVEIS_SHEET = "Planilha1"
OUTPUT_DIR = r"C:\Users\lorenna.santos\OneDrive - Subsecretaria de Tecnologia da Informação\Documentos\Projeções 2070"
ANOS = [str(ano) for ano in range(2000, 2071)]

# =============================================================================
# 2. CARREGAMENTO DE DADOS
# =============================================================================

def load_excel(file_path, sheet_name, skiprows=None):
    """Carrega uma planilha Excel e converte nomes de colunas para string."""
    try:
        df = pd.read_excel(
            file_path, 
            sheet_name=sheet_name, 
            skiprows=skiprows, 
            engine='openpyxl'
        )
        df.columns = [str(col).strip().replace('.', '').replace('CÓD', 'COD') for col in df.columns]
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"Erro: Arquivo '{file_path}' não encontrado. Verifique se os arquivos de entrada estão na mesma pasta do script.")
    except ValueError as e:
        if "Worksheet named" in str(e):
            raise ValueError(f"Erro: Planilha '{sheet_name}' não encontrada no arquivo '{file_path}'. Verifique o nome da aba.")
        raise

print(f"Carregando {PROJECOES_FILE}, aba '{PROJECOES_SHEET}'...")
df_projecoes = load_excel(PROJECOES_FILE, PROJECOES_SHEET, skiprows=5)
df_projecoes.columns = [c.replace('GRUPO ETÁRIO', 'GRUPO_ETARIO').replace('GRUPO ETARIO', 'GRUPO_ETARIO') for c in df_projecoes.columns]

print(f"Carregando {VARIAVEIS_FILE}, aba '{VARIAVEIS_SHEET}'...")
df_variaveis = load_excel(VARIAVEIS_FILE, VARIAVEIS_SHEET, skiprows=None)
df_variaveis['VAR'] = df_variaveis['VAR'].astype(str).str.strip()

# =============================================================================
# 3. FILTRO, CÁLCULO E MESCLAGEM
# =============================================================================

print("Filtrando projeções para SIGLA = 'GO'...")
df_go = df_projecoes[df_projecoes['SIGLA'] == 'GO'].copy()

print("Extraindo e calculando grupos etários agregados (979 a 983)...")

df_go_ambos = df_go[df_go['SEXO'] == 'Ambos'].copy()

print("\n" + "="*80)
print("DEBUG: Investigação dos Grupos Etários")
print("="*80)

print(f"\nLinhas com SEXO='Ambos': {len(df_go_ambos)}")
print("\nGrupos etários únicos em df_go_ambos (ANTES de qualquer processamento):")
grupos_unicos = sorted(df_go_ambos['GRUPO_ETARIO'].unique())
for i, grupo in enumerate(grupos_unicos, 1):
    count = len(df_go_ambos[df_go_ambos['GRUPO_ETARIO'] == grupo])
    print(f"  {i}. '{grupo}' - {count} linhas")

# Padronizar grupos etários ANTES de fazer agregações
print("\n\nPadronizando grupos etários...")
df_go_ambos['GRUPO_ETARIO_PADRAO'] = (
    df_go_ambos['GRUPO_ETARIO'].astype(str).str.strip()
    .str.replace('00-04', '0-4', regex=False)
)

print("Grupos após padronização:")
grupos_padrao = sorted(df_go_ambos['GRUPO_ETARIO_PADRAO'].unique())
for i, grupo in enumerate(grupos_padrao, 1):
    count = len(df_go_ambos[df_go_ambos['GRUPO_ETARIO_PADRAO'] == grupo])
    print(f"  {i}. '{grupo}' - {count} linhas")

# Definir mapeamento de agregados com base nos grupos reais
AGREGADOS = {
    '0-14': ['0-4', '5-9', '10-14'],
    '15-29': ['15-19', '20-24', '25-29'],
    '30-64': ['30-34', '35-39', '40-44', '45-49', '50-54', '55-59', '60-64'],
    '65 ou mais': ['65-69', '70-74', '75-79', '80-84', '85-89', '90 ou mais']
}

new_rows = []

print(f"\n\nCriando agregados:")
for grupo_agregado, grupos_quinquenais in AGREGADOS.items():
    df_sum = df_go_ambos[df_go_ambos['GRUPO_ETARIO_PADRAO'].isin(grupos_quinquenais)]
    
    print(f"\n  Agregado '{grupo_agregado}':")
    print(f"    Procurando por: {grupos_quinquenais}")
    print(f"    Linhas encontradas: {len(df_sum)}")
    
    if not df_sum.empty:
        soma_anual = df_sum[ANOS].sum()
        
        new_row = {
            'GRUPO_ETARIO': grupo_agregado,
            'SEXO': 'Ambos',
            'SIGLA': 'GO',
            **soma_anual.to_dict()
        }
        new_rows.append(new_row)
        print(f"    ✓ Agregado criado com sucesso")
    else:
        print(f"    ✗ Nenhuma linha encontrada para este agregado")
        print(f"    Grupos disponíveis: {grupos_padrao}")

# EXTRAÇÃO DO GRUPO 979 (Sexo = 'Mulheres', Grupo = '90 ou mais')
print(f"\n\n  Código 979 - Mulheres 90 ou mais:")
df_979 = df_go[(df_go['SEXO'] == 'Mulheres') & (df_go['GRUPO_ETARIO'].str.strip() == '90 ou mais')].copy()
print(f"    Linhas encontradas: {len(df_979)}")

if not df_979.empty:
    for _, row in df_979.iterrows():
        new_rows.append(row.to_dict())
    print(f"    ✓ Código 979 adicionado com sucesso")
else:
    print(f"    ✗ Nenhuma linha encontrada")

# Concatenar os agregados ao DataFrame principal
print(f"\n\nConcatenando agregados:")
print(f"  df_go original: {len(df_go)} linhas")

df_agregados = pd.DataFrame(new_rows, columns=df_go.columns)
print(f"  Agregados a adicionar: {len(df_agregados)} linhas")

df_go = pd.concat([df_go, df_agregados], ignore_index=True)
print(f"  df_go após concatenação: {len(df_go)} linhas")

# Remover coluna auxiliar de padronização
if 'GRUPO_ETARIO_PADRAO' in df_go.columns:
    df_go = df_go.drop('GRUPO_ETARIO_PADRAO', axis=1)

# =============================================================================
# C. CRIAÇÃO DAS CHAVES DE MESCLAGEM
# =============================================================================

# 1. Padronização de GRUPO ETÁRIO em df_go
df_go['GRUPO_PADRONIZADO'] = (
    df_go['GRUPO_ETARIO'].astype(str).str.strip().str.lower()
    .str.replace(' ', '')
    .str.replace('00-', '0-', regex=False)
    .str.replace(r'(\d+)-(\d+)', r'\1-\2', regex=True) 
)
df_go['GRUPO_PADRONIZADO'] = df_go['GRUPO_PADRONIZADO'].str.replace('90oumais', '90+', regex=False)
df_go['GRUPO_PADRONIZADO'] = df_go['GRUPO_PADRONIZADO'].str.replace('65oumais', '65+', regex=False)

# 2. Padronização de SEXO em df_go
df_go['SEXO_PADRONIZADO'] = df_go['SEXO'].replace({
    'Ambos': 'total',
    'Homens': 'masculina',
    'Mulheres': 'feminina'
})

# 3. Criação da MergeKey em df_go
df_go['MergeKey'] = df_go['GRUPO_PADRONIZADO'] + '|' + df_go['SEXO_PADRONIZADO']

# 4. Criação da MergeKey em df_variaveis
def extract_group_sex_variaveis(var_str):
    """Extrai grupo e sexo de VAR e padroniza para a chave de mesclagem."""
    
    if 'Feminina' in var_str:
        sexo = 'feminina'
    elif 'Masculina' in var_str:
        sexo = 'masculina'
    else:
        sexo = 'total'

    var_lower = var_str.lower()
    
    # PADRONIZAÇÃO DOS GRUPOS AGREGADOS (980-983 e 979)
    if '0 a 14 anos' in var_lower:
        grupo = '0-14'
    elif '15 a 29 anos' in var_lower:
        grupo = '15-29'
    elif '30 a 64 anos' in var_lower:
        grupo = '30-64'
    elif '65 anos ou mais' in var_lower:
        grupo = '65+'
    elif '90 anos ou mais' in var_lower or '90+' in var_lower:
        grupo = '90+'
    elif 'total' in var_lower:
        grupo = 'total'
    else:
        # Quinquenais
        match_idade = re.search(r'(\d{1,2})\s?a\s?(\d{1,2})\sanos', var_lower)
        if match_idade:
            grupo = f"{match_idade.group(1)}-{match_idade.group(2)}"
        else:
            grupo = 'desconhecido'
        
    return grupo, sexo

df_variaveis[['GRUPO_VAR', 'SEXO_VAR']] = df_variaveis['VAR'].apply(
    lambda x: pd.Series(extract_group_sex_variaveis(x))
)

df_variaveis['MergeKey'] = df_variaveis['GRUPO_VAR'] + '|' + df_variaveis['SEXO_VAR']

# =============================================================================
# DEBUG: Verificar MergeKeys
# =============================================================================

print("\n" + "="*80)
print("DEBUG: Verificação de MergeKeys")
print("="*80)

print("\nMergeKeys em df_variaveis para códigos 980-983:")
df_980_983 = df_variaveis[df_variaveis['VAR_COD'].isin([980, 981, 982, 983])]
print(df_980_983[['VAR_COD', 'GRUPO_VAR', 'SEXO_VAR', 'MergeKey']].to_string())

print("\n\nComparação - Agregados vs Variáveis:")
merge_keys_go = df_go['MergeKey'].unique()
print("Agregados (df_go)      | Status em df_variaveis")
print("="*50)
for cod in [980, 981, 982, 983]:
    linha = df_variaveis[df_variaveis['VAR_COD'] == cod]
    if len(linha) > 0:
        key = linha.iloc[0]['MergeKey']
        existe = key in merge_keys_go
        status = "✓ MATCH ENCONTRADO" if existe else "✗ NÃO ENCONTRADA"
        print(f"{key:<25} | {status}")

# =============================================================================
# D. MESCLAGEM FINAL
# =============================================================================

df_final = pd.merge(
    df_go,
    df_variaveis[['VAR_COD', 'MergeKey']],
    on='MergeKey',
    how='left'
)

# E. LIMPEZA E VERIFICAÇÃO FINAL
print("\n" + "="*80)
print("--- Resultado do Mapeamento Final ---")
print("="*80)

df_final = df_final.dropna(subset=['VAR_COD'])
df_final['VAR_COD'] = df_final['VAR_COD'].astype(int)
print(f"\nTotal de linhas Mapeadas: {len(df_final)}")

# Verificação final dos códigos que estavam faltando
codigos_faltantes = [979, 980, 981, 982, 983]
print("\nVerificação dos códigos 979-983:")
for cod in codigos_faltantes:
    if cod in df_final['VAR_COD'].values:
        count = len(df_final[df_final['VAR_COD'] == cod])
        print(f"✓ Código {cod} foi mapeado com sucesso. ({count} linhas)")
    else:
        print(f"✗ Aviso: O código {cod} AINDA está faltando no resultado final.")

# =============================================================================
# 4. GERAÇÃO DOS ARQUIVOS CSV
# =============================================================================

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

print(f"\nGerando {len(ANOS)} arquivos CSV no diretório: {OUTPUT_DIR}...")

for ano in ANOS:
    if ano in df_final.columns:
        df_output = df_final[['VAR_COD', ano]].copy()
        df_output.rename(columns={ano: 'VALOR'}, inplace=True)
        
        file_name = f"GO_{ano}.csv"
        file_path = os.path.join(OUTPUT_DIR, file_name)
        
        df_output.to_csv(file_path, index=False)
        # 4. GERAÇÃO DOS ARQUIVOS CSV (AJUSTADO CONFORME SOLICITADO)
# =============================================================================

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

print(f"\nGerando {len(ANOS)} arquivos CSV no diretório: {OUTPUT_DIR}...")

for ano in ANOS:
    if ano in df_final.columns:
        # 1. Seleciona VAR_COD e a coluna do ano (que contém o valor da população)
        df_output = df_final[['VAR_COD', ano]].copy()
        
        # 2. Adiciona as colunas solicitadas LOC_NOME e LOC_COD
        df_output.insert(0, 'LOC_NOME', 'Estado de Goiás') 
        df_output.insert(1, 'LOC_COD', 1000)
        
        # 3. Renomeia a coluna do ano (e.g., '2000') para o formato solicitado (e.g., 'd_2000')
        col_d_ano = f"d_{ano}"
        df_output.rename(columns={ano: col_d_ano}, inplace=True)
        df_output[col_d_ano] = df_output[col_d_ano].apply(
            lambda x: f"{x:,.0f}".replace(",", "TEMP_SEP").replace(".", ",").replace("TEMP_SEP", ".")
        )
        # 4. A coluna 'VALOR' é o valor da população, mas a sua solicitação implicou que
        #    a coluna de população DEVERIA se chamar d_ANO. 
        #    Assumimos que você quer: LOC_NOME, LOC_COD, VAR_COD, e a coluna de valor dinâmico d_ANO
        #    Se você precisava de uma coluna VALOR separada, por favor, me informe.
        
        file_name = f"GO_{ano}.csv"
        file_path = os.path.join(OUTPUT_DIR, file_name)
        
        df_output.to_csv(file_path, index=False, sep=';', encoding='latin-1')
    
print("\nProcesso concluído com sucesso!")
    
print("\nProcesso concluído com sucesso!")