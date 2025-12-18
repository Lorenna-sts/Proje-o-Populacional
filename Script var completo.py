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
        raise FileNotFoundError(f"Erro: Arquivo '{file_path}' não encontrado.")
    except ValueError as e:
        if "Worksheet named" in str(e):
            raise ValueError(f"Erro: Planilha '{sheet_name}' não encontrada.")
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

print("Extraindo e calculando grupos etários agregados...")

# Investigação dos grupos etários
df_go_ambos = df_go[df_go['SEXO'] == 'Ambos'].copy()

print("\n" + "="*80)
print("DEBUG: Grupos Etários Disponíveis")
print("="*80)

print(f"\nLinhas com SEXO='Ambos': {len(df_go_ambos)}")
print("Grupos etários únicos em df_go_ambos:")
grupos_unicos = sorted(df_go_ambos['GRUPO_ETARIO'].unique())
for grupo in grupos_unicos:
    count = len(df_go_ambos[df_go_ambos['GRUPO_ETARIO'] == grupo])
    print(f"  '{grupo}' - {count} linhas")

# Padronizar grupos etários
df_go['GRUPO_ETARIO_PADRAO'] = (
    df_go['GRUPO_ETARIO'].astype(str).str.strip()
    .str.replace('00-04', '0-4', regex=False)
    .str.replace('05-09', '5-9', regex=False)
    .str.replace('10-14', '10-14', regex=False)
)

# Definir mapeamento de agregados
AGREGADOS = {
    '0-14': ['0-4', '5-9', '10-14'],
    '15-29': ['15-19', '20-24', '25-29'],
    '30-64': ['30-34', '35-39', '40-44', '45-49', '50-54', '55-59', '60-64'],
    '65 ou mais': ['65-69', '70-74', '75-79', '80-84', '85-89', '90 ou mais']
}

new_rows = []

# =============================================================================
# CRIAR AGREGADOS: 980-983, 979, 939-944
# =============================================================================

print("\nCriando agregados...")

# CÓDIGOS 980-983: Agregados por faixa etária (Ambos)
for grupo_agregado, grupos_quinquenais in AGREGADOS.items():
    df_go_ambos = df_go[df_go['SEXO'] == 'Ambos'].copy()
    df_sum = df_go_ambos[df_go_ambos['GRUPO_ETARIO_PADRAO'].isin(grupos_quinquenais)]
    
    if not df_sum.empty:
        soma_anual = df_sum[ANOS].sum()
        new_row = {
            'GRUPO_ETARIO': grupo_agregado,
            'SEXO': 'Ambos',
            'SIGLA': 'GO',
            **soma_anual.to_dict()
        }
        new_rows.append(new_row)
        print(f"  ✓ Agregado '{grupo_agregado}' (Ambos) criado")

# CÓDIGO 979: Mulheres 90 ou mais
df_979 = df_go[(df_go['SEXO'] == 'Mulheres') & (df_go['GRUPO_ETARIO'].str.strip() == '90 ou mais')].copy()
if not df_979.empty:
    for _, row in df_979.iterrows():
        new_rows.append(row.to_dict())
    print(f"  ✓ Código 979 (Mulheres 90+) criado")

# CÓDIGO 939: Total Geral (soma de todos SEXO = Ambos)
df_ambos = df_go[df_go['SEXO'] == 'Ambos'].copy()
if not df_ambos.empty:
    soma_total = df_ambos[ANOS].sum()
    new_row = {
        'GRUPO_ETARIO': 'Total',
        'SEXO': 'Ambos',
        'SIGLA': 'GO',
        **soma_total.to_dict()
    }
    new_rows.append(new_row)
    print(f"  ✓ Código 939 (Total Geral) criado")

# CÓDIGO 940: Total Homens (soma de todas as faixas etárias, SEXO = Homens)
df_homens = df_go[df_go['SEXO'] == 'Homens'].copy()
if not df_homens.empty:
    soma_homens = df_homens[ANOS].sum()
    new_row = {
        'GRUPO_ETARIO': 'Total',
        'SEXO': 'Homens',
        'SIGLA': 'GO',
        **soma_homens.to_dict()
    }
    new_rows.append(new_row)
    print(f"  ✓ Código 940 (Total Homens) criado")

# CÓDIGO 941: Total Mulheres (soma de todas as faixas etárias, SEXO = Mulheres)
df_mulheres_all = df_go[df_go['SEXO'] == 'Mulheres'].copy()
if not df_mulheres_all.empty:
    soma_mulheres = df_mulheres_all[ANOS].sum()
    new_row = {
        'GRUPO_ETARIO': 'Total',
        'SEXO': 'Mulheres',
        'SIGLA': 'GO',
        **soma_mulheres.to_dict()
    }
    new_rows.append(new_row)
    print(f"  ✓ Código 941 (Total Mulheres) criado")

# CÓDIGOS 942-944: Faixas etárias específicas (Homens)
faixas_homens = [
    ('942', '0-4'),
    ('943', '5-9'),
    ('944', '10-14')
]

for codigo, faixa in faixas_homens:
    df_faixa = df_go[(df_go['SEXO'] == 'Homens') & (df_go['GRUPO_ETARIO_PADRAO'] == faixa)].copy()
    if not df_faixa.empty:
        # Se houver múltiplas linhas, somar
        if len(df_faixa) > 1:
            soma_faixa = df_faixa[ANOS].sum()
        else:
            soma_faixa = df_faixa.iloc[0][ANOS]
        
        new_row = {
            'GRUPO_ETARIO': faixa,
            'SEXO': 'Homens',
            'SIGLA': 'GO',
            **soma_faixa.to_dict()
        }
        new_rows.append(new_row)
        print(f"  ✓ Código {codigo} ({faixa} Homens) criado")

# Concatenar os agregados ao DataFrame principal
print(f"\nConcatenando agregados...")
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
    
    if '0 a 4 anos' in var_lower or '0-4' in var_lower or '0 a 4' in var_lower:
        grupo = '0-4'
    elif '5 a 9 anos' in var_lower or '5-9' in var_lower or '5 a 9' in var_lower:
        grupo = '5-9'
    elif '10 a 14 anos' in var_lower or '10-14' in var_lower or '10 a 14' in var_lower:
        grupo = '10-14'
    elif '0 a 14 anos' in var_lower:
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
        # Diferenciar entre diferentes tipos de total
        if 'mulheres' in var_lower and 'feminina' in var_lower:
            grupo = 'total'  # Para Mulheres - Total
            sexo = 'feminina'
        elif ('homens' in var_lower or 'masculina' in var_lower) and 'total' in var_lower:
            grupo = 'total'  # Para Homens - Total
            sexo = 'masculina'
        else:
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

print("\nMergeKeys em df_variaveis para códigos 939-944 e 979-983:")
df_especiais = df_variaveis[df_variaveis['VAR_COD'].isin([939, 940, 941, 942, 943, 944, 979, 980, 981, 982, 983])]
print(df_especiais[['VAR_COD', 'GRUPO_VAR', 'SEXO_VAR', 'MergeKey']].to_string())

print("\n\nComparação - df_go vs df_variaveis:")
merge_keys_go = df_go['MergeKey'].unique()
print("MergeKey (df_variaveis)    | Status em df_go")
print("="*50)
for cod in [939, 940, 941, 942, 943, 944, 979, 980, 981, 982, 983]:
    linha = df_variaveis[df_variaveis['VAR_COD'] == cod]
    if len(linha) > 0:
        key = linha.iloc[0]['MergeKey']
        existe = key in merge_keys_go
        status = "✓" if existe else "✗"
        print(f"Código {cod}: {key:<25} {status}")

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

# Verificação final dos códigos especiais
codigos_especiais = [939, 940, 941, 942, 943, 944, 979, 980, 981, 982, 983]
print("\nVerificação dos códigos especiais:")
for cod in codigos_especiais:
    if cod in df_final['VAR_COD'].values:
        count = len(df_final[df_final['VAR_COD'] == cod])
        print(f"✓ Código {cod} foi mapeado com sucesso. ({count} linhas)")
    else:
        print(f"✗ Aviso: O código {cod} AINDA está faltando no resultado final.")

# =============================================================================
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
        
        file_name = f"GO_{ano}.csv"
        file_path = os.path.join(OUTPUT_DIR, file_name)
        
        df_output.to_csv(file_path, index=False, sep=';', encoding='latin-1')
    
print("\nProcesso concluído com sucesso!")