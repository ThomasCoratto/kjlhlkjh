import pandas as pd


arquivo_resultado_final = r'C:\Users\moesios\Desktop\tratamento contorno python\Resultado_final.xlsx'
df = pd.read_excel(arquivo_resultado_final)

df['Validação_FORMS'] = df.apply(lambda row: 'INVÁLIDO' if pd.isna(row['ID_FORMULARIO']) else ('INVÁLIDO' if df[df['ID_DADOS'] == row['ID_DADOS']]['ID_DADOS'].count() != row['NÚMERO DE PASSAGEIROS'] else 'VÁLIDO'), axis=1)


def validate_logradouro(row, coluna_logradouro):
    if pd.isna(row[coluna_logradouro]):
        return 'INVÁLIDO'
    elif any(char.isdigit() for char in str(row[coluna_logradouro])):
        return 'Contém número'
    else:
        return 'Verificar via Ponto de referência'

df['Validação_LOGRADOURO_ORIGEM'] = df.apply(lambda row: validate_logradouro(row, 'LOGRADOURO_ORIGEM'), axis=1)
df['Validação_LOGRADOURO_DESTINO'] = df.apply(lambda row: validate_logradouro(row, 'LOGRADOURO_DESTINO'), axis=1)

def validar_municipio_origem(row):
    if pd.notna(row['MUNICÍPIO DE ORIGEM DA VIAGEM']) and row['MUNICÍPIO DE ORIGEM DA VIAGEM'] != 'Porto Alegre':
        return 'VÁLIDO'
    return validate_logradouro(row, 'LOGRADOURO_ORIGEM')

def validar_municipio_destino(row):
    if pd.notna(row['MUNICÍPIO DE DESTINO DA VIAGEM']) and row['MUNICÍPIO DE DESTINO DA VIAGEM'] != 'Porto Alegre':
        return 'VÁLIDO'
    return validate_logradouro(row, 'LOGRADOURO_DESTINO')

df['Validação_LOGRADOURO_ORIGEM'] = df.apply(validar_municipio_origem, axis=1)
df['Validação_LOGRADOURO_DESTINO'] = df.apply(validar_municipio_destino, axis=1)

def validar_estado_origem_destino(row):
    if pd.isna(row['ESTADO DE ORIGEM DA VIAGEM']):
        return 'INVÁLIDO'
    elif row['ESTADO DE ORIGEM DA VIAGEM'] != 'RIO GRANDE DO SUL (RS)' or row['MUNICÍPIO DE ORIGEM DA VIAGEM'] != 'Porto Alegre':
        return 'Requer Análise'
    else:
        return 'VÁLIDO'

df['Validação_ORIGEM'] = df.apply(validar_estado_origem_destino, axis=1)

def validar_estado_destino(row):
    if pd.isna(row['ESTADO DE DESTINO DA VIAGEM']):
        return 'INVÁLIDO'
    elif row['MUNICÍPIO DE DESTINO DA VIAGEM'] != 'Rio Grande do Sul (RS)':
        return 'Requer Análise'
    else:
        return 'VÁLIDO'

df['Validação_DESTINO'] = df.apply(validar_estado_destino, axis=1)

condicao_carga_peso_vazia = (df['TIPO DE VEÍCULO'].str.contains('CARGA', case=False)) & (df['QUAL O PESO DA CARGA? (EM TONELADAS)'].isna())
df.loc[condicao_carga_peso_vazia, 'QUAL O PESO DA CARGA? (EM TONELADAS)'] = 0

df.to_excel(arquivo_resultado_final, index=False)
