import pandas as pd
import zipfile

# Leer archivo destino
df = pd.read_excel('enero_a_junio_unificado.xlsx', sheet_name='Sheet1')

# Leer históricos desde ZIP
historico = pd.DataFrame()
with zipfile.ZipFile('historicos.zip', 'r') as z:
    for archivo in z.namelist():
        if '.xlsx' in archivo:
            with z.open(archivo) as f:
                try:
                    temp = pd.read_excel(f, sheet_name='Historico')
                except:
                    temp = pd.read_excel(f, sheet_name=0)
                historico = pd.concat([historico, temp])

# Zonas
zonas = {'ABA': 'ABASOLO', 'CAL': 'CALVILLO', 'IRA': 'IRAPUATO', 
         'CEL': 'CELAYA', 'SIL': 'SILAO', 'AGS': 'AGS', 'APA': 'APASEO'
         , 'SN MIGUEL DE ALLENDE': 'SAN MIGUEL DE ALLENDE', 'SNJL': 'SAN JUAN','LEO': 'LEON'}
         

# Calcular
def calc(row):
    zona = zonas.get(row['Zona'], row['Zona'])
    entregas = historico[
        (historico['zona'] == zona) & 
        (historico['anio'] == int(row['AÑO_DETECTADO'])) & 
        (historico['mes'] == int(row['MES_DETECTADO']))
    ]['Entregas'].sum()
    return row['VENTA_MES'] - entregas

df['ventas_2'] = df.apply(calc, axis=1)

# Guardar
df.to_excel('resultado.xlsx', index=False)
print("✓ Listo! Revisa resultado.xlsx")
