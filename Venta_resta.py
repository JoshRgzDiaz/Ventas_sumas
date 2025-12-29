import pandas as pd
import zipfile
import sys

# Leer archivo
archivo = sys.argv[1] if len(sys.argv) > 1 else 'enero_a_junio_unificado.xlsx'
df = pd.read_excel(archivo)

# Eliminar columnas "_DIA"
columnas_dia = [col for col in df.columns if '_DIA' in col.upper() or col.upper().endswith('DIA')]
if columnas_dia:
    df = df.drop(columns=columnas_dia)

# Separar permiso_zona en Permiso y Zona
if 'permiso_zona' in df.columns:
    df['Permiso'] = df['permiso_zona'].str.split(' - ').str[0].str.strip()
    df['Zona'] = df['permiso_zona'].str.split(' - ').str[1].str.strip()
    # Eliminar columna original
    df = df.drop(columns=['permiso_zona'])

# Convertir columnas "_MES" a litros (÷ 0.54) - UNA SOLA VEZ
columnas_mes = [col for col in df.columns if '_MES' in col.upper() and col != 'MES_DETECTADO']
for col in columnas_mes:
    df[col] = df[col] / 0.54

# Leer históricos y calcular Ventas_2 (con valores YA en litros)
try:
    historico = pd.DataFrame()
    with zipfile.ZipFile('historicos.zip', 'r') as z:
        for archivo_zip in z.namelist():
            if '.xlsx' in archivo_zip:
                with z.open(archivo_zip) as f:
                    try:
                        temp = pd.read_excel(f, sheet_name='Historico')
                    except:
                        temp = pd.read_excel(f, sheet_name=0)
                    historico = pd.concat([historico, temp])
    
    # Zonas
    zonas = {
        'ABA': 'ABASOLO', 
        'CAL': 'CALVILLO', 
        'IRA': 'IRAPUATO', 
        'CEL': 'CELAYA', 
        'SIL': 'SILAO', 
        'AGS': 'AGS', 
        'APA': 'APASEO',
        'SN MIGUEL DE ALLENDE': 'SAN MIGUEL DE ALLENDE',
        'Sn Miguel de Allende': 'SAN MIGUEL DE ALLENDE',
        'SNJL': 'SAN JUAN',
        'SnJL': 'SAN JUAN',
        'LEO': 'LEON',
        'QRO': 'QRO'
    }
    
    # Calcular Ventas_2 = VENTA_MES (litros) - ENTREGAS (litros)
    def calc_ventas_2(row):
        zona = zonas.get(row['Zona'], row['Zona'])
        entregas = historico[
            (historico['zona'] == zona) & 
            (historico['anio'] == int(row['AÑO_DETECTADO'])) & 
            (historico['mes'] == int(row['MES_DETECTADO']))
        ]['Entregas'].sum()
        return row['VENTA_MES'] - entregas
    
    df['Ventas_2'] = df.apply(calc_ventas_2, axis=1)
    
except:
    pass

# Reordenar columnas en el orden correcto (sin EXISTENCIA_FINAL)
columnas_orden = ['Permiso', 'Zona', 'AÑO_DETECTADO', 'MES_DETECTADO', 
                  'EXIST. INI._MES', 'COMPRAS_MES', 'VENTA_MES', 'Ventas_2']

# Verificar qué columnas existen y ordenar
columnas_existentes = [col for col in columnas_orden if col in df.columns]
df = df[columnas_existentes]

# Guardar resultado
nombre_salida = archivo.replace('.xlsx', '_procesado.xlsx')
df.to_excel(nombre_salida, index=False)

print(f"✓ Listo: {nombre_salida}")