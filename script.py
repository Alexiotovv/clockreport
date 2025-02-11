import pandas as pd
from datetime import datetime, timedelta
import forms

# Cargar el archivo Excel
df = pd.read_excel('marcaciones.xlsx')

horarios=(forms.obtener_horarios())

# Definir la tolerancia de 5 minutos
horarios_permitidos = timedelta(minutes=5)

# Definir los horarios permitidos
# horarios_permitidos = [
#     datetime.strptime('05:00:00', '%H:%M:%S').time(),
#     datetime.strptime('13:00:00', '%H:%M:%S').time(),
#     datetime.strptime('21:00:00', '%H:%M:%S').time()
# ]

def procesar_descripcion(descripcion):
    # Dividir los eventos (entradas y salidas)
    eventos = descripcion.split(', ')
    
    # Inicializar una lista para almacenar los resultados de cada entrada
    resultados = []
    
    for evento in eventos:
        if evento.startswith('Entrada:'):
            # Extraer la hora de entrada
            entrada = evento.split('Entrada: ')[1].strip()
            
            # Dividir la hora, minutos y segundos
            try:
                hora_entrada, minutos_entrada, segundos_entrada = map(int, entrada.split(':'))
                
                # Aplicar la condición
                if (hora_entrada == 21 or hora_entrada == 13 or hora_entrada == 5) and (minutos_entrada >= 0 and segundos_entrada >= 0):
                    # Agregar el resultado de la tardanza
                    resultados.append(f"Dscto {minutos_entrada} min. {segundos_entrada} seg.")
                else:
                    # No hacer nada si cumple la condición
                    continue
            except ValueError:
                # Si no se puede convertir a int, ignorar este evento
                continue
    
    # Unir todos los resultados en una sola cadena separada por comas
    return ', '.join(resultados) if resultados else ""

# Función para evaluar puntualidad o tardanza
def evaluar_tiempo(hora, tipo):
    if tipo == 'Entrada':
        return 'Puntual' if any(hora <= h for h in horarios_permitidos) else 'Tardanza'
    else:
        return 'Puntual' if any(hora >= h for h in horarios_permitidos) else 'Antes'

# Procesar las marcaciones
def procesar_marcaciones(df):
    df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y')
    df = df.sort_values(by=['Id del Empleado', 'Fecha'])

    resultados = []
    ultima_accion = {}  # Para rastrear la última acción (entrada o salida) por empleado
    ultima_hora = {}    # Para rastrear la última hora de marcación

    for idx, row in df.iterrows():
        empleado_id = row['Id del Empleado']
        nombre = row['Nombres']  # Usar la columna "Nombres"
        fecha = row['Fecha']
        horas = row['Hora'].split(',')
        horas = [datetime.strptime(h.strip(), '%H:%M:%S').time() for h in horas]
        horas.sort()

        descripcion = []

        for hora in horas:
            if empleado_id not in ultima_accion or ultima_accion[empleado_id] == 'Salida':
                if empleado_id in ultima_hora:
                    tiempo_trabajado = datetime.combine(fecha, hora) - datetime.combine(fecha, ultima_hora[empleado_id])
                    if tiempo_trabajado > timedelta(hours=8):
                        descripcion.append(f'Entrada: {hora.strftime("%H:%M:%S")}')
                        ultima_accion[empleado_id] = 'Entrada'
                        ultima_hora[empleado_id] = hora
                        continue
                descripcion.append(f'Entrada: {hora.strftime("%H:%M:%S")}')
                ultima_accion[empleado_id] = 'Entrada'
            else:
                descripcion.append(f'Salida: {hora.strftime("%H:%M:%S")}')
                ultima_accion[empleado_id] = 'Salida'
            
            ultima_hora[empleado_id] = hora

        resultados.append({**row, 'Descripción': ', '.join(descripcion)})

    return pd.DataFrame(resultados)

# Calcular el total de minutos de tardanzas por trabajador
def calcular_tardanzas_totales(df):
    df['Minutos_Tardanza'] = df['Dscto'].apply(lambda x: sum(int(t.split()[1]) for t in x.split(', ') if 'Dscto' in t) if x else 0)
    return df.groupby(['Id del Empleado', 'Nombres'])['Minutos_Tardanza'].sum().reset_index()

# Calcular la cantidad de días trabajados por trabajador
def calcular_dias_trabajados(df):
    return df.groupby(['Id del Empleado', 'Nombres'])['Fecha'].nunique().reset_index()

# Crear un resumen de los trabajadores
def crear_resumen(df):
    # Calcular el total de minutos de tardanzas
    tardanzas_totales = calcular_tardanzas_totales(df)
    
    # Calcular la cantidad de días trabajados
    dias_trabajados = calcular_dias_trabajados(df)
    
    # Unir los resultados en un solo DataFrame
    resumen = pd.merge(tardanzas_totales, dias_trabajados, on=['Id del Empleado', 'Nombres'], how='outer')
    resumen.rename(columns={'Fecha': 'Dias_Trabajados'}, inplace=True)
    
    # Rellenar los valores nulos en "Minutos_Tardanza" con 0
    resumen['Minutos_Tardanza'] = resumen['Minutos_Tardanza'].fillna(0)
    
    return resumen

# Aplicar el procesamiento
df = procesar_marcaciones(df)

# Aplicar la función a la columna "Descripción" y guardar en "Dscto"
df['Dscto'] = df['Descripción'].apply(procesar_descripcion)

# Crear el resumen
resumen = crear_resumen(df)

# Guardar el resultado en un nuevo archivo Excel con dos hojas
with pd.ExcelWriter('marcaciones_procesadas.xlsx') as writer:
    df.to_excel(writer, sheet_name='Marcaciones', index=False)
    resumen.to_excel(writer, sheet_name='Resumen', index=False)

print("Procesamiento completado. El archivo 'marcaciones_procesadas.xlsx' ha sido generado.")
