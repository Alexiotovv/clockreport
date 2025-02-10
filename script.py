import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel
df = pd.read_excel('marcaciones.xlsx')

# Definir la tolerancia de 5 minutos
tolerancia = timedelta(minutes=5)

# Definir los horarios permitidos
horarios_permitidos = [
    datetime.strptime('05:00:00', '%H:%M:%S').time(),
    datetime.strptime('13:00:00', '%H:%M:%S').time(),
    datetime.strptime('21:00:00', '%H:%M:%S').time()
]

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
        fecha = row['Fecha']
        horas = row['Hora'].split(',')
        horas = [datetime.strptime(h.strip(), '%H:%M:%S').time() for h in horas]
        horas.sort()

        descripcion = []
        # tardanzas = []

        for hora in horas:
            if empleado_id not in ultima_accion or ultima_accion[empleado_id] == 'Salida':
                if empleado_id in ultima_hora:
                    tiempo_trabajado = datetime.combine(fecha, hora) - datetime.combine(fecha, ultima_hora[empleado_id])
                    if tiempo_trabajado > timedelta(hours=8):
                        descripcion.append(f'Entrada: {hora.strftime("%H:%M:%S")}')
                        # tardanzas.append(f'Entrada: {evaluar_tiempo(hora, "Entrada")}')
                        ultima_accion[empleado_id] = 'Entrada'
                        ultima_hora[empleado_id] = hora
                        continue
                descripcion.append(f'Entrada: {hora.strftime("%H:%M:%S")}')
                # tardanzas.append(f'Entrada: {evaluar_tiempo(hora, "Entrada")}')
                ultima_accion[empleado_id] = 'Entrada'
            else:
                descripcion.append(f'Salida: {hora.strftime("%H:%M:%S")}')
                # tardanzas.append(f'Salida: {evaluar_tiempo(hora, "Salida")}')
                ultima_accion[empleado_id] = 'Salida'
            
            ultima_hora[empleado_id] = hora

        # resultados.append({**row, 'Descripción': ', '.join(descripcion), 'Tardanzas': ', '.join(tardanzas)})
        resultados.append({**row, 'Descripción': ', '.join(descripcion)})

    return pd.DataFrame(resultados)

# Aplicar el procesamiento
df = procesar_marcaciones(df)
# df = encontrar_tardanza(df_resultado)

df['Dscto'] = df['Descripción'].apply(procesar_descripcion)

# Guardar el resultado en un nuevo archivo Excel
df.to_excel('marcaciones_procesadas.xlsx', index=False)

print("Procesamiento completado. El archivo 'marcaciones_procesadas.xlsx' ha sido generado.")

