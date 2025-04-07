import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

def leer_datos(archivo):
    print(f"\nLeyendo el archivo '{archivo}'")
    df = pd.read_excel(archivo)
    print("Nombres de las columnas:", df.columns.tolist())
    df['Fecha'] = pd.to_datetime(df['Fecha'])
    
    # Convertir tasas de fracción a porcentaje
    if df['Tasa de Desempleo  (%)'].max() < 1:
        print("Convirtiendo tasas a porcentaje...")
        df['Tasa de Desempleo  (%)'] = df['Tasa de Desempleo  (%)'] * 100
    # Ordenamos por fecha para asegurar la secuencia correcta
    df = df.sort_values(by='Fecha').reset_index(drop=True)
    return df

def generar_predicciones_con_ventana(df, dias_prediccion=12, ventana=5):
    print(f"\nGenerando predicciones con ventana de tiempo...")
    fecha_ultimo = df['Fecha'].max()
    print(f"Última fecha en los datos: {fecha_ultimo}")
    datos_completos = df['Tasa de Desempleo  (%)'].tolist()
    
    # Mostrar las últimas 5 fechas y sus valores
    print("\nÚltimas 5 fechas y sus valores de Tasa de Desempleo (%):")
    ultimas_fechas = df['Fecha'].tail(5).tolist()
    ultimos_valores = df['Tasa de Desempleo  (%)'].tail(5).tolist()
    for fecha, valor in zip(ultimas_fechas, ultimos_valores):
        print(f"Fecha: {fecha.strftime('%d/%m/%Y')}, Tasa de Desempleo: {valor:.6f}%")
    
    fechas_predichas = []
    predicciones = []
    
    print(f"\nGenerando predicciones para los próximos {dias_prediccion} días...")
    for i in range(1, dias_prediccion + 1):
        nueva_fecha = fecha_ultimo + timedelta(days=i)
        # Se toma el promedio de los últimos "ventana" dias
        ventana_valores = datos_completos[-ventana:]
        prediccion = np.mean(ventana_valores)
        
        predicciones.append(prediccion)
        fechas_predichas.append(nueva_fecha)
        datos_completos.append(prediccion)
    
    # Crear DataFrame para las predicciones
    print(f"\nPredicciones generadas:")
    for fecha, pred in zip(fechas_predichas, predicciones):
        print(f"Fecha: {fecha.strftime('%d/%m/%Y')}, Predicción: {pred:.6f}%")
        
    df_predicciones = pd.DataFrame({
        'Fecha': fechas_predichas,
        'Tasa_Predicha (%)': predicciones
    })
    return df_predicciones

def graficar_datos(df, df_pred):
    print(f"\nGenerando gráfico ...")
    plt.figure(figsize=(12, 7))
    
    # Datos históricos
    plt.plot(df['Fecha'], df['Tasa de Desempleo  (%)'], 
             label='Datos históricos', color='blue', marker='o', linestyle='-', linewidth=2)
    
    # Predicciones
    plt.plot(df_pred['Fecha'], df_pred['Tasa_Predicha (%)'], 
             label='Predicciones', color='red', marker='o', linestyle='-', linewidth=2)
    
    # Línea de transición
    ultimo_historico = df['Fecha'].max()
    primer_prediccion = df_pred['Fecha'].min()
    if ultimo_historico < primer_prediccion:
        valor_ultimo_historico = df.loc[df['Fecha'] == ultimo_historico, 'Tasa de Desempleo  (%)'].values[0]
        valor_primer_prediccion = df_pred.loc[df_pred['Fecha'] == primer_prediccion, 'Tasa_Predicha (%)'].values[0]

        plt.plot([ultimo_historico, primer_prediccion], 
                 [valor_ultimo_historico, valor_primer_prediccion], 
                 color='blue', linestyle='--', linewidth=2, label='Transición')
    
    plt.xlabel('Fecha', fontsize=12)
    plt.ylabel('Tasa de Desempleo (%)', fontsize=12)
    plt.title('Tasa de Desempleo Histórica y Predicciones', fontsize=14, fontweight='bold')
    plt.legend(fontsize=12)
    plt.grid(True, which='both', linestyle='--', linewidth=0.5, alpha=0.7)
    plt.xticks(rotation=45, fontsize=10)
    plt.yticks(fontsize=10)
    plt.tight_layout()
    plt.show()

def guardar_predicciones(df_pred, archivo_salida):
    print(f"\nGuardando predicciones en '{archivo_salida}'...")
    # Formatear las fechas para que solo incluyan el día, mes y año
    df_pred['Fecha'] = df_pred['Fecha'].dt.strftime('%d/%m/%Y')
    df_pred.to_excel(archivo_salida, index=False)
    print(f"Predicciones guardadas exitosamente en '{archivo_salida}'")
    
# Ejecución principal
if __name__ == "__main__":
    archivo_entrada = 'desempleo.xlsx'
    archivo_salida = 'predicciones_desempleo.xlsx'
    
    # Leer datos
    df_datos = leer_datos(archivo_entrada)  
    df_predicciones = generar_predicciones_con_ventana(df_datos, dias_prediccion=12, ventana=5)    
    graficar_datos(df_datos, df_predicciones)
    guardar_predicciones(df_predicciones, archivo_salida)