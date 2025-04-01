# Procesador de Facturas CSV

Este proyecto proporciona herramientas para procesar, analizar y gestionar archivos CSV de facturas empresariales de manera segura, evitando la exposición de datos sensibles al repositorio Git.

## Características

- Procesamiento de archivos CSV de facturación
- Análisis de datos financieros
- Generación de reportes y visualizaciones
- Implementación segura para proteger datos confidenciales
- Compatibilidad con formatos de facturación estándar

## Instalación

1. Clone el repositorio:
   ```bash
   git clone https://github.com/usuario/procesador-facturas-csv.git
   cd procesador-facturas-csv
   ```

2. Cree un entorno virtual:
   ```bash
   python -m venv venv
   source venv/bin/activate  # En Windows: venv\Scripts\activate
   ```

3. Instale las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

4. Configure el archivo `.gitignore` para proteger datos sensibles:
   ```bash
   # Ejemplo de contenido para .gitignore
   /Input/*.csv
   !/Input/ejemplo_facturacion.csv  # Mantiene el archivo de ejemplo
   ```

## Uso

1. Coloque sus archivos CSV de facturación en la carpeta `Input/`
2. Ejecute el procesador principal:
   ```bash
   python -m src.processor
   ```
3. Los resultados se guardarán en la carpeta `Output/`

## Personalización

El proyecto está diseñado para ser fácilmente personalizable:

### Modificar el formato de entrada

Para adaptar el código a diferentes formatos de CSV:

1. Abra `src/processor.py`
2. Localice la función `read_csv()` (aproximadamente línea 25)
3. Ajuste las columnas y el mapeo según su formato:
   ```python
   # Ejemplo de personalización
   def read_csv(filepath):
       # Modificar nombres de columnas aquí
       column_mapping = {
           'Su Columna': 'column_name_internal',
           'Otra Columna': 'another_column_internal',
           # Añada más mapeos según sea necesario
       }
       
       # Modificar tipos de datos aquí
       dtype_mapping = {
           'Columna Numérica': 'float64',
           'Columna Fecha': 'datetime64',
           # Añada más tipos según sea necesario
       }
       
       return pd.read_csv(
           filepath, 
           dtype=dtype_mapping,
           parse_dates=['Columna Fecha'],  # Ajuste según sea necesario
           # Otras opciones de configuración
       )
   ```

### Añadir nuevas funcionalidades

Para añadir nuevas capacidades de procesamiento:

1. Cree un nuevo módulo en la carpeta `src/`
2. Importe y utilice las funciones existentes
3. Actualice el procesador principal para incluir su nueva funcionalidad

Ejemplo:
```python
# En un nuevo archivo src/custom_processor.py

from .processor import read_csv
import pandas as pd

def my_custom_function(input_filepath, output_filepath):
    df = read_csv(input_filepath)
    
    # Su lógica personalizada aquí
    df['Nueva_Columna'] = df['Amount'] * 1.1  # Ejemplo
    
    # Guardar resultados
    df.to_csv(output_filepath, index=False)
    return df
```

### Configuración de seguridad

Para ajustar qué archivos se excluyen del control de versiones:

1. Edite el archivo `.gitignore` en la raíz del proyecto
2. Agregue o modifique patrones para proteger datos sensibles
3. Verifique que los archivos sensibles no se estén rastreando:
   ```bash
   git status
   ```

## Contribuciones

Las contribuciones son bienvenidas. Por favor, siga estos pasos:

1. Haga fork del repositorio
2. Cree una rama para su característica (`git checkout -b feature/nueva-caracteristica`)
3. Haga commit de sus cambios (`git commit -m 'Añadir nueva característica'`)
4. Haga push a la rama (`git push origin feature/nueva-caracteristica`)
5. Abra un Pull Request

## Open Source 

Este proyecto es open source y puede ser descargado y personalizado. 
