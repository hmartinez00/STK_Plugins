# STK Plugins - Análisis de accesos y actualización orbital

Este repositorio contiene un conjunto de scripts auxiliares orientados a trabajar con AGI STK desde Python. La idea principal es facilitar tareas de preparación, actualización y análisis de escenarios espaciales sin tener que realizar todo manualmente dentro de la interfaz gráfica de STK.

## Propósito general

Los algoritmos incluidos aquí están pensados para apoyar workflows de simulación y análisis orbital en STK, especialmente en tareas relacionadas con:

- actualización de la órbita de un satélite a partir de efemérides externas,
- creación automática de targets y chains a partir de archivos CSV,
- análisis de accesos y selección del mejor momento de paso sobre un objetivo,
- generación de reportes exportables para su revisión posterior.

Aunque no están diseñados como un plugin compilado de STK en sí mismos, estos scripts funcionan como herramientas de automatización y extensión del flujo de trabajo de STK.

## Archivos principales

### 1. eph_updater.py
Este script permite leer un archivo de efemérides en formato EPH y actualizar la órbita del satélite dentro de un escenario abierto en STK.

#### Qué hace
- abre un archivo EPH seleccionado por el usuario,
- extrae parámetros orbitales como semieje mayor, excentricidad, inclinación, RAAN, argumento del perigeo y anomalía media,
- construye un intervalo temporal nuevo para el escenario,
- envía los comandos necesarios a STK para actualizar el estado orbital del satélite.

#### Uso típico
Se usa cuando se dispone de una nueva efeméride del satélite y se desea reflejar esa información en un escenario de STK sin hacerlo manualmente.

### 2. stk3_conn.py
Este script conecta Python con STK para crear targets y chains automáticamente a partir de un archivo CSV con coordenadas geográficas.

#### Qué hace
- solicita al usuario un archivo CSV con datos de targets,
- lee coordenadas de latitud y longitud,
- crea objetos Target en STK,
- crea Chains que relacionan un sensor del satélite con cada target,
- calcula los accesos asociados a esos chains.

#### Uso típico
Sirve para automatizar la generación de múltiples objetivos y su relación con un sensor, especialmente cuando se trabaja con listas largas de puntos geográficos.

### 3. add_ct.py
Este script procesa salidas de accesos y ángulos de STK para identificar el mejor instante de paso sobre un objetivo.

#### Qué hace
- lee archivos CSV que contienen bloques de datos de Access y Angles,
- identifica los intervalos de acceso de cada target,
- busca el punto dentro del acceso donde la distancia (range) es mínima,
- genera un resumen con el mejor tiempo y los valores angulares asociados,
- exporta el resultado a un archivo Excel.

#### Generación de Reporte de Ángulos Satelitales en STK

Sigue estos pasos generales para calcular y extraer los ángulos de acceso (Along Track y Cross Track) entre un satélite y un objetivo en STK:

1. **Iniciar el análisis de acceso**: Abre la herramienta **Access Tool** en el objeto satélite deseado y selecciona el objeto objetivo (*target*) de interés (por ejemplo, una instalación, región u otro satélite).
2. **Calcular acceso**: Haz clic en el botón **Compute** para calcular los intervalos de acceso entre los objetos.
3. **Abrir el gestor de reportes**: En la ventana de *Access*, navega a la sección **Reports** y haz clic en el botón **Report & Graph Manager...**.
4. **Editar el estilo de reporte**: En el *Report & Graph Manager*, selecciona el estilo de reporte **Access**, haz clic derecho sobre él y elige **Properties** (o **Edit**).
5. **Seleccionar el proveedor de datos**: En la ventana de *Data Providers*, busca y agrega el proveedor **Sat Angles Data** (o el proveedor de datos angulares que requieras).
6. **Configurar las variables**: Dentro del proveedor de datos seleccionado, marca los elementos específicos que deseas exportar, tales como:
   - **Time**
   - **Along Track**
   - **Cross Track**
   - **Range**
7. **Generar el reporte**: Haz clic en **OK** para guardar la configuración de los datos y luego en **Generate** para crear el reporte.
8. **Interpretar los resultados**: Revisa la tabla generada. Esta contendrá los parámetros seleccionados, donde el valor **Cross Track** sirve como aproximación del ángulo de *roll* del satélite respecto al objetivo.

#### Uso típico
Se utiliza cuando se quiere obtener una vista más compacta y útil de los accesos, destacando el momento más favorable de observación o cobertura.

## Flujo general de trabajo

Un flujo típico de uso podría ser:

1. actualizar la órbita del satélite con eph_updater.py,
2. crear targets y chains con stk3_conn.py,
3. ejecutar el análisis de accesos con add_ct.py,
4. revisar los resultados exportados para tomar decisiones de diseño o análisis.

## Requisitos

Las herramientas requieren Python y dependencias como:

- pandas
- comtypes
- tqdm
- openpyxl

Se recomienda instalar las dependencias desde el archivo requirements.txt del proyecto.

## Notas importantes

- Los scripts están orientados a un entorno Windows y a una integración específica con STK 11.
- La interacción con STK se realiza mediante comtypes y comandos de la API de STK.
- Algunos archivos y rutas están pensados para un escenario concreto, por lo que pueden requerir ajustes según el nombre del satélite, el sensor o la estructura del escenario.

## Objetivo del proyecto

El objetivo de este repositorio es simplificar y automatizar tareas repetitivas en STK, convirtiendo procesos manuales en scripts que puedan ejecutarse de forma más rápida, consistente y escalable.
