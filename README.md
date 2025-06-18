
# Generador Automático de Certificados

## Tabla de Contenidos

1. [Introducción](#introducción)
2. [Objetivo](#objetivo)
3. [Beneficios](#beneficios)
4. [Cómo funciona](#cómo-funciona)
5. [Requisitos](#requisitos)
6. [Instrucciones de uso](#instrucciones-de-uso)
7. [Archivos generados](#archivos-generados)
8. [Licencia](#licencia)

---

## Introducción

El **Generador Automático de Certificados** es una herramienta diseñada para automatizar el proceso de emisión de certificados de participación. Utilizando **Microsoft Excel** y **Microsoft Word**, el sistema reemplaza los datos de los participantes en una plantilla de Word, generando un archivo individual para cada uno, sin intervención manual. Esta herramienta está orientada a organizaciones y eventos que requieren emitir certificados a gran escala de forma eficiente y precisa.

## Objetivo

El principal objetivo de este proyecto es reducir el tiempo y el esfuerzo involucrado en la creación manual de certificados. Este sistema permite a las organizaciones crear certificados personalizados para cada participante de manera automatizada, mejorando la eficiencia y reduciendo los errores humanos.

## Beneficios

- **Eficiencia**: Ahorro significativo de tiempo al automatizar el proceso de creación de certificados.
- **Reducción de Errores**: Minimización de errores humanos al sustituir los datos automáticamente.
- **Escalabilidad**: Ideal para eventos de gran escala donde se necesiten generar muchos certificados.
- **Consistencia**: Garantiza que todos los certificados tengan el mismo formato y sean correctos.
  
## Cómo funciona

Este sistema funciona utilizando dos componentes clave:

1. **Hoja de Excel**: Contiene los datos de los participantes (nombre, DNI, nivel de curso, etc.).
2. **Plantilla de Word**: Un documento de Word que contiene marcadores de texto como `<Nombre>`, `<Doc>`, `<Nivel>`, etc., que serán reemplazados por los valores correspondientes de Excel.
3. **Macro en Excel**: La macro en VBA extrae los datos de Excel y reemplaza automáticamente los marcadores en la plantilla de Word, generando un documento de certificado para cada fila de datos.

## Requisitos

Para utilizar esta herramienta, necesitas:

- **Microsoft Excel** con soporte para macros VBA.
- **Microsoft Word**.
- Una plantilla de **Word** con los marcadores de texto correspondientes.
- Un archivo **Excel** con los datos de los participantes.

## Instrucciones de uso

### 1. Preparar los datos en Excel

- Crea una hoja de Excel con las siguientes columnas:
    - **Nombre**
    - **DNI**
    - **Nivel**
    - **Lengua**
    - **Fecha de Inicio**
    - **Fecha de Expiración**
    - **Código**
  
  Asegúrate de completar las filas con la información de cada participante.

### 2. Configurar la plantilla de Word

- Abre el archivo **data.docx** (plantilla de Word).
- Coloca los marcadores de texto en el lugar adecuado, como:
    - `<Nombre>`
    - `<Doc>`
    - `<Nivel>`
    - `<Lengua>`
    - `<Fecha1>`
    - `<Fecha2>`
    - `<Codigo>`
  
### 3. Ejecutar la macro

- Abre el archivo de Excel que contiene los datos.
- Ejecuta la macro **CombinarWord** desde el editor de VBA (presiona `Alt + F11` para abrir el editor de VBA).
- La macro reemplazará automáticamente los marcadores de texto en la plantilla de Word con los datos de Excel y generará un archivo de certificado para cada participante.

### 4. Archivos generados

Los certificados se guardarán en la carpeta seleccionada con el formato de nombre:
