# Generador de Actas — Plan Concertado (GOR-F-084)

Genera el acta de plan concertado en Word, por colegio y por ficha, tomando
las competencias, RAP, guías, temas y actividades directamente del archivo
`RESULTADOS_A_EVALUAR_10_11_GUIA_TEMA_ACTIVIDAD.xlsx`.

## Contenido de la carpeta

| Archivo | Para qué sirve |
|---|---|
| `generar_plan_concertado.py` | El programa. Se ejecuta y abre una ventana. |
| `config_colegios.json` | **El único archivo que usted edita.** Colegios, fichas, voceros, horarios, ambiente, instructor. |
| `GOR-F-084FormatodeActaV02 PLAN CONCERTADO.docx` | Plantilla oficial. No la modifique. |
| `RESULTADOS_A_EVALUAR_10_11_GUIA_TEMA_ACTIVIDAD.xlsx` | Fuente de competencias, RAP, guías y actividades. |
| `Actas generadas/` | Aquí quedan los documentos producidos. |

## Instalación (una sola vez)

```bash
pip install python-docx openpyxl
```

Tkinter viene incluido con Python en Windows y macOS. En Linux:
`sudo apt install python3-tk`.

## Uso

```bash
python generar_plan_concertado.py
```

En la ventana:

1. Elija **colegio**, **grado** y **trimestre**. La ficha, el vocero, el día
   y las horas se cargan solos desde `config_colegios.json`; puede corregirlos
   antes de generar.
2. Escriba el **N° de acta** y ajuste la **fecha** de la reunión.
3. **Generar acta** produce un documento.
   **Generar todas las fichas** produce un acta por cada ficha registrada en la
   configuración, para el trimestre seleccionado.

El archivo queda en `Actas generadas/` con el nombre
`PLAN CONCERTADO T2 - COLEGIO ... - FICHA 3191165 - GRADO 11.docx`.

## Qué llena automáticamente

- **Nombre de la reunión**: trimestre en letras, programa, ficha y grado.
- **Ciudad y fecha, horas, lugar, ambiente y centro**.
- **Agenda**: fechas de inicio y fin del trimestre; las guías del trimestre
  quedan como viñetas redondas dentro del punto 3, una por guía.
- **Objetivo**: trimestre y nombre del colegio.
- **Desarrollo**: lista multinivel real de Word — `1.` para cada párrafo guía,
  `a.` para cada competencia (en negrita) y `i.` para cada RAP. Los temas van
  como `a.`, `b.` bajo el punto de ejes de formación. La numeración es de Word,
  no texto escrito a mano: si usted agrega o borra un ítem, se renumera solo.
- **Formato de oración**: los ejes de formación y las actividades de
  compromisos vienen del Excel EN MAYÚSCULAS SOSTENIDAS y se convierten a
  minúsculas con mayúscula inicial de cada oración (después de dos puntos va
  minúscula, como corresponde en español). Los términos técnicos conservan su
  grafía: `SQL`, `MySQL`, `PostgreSQL`, `MongoDB`, `ACID`, `MER`, `HTML5`… La
  lista está en `terminos_tecnicos` dentro de `config_colegios.json` y usted
  puede agregar los que le falten, con la clave en minúscula. Un texto que ya
  venga redactado en minúsculas no se toca (el umbral es 60% de mayúsculas).
- **Conclusiones**: fechas del trimestre, día y horario de sesión y el número de
  actividades; las guías bajan como sub-viñetas `o`, alineadas a la izquierda.
- **Compromisos**: una fila por actividad. La tabla crece o se reduce según
  cuántas actividades tenga el trimestre (no quedan filas sobrantes de la
  plantilla), y se renumeran desde 1. Las columnas ACTIVIDAD/DECISIÓN y
  RESPONSABLE van alineadas a la izquierda, y cada actividad se pasa a formato
  de oración igual que los ejes (ver más abajo).
- **Asistentes**: instructor y vocero en la columna NOMBRE; la fila de
  aprendices con el número de ficha.

Se conservan el encabezado, el logo, las fuentes y los bordes de la plantilla,
porque el programa escribe sobre el documento original en lugar de rehacerlo.

## Antes de usarlo en producción

1. **Complete los números de ficha en `config_colegios.json`.** Solo está
   cargada la ficha `3191165` (grado 11, Magdalena Ortega) tomada del acta de
   ejemplo; verifíquela y agregue las demás.
2. Revise días, horas y ambientes de cada colegio: están puestos como ejemplo.
3. En la hoja **11°** del Excel, el RAP `611187` aparece cortado
   (“…PERTINENTES A LAS”) y se copia así al acta. Si lo completa en el Excel,
   el acta sale bien sin tocar el programa.
4. En la hoja **10°**, trimestre 1, la celda de actividades tiene la numeración
   dañada (`…contexto escolar.52) identificar…`). El programa la separa y
   renumera correctamente, pero conviene corregirla en el origen.

## Cambiar el texto fijo del acta

Los párrafos de agenda, objetivo, desarrollo y conclusiones están en la función
`generar_acta()` del script, marcados con comentarios (`# --- Agenda ---`, etc.).
Edítelos ahí y aplican a todos los documentos.

Si en algún momento el SENA cambia la plantilla y se mueven las filas, ajuste
las constantes `R_*` del inicio del archivo, que indican en qué fila de la tabla
va cada bloque.

## Ajustar las viñetas y los tabuladores

Las sangrías están al inicio del script, en twips (1 cm = 567 twips). El primer
número es la sangría izquierda y el segundo la francesa:

```python
IND_NIVEL        = {0: (312, 236), 1: (672, 240), 2: (1032, 240)}  # desarrollo
IND_AGENDA       = {0: (312, 236), 1: (1134, 284)}                 # agenda
IND_CONCLUSIONES = {0: (312, 236), 1: (1134, 284)}                 # conclusiones
```

Los símbolos están en `FORMATOS_NIVEL`, `FORMATOS_AGENDA` y
`FORMATOS_CONCLUSIONES`. Cada nivel es `(tipo, símbolo, fuente)`:

| Tipo | Se ve como | Fuente |
|---|---|---|
| `("decimal", "%1.", None)` | 1. 2. 3. | — |
| `("lowerLetter", "%2.", None)` | a. b. c. | — |
| `("lowerRoman", "%3.", None)` | i. ii. iii. | — |
| `("bullet", "", "Symbol")` | • | Symbol |
| `("bullet", "o", "Courier New")` | o | Courier New |
| `("bullet", "-", "Calibri")` | - | Calibri |

El número dentro de `%1.` es el nivel, empezando en 1.

Qué texto va en cada nivel se decide en la sección `# --- Desarrollo de la
reunión ---` de `generar_acta()`: cada línea es una tupla
`(texto, nivel, negrita)`, donde el nivel `None` produce una línea en blanco
separadora que no recibe numeración.
