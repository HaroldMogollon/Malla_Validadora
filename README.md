# Malla de validación de Excel

Aplicación Streamlit para cargar un libro `.xlsx`, seleccionar la hoja a auditar, emparejar reglas con las columnas, revisar tipos de dato y descargar el log de observaciones en Excel o CSV.

## Estructura esperada de reglas

La app admite dos formatos de regla:

- Formato estructurado: `Campo de regla`, `Tipo_Campo`, `Obligatorio`, `No Aplica` y `Valores Aceptados`. `No Aplica` marcado con `X` habilita exactamente ese texto como excepción al tipo de dato. El valor `N/A` se rechaza.
- Formato narrativo: dos columnas con encabezados `COLUMNA` y `DESCRIPCIÓN` (también reconoce `CAMPO` y `REGLA`). Las filas posteriores sin nombre de campo se agregan como continuación de la descripción.

La app propone el emparejamiento con la hoja auditada. Revísalo antes de ejecutar, sobre todo si los nombres difieren entre las hojas. En el formato estructurado carga los tipos, obligatoriedad, indicador `No Aplica` y listas permitidas. Puede extraer un límite máximo de fecha de frases como `No debe ser mayor a DD/MM/AAAA`. En cualquiera de los formatos puedes revisar y ajustar las opciones antes de validar.

## Ejecutar localmente

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Desplegar en Streamlit Community Cloud

1. Sube `app.py`, `requirements.txt` y `README.md` a la raíz de un repositorio GitHub.
2. En Streamlit Community Cloud, crea una app desde ese repositorio y selecciona `app.py` como archivo principal.
3. Despliega y carga el libro desde la app.

## Contenido del log

El Excel descargado contiene `Resumen`, `Observaciones` (una fila por cada error de campo) y `Reglas aplicadas`. `Fila Excel` apunta al número real de fila del libro de origen.

## Consideraciones

- La malla aplica tipos, obligatoriedad, listas permitidas, excepción `No Aplica` y límite máximo de fecha cuando están definidos en el formato estructurado. Las reglas de unicidad se pueden activar en la interfaz.
- Las celdas vacías solo se marcan cuando se activa `Campo obligatorio`.
- En campos numéricos se admiten valores numéricos y números con separadores decimales comunes. En fecha se aceptan fechas nativas de Excel y formatos usuales, incluido `DD/MM/AAAA`.
- El texto exacto `No Aplica` solo se acepta en los campos que tengan `X` en la columna `No Aplica`. `N/A` se rechaza.
- El libro se procesa en memoria durante la sesión y no se escribe en el repositorio.
