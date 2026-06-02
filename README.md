# Generador FD400 — Banco de Chile Pagos Masivos

Scripts Python que leen un Excel y generan el archivo `.txt` de nómina en formato **FD400** listo para subir al portal de pagos masivos del Banco de Chile.

Soporta dos tipos de nómina:
- **Remuneraciones** (sueldos) — registros tipo 1 y 2
- **Proveedores** — registros tipo 1, 2, 3 (email) y 4 (detalle facturas)

---

## Requisitos

- Python 3.8 o superior
- Instalar dependencias:

```bash
pip install -r requirements.txt
```

---

## Uso rápido

### 1. Editar variables de empresa

Al inicio de cada script hay un bloque que **debes editar antes de correr**:

```python
RUT_EMPRESA   = ""   # sin DV, sin puntos ni guiones
DV_EMPRESA    = ""
COD_CONVENIO  = ""   # codigo de tu contrato con el banco
NUM_NOMINA    = ""   # numero unico que asignas a esta nomina (max 5 digitos)
NOMBRE_NOMINA = ""   # nombre libre, max 25 caracteres
FECHA_PAGO    = ""   # AAAAMMDD
```

### 2. Preparar el Excel

Usa las plantillas incluidas como punto de partida:
- `remuneraciones_template.xlsx` → hoja `Remuneraciones`
- `proveedores_template.xlsx` → hoja `Proveedores`

### 3. Ejecutar

```bash
python gen_remuneraciones.py remuneraciones.xlsx
python gen_proveedores.py    proveedores.xlsx
```

Si el archivo se llama exactamente `remuneraciones.xlsx` / `proveedores.xlsx` y está en la misma carpeta, puedes omitir el nombre:

```bash
python gen_remuneraciones.py
python gen_proveedores.py
```

### 4. Resultado

Se genera un archivo en la misma carpeta:
- `remuneraciones_YYYYMMDD.txt`
- `proveedores_YYYYMMDD.txt`

---

## Estructura del Excel de remuneraciones

Fila 1 = encabezado (ignorada). Datos desde fila 2.

| Col | Campo | Ejemplo | Notas |
|-----|-------|---------|-------|
| A | RUT beneficiario | `16746523` | Solo dígitos, sin puntos ni guión |
| B | DV beneficiario | `4` o `k` | Un carácter |
| C | Nombre completo | `DIEGO VARGAS` | Máx. 60 caracteres |
| D | Código banco | `001` | Ver tabla de bancos |
| E | Número de cuenta | `009-008-785564` | Tal como aparece en el banco |
| F | Medio de pago | `07` | Ver tabla de medios de pago |
| G | Monto (pesos enteros) | `850000` | Sin puntos ni comas |
| H | Descripción de pago | `REMMAYO` | Máx. 119 caracteres |

---

## Estructura del Excel de proveedores

Fila 1 = encabezado. Datos desde fila 2.

| Col | Campo | Ejemplo | Notas |
|-----|-------|---------|-------|
| A | RUT proveedor | `12345678` | Solo dígitos |
| B | DV proveedor | `9` | Un carácter |
| C | Nombre | `EMPRESA LTDA` | Máx. 60 caracteres |
| D | Código banco | `016` | Ver tabla de bancos |
| E | Número de cuenta | `12345678` | |
| F | Medio de pago | `07` | |
| G | Monto total (pesos) | `1900000` | |
| H | Descripción pago | `PAGO SERVICIOS MAYO` | Máx. 119 caracteres |
| I | Email | `pago@empresa.cl` | Opcional — para notificación |
| J | Glosa email | `Pago por servicios mayo` | Máx. 250 caracteres |
| K | Glosa factura 1 | `Factura 001 servicios TI` | Dejar vacío si no aplica |
| L | Monto factura 1 | `950000` | |
| M | Glosa factura 2 | `Factura 002 soporte` | Dejar vacío si no aplica |
| N | Monto factura 2 | `950000` | |
| O | Glosa factura 3 | | Vacío si no aplica |
| P | Monto factura 3 | | |
| Q | Glosa factura 4 | | Vacío si no aplica |
| R | Monto factura 4 | | |

---

## Tablas de referencia

### Medios de pago (columna F)

| Código | Descripción |
|--------|-------------|
| 01 | Cta. Cte. Banco de Chile |
| 06 | Cta. Ahorro Banco de Chile |
| 07 | Cta. Cte. otros bancos |
| 08 | Cta. Ahorro otros bancos |
| 11 | Bancuenta Credichile |
| 12 | Pago efectivo Servipag |

### Bancos (columna D)

| Código | Banco |
|--------|-------|
| 001 | Banco de Chile |
| 009 | Banco Internacional |
| 012 | Banco Estado |
| 014 | Scotiabank |
| 016 | BCI |
| 028 | BICE |
| 031 | HSBC |
| 037 | Santander |
| 039 | Itaú / Corpbanca |
| 041 | JP Morgan |
| 049 | Security |
| 051 | Falabella |
| 053 | Banco Ripley |
| 055 | Consorcio |
| 672 | Coopeuch |
| 729 | Prepago Los Héroes |
| 730 | Tenpo Prepago |

---

## Formato del archivo de salida

| Propiedad | Valor |
|-----------|-------|
| Encoding | UTF-8 |
| Fin de línea | CRLF (Windows) |
| Largo por línea | exactamente 400 caracteres |
| Campos numéricos | relleno con `0` a la izquierda |
| Campos texto | relleno con espacios a la derecha |

---

## Limitaciones (v1)

- Máximo 4 facturas por proveedor en el detalle tipo 4.
- Los datos de empresa (RUT, convenio, fecha) se editan directamente en el script.
- No valida RUT ni dígito verificador.
- No valida que los montos de las facturas sumen el monto total.
