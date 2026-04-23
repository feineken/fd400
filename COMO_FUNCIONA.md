# Generador FD400 — Banco de Chile Pagos Masivos

Dos scripts Python independientes que leen un Excel y generan el archivo `.txt` de nómina en formato FD400.

---

## Requisito

```
pip install openpyxl
```

---

## Uso

```bash
python gen_remuneraciones.py remuneraciones.xlsx
python gen_proveedores.py    proveedores.xlsx
```

Si el Excel está en la misma carpeta y se llama exactamente `remuneraciones.xlsx` / `proveedores.xlsx`, puedes omitir el nombre:

```bash
python gen_remuneraciones.py
python gen_proveedores.py
```

Cada script genera un archivo en la misma carpeta con nombre:
- `remuneraciones_YYYYMMDD.txt`
- `proveedores_YYYYMMDD.txt`

---

## Variables de empresa

Al inicio de cada script hay un bloque que **debes editar antes de correr**:

```python
RUT_EMPRESA   = "16746523"   # sin DV, sin puntos ni guiones
DV_EMPRESA    = "4"
COD_CONVENIO  = "046"        # codigo de tu contrato con el banco
NUM_NOMINA    = "00001"      # numero unico que tu le asignas a esta nomina
NOMBRE_NOMINA = "NOMINAMAYO" # nombre libre, max 25 caracteres
FECHA_PAGO    = "20250430"   # AAAAMMDD
```

El **numero de nomina** lo asigna la empresa. Puede ser un correlativo mensual o cualquier identificador que uses internamente.

---

## Estructura del Excel de remuneraciones

Fila 1 = encabezado (ignorada). Datos desde fila 2.

| Col | Campo | Ejemplo | Notas |
|-----|-------|---------|-------|
| A | RUT beneficiario | `16746523` | Solo digitos, sin puntos ni guion |
| B | DV beneficiario | `4` o `k` | Un caracter |
| C | Nombre completo | `DIEGO VARGAS` | Max 60 caracteres |
| D | Codigo banco | `001` | Ver tabla de bancos abajo |
| E | Numero de cuenta | `009-008-785564` | Tal como esta en el banco |
| F | Medio de pago | `07` | Ver tabla de medios abajo |
| G | Monto (pesos enteros) | `850000` | Sin puntos ni comas |
| H | Descripcion de pago | `REMMAYO` | Max 119 caracteres |

### Registros generados: tipo 1 (empresa) + tipo 2 por cada fila

---

## Estructura del Excel de proveedores

Fila 1 = encabezado. Datos desde fila 2.

| Col | Campo | Ejemplo | Notas |
|-----|-------|---------|-------|
| A | RUT proveedor | `12345678` | Solo digitos |
| B | DV proveedor | `9` | Un caracter |
| C | Nombre | `EMPRESA LTDA` | Max 60 caracteres |
| D | Codigo banco | `016` | Ver tabla de bancos abajo |
| E | Numero de cuenta | `12345678` | |
| F | Medio de pago | `07` | |
| G | Monto total (pesos) | `1900000` | |
| H | Descripcion pago | `PAGO SERVICIOS MAYO` | Max 119 caracteres |
| I | Email | `pago@empresa.cl` | Para notificacion |
| J | Glosa email | `Pago por servicios mayo` | Max 250 caracteres |
| K | Glosa factura 1 | `Factura 001 servicios TI` | Dejar vacio si no aplica |
| L | Monto factura 1 | `950000` | |
| M | Glosa factura 2 | `Factura 002 soporte` | Dejar vacio si no aplica |
| N | Monto factura 2 | `950000` | |
| O | Glosa factura 3 | | Vacio si no aplica |
| P | Monto factura 3 | | |
| Q | Glosa factura 4 | | Vacio si no aplica |
| R | Monto factura 4 | | |

### Registros generados por proveedor:
- **Tipo 2**: siempre
- **Tipo 3**: siempre (email con glosa)
- **Tipo 4**: solo si la columna K tiene contenido (detalle de facturas)

---

## Tablas de referencia

### Tipo de pago (Registro tipo 1)
| Codigo | Descripcion |
|--------|-------------|
| 0101 | Remuneraciones (sueldos) |
| 0201 | Proveedores |

### Medios de pago (columna F)
| Codigo | Descripcion |
|--------|-------------|
| 01 | Abono cuenta corriente Banco de Chile |
| 06 | Abono cuenta ahorro Banco de Chile |
| 07 | Abono cuenta corriente otros bancos |
| 08 | Abono cuenta ahorro otros bancos |
| 11 | Bancuenta Credichile |

### Bancos (columna D)
| Codigo | Banco |
|--------|-------|
| 001 | Banco de Chile |
| 009 | Banco Internacional |
| 012 | Banco Estado |
| 014 | Scotiabank |
| 016 | BCI |
| 028 | BICE |
| 037 | Santander |
| 039 | Itau / Corpbanca |
| 049 | Security |
| 051 | Falabella |
| 055 | Consorcio |
| 672 | Coopeuch |
| 729 | Prepago Los Heroes |
| 730 | Tenpo Prepago |

---

## Formato del archivo de salida

- Encoding: **UTF-8**
- Fin de linea: **CRLF** (Windows)
- Largo de cada linea: **exactamente 400 caracteres**
- Campos numericos: relleno con `0` a la izquierda
- Campos texto: relleno con espacios a la derecha

---

## Limitaciones actuales (v1)

- Maximo **4 facturas por proveedor** en el detalle tipo 4. Si un proveedor tiene mas de 4 facturas, hay que agregar logica para generar multiples registros tipo 4.
- Los datos de empresa (RUT, convenio, fecha) se editan **directamente en el script**. En una version futura podrian venir de un archivo de configuracion o de una hoja del Excel.
- No valida RUT ni digito verificador.
- No valida que los montos de las facturas (cols K-R) sumen el monto total (col G).
