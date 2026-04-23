# gen_remuneraciones.py
# Genera archivo FD400 para nomina de remuneraciones (Banco de Chile)
# Uso: python gen_remuneraciones.py [remuneraciones.xlsx]

import sys
import openpyxl
from datetime import datetime

# --- VARIABLES EMPRESA (editar antes de cada corrida) ---
RUT_EMPRESA   = "16746523"   # sin DV, sin puntos ni guiones
DV_EMPRESA    = "4"
COD_CONVENIO  = "046"
NUM_NOMINA    = "00001"
NOMBRE_NOMINA = "NOMINAMAYO"
FECHA_PAGO    = "20250430"   # AAAAMMDD
# --------------------------------------------------------


def pad_num(v, n):
    """Rellena con ceros a la izquierda hasta largo n."""
    return str(v).strip().zfill(n)


def pad_char(v, n):
    """Rellena con espacios a la derecha hasta largo n, trunca si excede."""
    s = str(v).strip() if v is not None else ""
    return s.ljust(n)[:n]


def fmt_monto(pesos):
    """Convierte monto en pesos a string de 13 digitos con 2 decimales implicitos."""
    cents = int(round(float(pesos) * 100))
    return str(cents).zfill(13)


def build_reg1(total):
    """Registro tipo 1: Empresa. Unico por archivo."""
    r  = "01"                          # pos 1-2:   tipo registro
    r += pad_num(RUT_EMPRESA, 9)       # pos 3-11:  RUT empresa
    r += pad_char(DV_EMPRESA, 1)       # pos 12:    DV empresa
    r += pad_num(COD_CONVENIO, 3)      # pos 13-15: codigo convenio
    r += pad_num(NUM_NOMINA, 5)        # pos 16-20: numero nomina
    r += pad_char(NOMBRE_NOMINA, 25)   # pos 21-45: nombre nomina
    r += "01"                          # pos 46-47: moneda = pesos
    r += FECHA_PAGO                    # pos 48-55: fecha pago AAAAMMDD
    r += fmt_monto(total)              # pos 56-68: monto total
    r += " " * 3                       # pos 69-71: filler
    r += "N"                           # pos 72:    tipo endoso nominativo
    r += " " * 322                     # pos 73-394: filler
    r += "01"                          # pos 395-396: tipo proceso = pago
    r += "0101"                        # pos 397-400: tipo pago = sueldos
    assert len(r) == 400, f"Reg1 largo incorrecto: {len(r)}"
    return r


def build_reg2(row, num_msg):
    """Registro tipo 2: Beneficiario. Uno por fila del Excel."""
    # Columnas Excel: A=RUT, B=DV, C=Nombre, D=Banco, E=Cuenta, F=MedioPago, G=Monto, H=Descripcion
    r  = "02"                          # pos 1-2:   tipo registro
    r += pad_num(RUT_EMPRESA, 9)       # pos 3-11:  RUT empresa
    r += pad_char(DV_EMPRESA, 1)       # pos 12:    DV empresa
    r += pad_num(COD_CONVENIO, 3)      # pos 13-15: codigo convenio
    r += "  "                          # pos 16-17: filler
    r += pad_num(NUM_NOMINA, 5)        # pos 18-22: numero nomina
    r += pad_num(row[5], 2)            # pos 23-24: medio de pago
    r += pad_num(row[0], 9)            # pos 25-33: RUT beneficiario
    r += pad_char(row[1], 1)           # pos 34:    DV beneficiario
    r += pad_char(row[2], 60)          # pos 35-94: nombre
    r += "0"                           # pos 95:    tipo direccion
    r += " " * 35                      # pos 96-130: direccion
    r += " " * 15                      # pos 131-145: comuna
    r += " " * 15                      # pos 146-160: ciudad
    r += " " * 7                       # pos 161-167: filler
    r += "  "                          # pos 168-169: actividad economica (solo vale vista)
    r += pad_num(row[3], 3)            # pos 170-172: codigo banco
    r += pad_char(row[4], 22)          # pos 173-194: numero cuenta (izquierda)
    r += "000"                         # pos 195-197: oficina destino
    r += fmt_monto(row[6])             # pos 198-210: monto pago
    r += pad_char(row[7], 119)         # pos 211-329: descripcion pago
    r += pad_num(num_msg, 4)           # pos 330-333: numero mensaje
    r += "N"                           # pos 334:    vale vista acumulado
    r += "   "                         # pos 335-337: tipo documento
    r += "0000000000"                  # pos 338-347: numero documento
    r += "+"                           # pos 348:    signo (pago a favor)
    r += "000000"                      # pos 349-354: correlativo vale vista
    r += "N"                           # pos 355:    vale vista virtual
    r += " " * 45                      # pos 356-400: filler
    assert len(r) == 400, f"Reg2 fila {num_msg} largo incorrecto: {len(r)}"
    return r


def main():
    archivo = sys.argv[1] if len(sys.argv) > 1 else "remuneraciones.xlsx"

    wb = openpyxl.load_workbook(archivo)
    ws = wb["Remuneraciones"]

    rows = []
    total = 0.0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        rows.append(row)
        total += float(row[6])

    if not rows:
        print("Error: no se encontraron filas con datos.")
        sys.exit(1)

    lines = [build_reg1(total)]
    for i, row in enumerate(rows, 1):
        lines.append(build_reg2(row, i))

    fecha = datetime.now().strftime("%Y%m%d")
    salida = f"remuneraciones_{fecha}.txt"
    with open(salida, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")

    print(f"OK: {salida}")
    print(f"    Beneficiarios : {len(rows)}")
    print(f"    Monto total   : ${total:,.0f}")


if __name__ == "__main__":
    main()
