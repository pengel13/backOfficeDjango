def passarParaDecimal(row):
    hora: int = 0
    minuto: float = 0.0
    if len(row) == 4:
        hora = int(row[0])
        minuto: float = int(row[2:]) / 60
    elif len(row) == 5:
        hora = int(row[:2])
        minuto = int(row[3:]) / 60
    elif len(row) == 6:
        hora = int(row[:3])
        minuto = int(row[4:]) / 60
    return hora + minuto