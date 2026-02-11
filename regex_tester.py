import re

# texto = "232 Comp L: 1,5,11-15 (2h) L:2-4,8-10,16(4h)"
# resultado = re.findall(r'(\d{2,3})\s+(([A-Za-z])+\s(([A-Z]:)\s*((\d*[,.-]*\d*)*)\s*(\(\d[hH]\))))*', texto)

def parse_asignatura(texto):
    # Expresión para capturar el código y nombre
    patron_inicial = re.compile(r'^(\d{2,3})\s+([A-Za-z]+)')
    
    # Expresión para cada sección de la segunda parte
    patron_seccion = re.compile(r'([A-Z]):\s*([\d\-,.]+)\s*\((\d+)h\)', re.IGNORECASE)
    
    # Buscar código y nombre
    match_inicial = patron_inicial.match(texto)
    if not match_inicial:
        return None
    
    codigo = match_inicial.group(1)
    nombre = match_inicial.group(2)
    resto = texto[match_inicial.end():].strip()
    
    # Buscar todas las secciones
    secciones = []
    for match in patron_seccion.finditer(resto):
        secciones.append({
            'tipo': match.group(1),
            'numeros': match.group(2),
            'horas': match.group(3)
        })
    
    # Parsear los números (incluyendo rangos y separadores mixtos)
    for seccion in secciones:
        numeros_str = seccion['numeros']
        # Reemplazar puntos por comas para estandarizar
        numeros_str = numeros_str.replace('.', ',')
        # Separar por comas
        partes = numeros_str.split(',')
        
        numeros_expandidos = []
        for parte in partes:
            parte = parte.strip()
            if '-' in parte:
                inicio, fin = map(int, parte.split('-'))
                numeros_expandidos.extend(range(inicio, fin + 1))
            else:
                if parte:
                    numeros_expandidos.append(int(parte))
        
        seccion['numeros_lista'] = numeros_expandidos
    
    return {
        'codigo': codigo,
        'nombre': nombre,
        'secciones': secciones
    }

texto = "232 Comp L: 1,5,11-15 (2h) S:2-4.8-10,16(4h)"
resultado = parse_asignatura(texto)

if resultado:
    print("Código:", resultado['codigo'])
    print("Nombre:", resultado['nombre'])
    print("\nSecciones:")
    for i, seccion in enumerate(resultado['secciones'], 1):
        print(f"  {seccion['tipo']}: {seccion['numeros']} ({seccion['horas']}h)")
        print(f"    Números expandidos: {seccion['numeros_lista']}")

# También puedes usar findall directamente para extraer las secciones
def encontrar_secciones(texto):
    patron = re.compile(r'([A-Z]):\s*([\d\-,.]+)\s*\((\d+)h\)', re.IGNORECASE)
    return patron.findall(texto)

# Ejemplo
secciones = encontrar_secciones(texto)
print("\nExtracción directa con findall:")
for tipo, numeros, horas in secciones:
    print(f"{tipo}: {numeros} ({horas}h)")
