import re
from docx import Document
from collections import defaultdict

FILE_PATH = 'ejemplo4.docx'

def limpiar_texto(texto):
    texto =  texto.replace('\n', ' ')
    texto = texto.replace(' - ', '-')
    texto = texto.replace(' -', '-')
    texto = texto.replace('- ', '-')
    texto = texto.replace(' – ', '-')
    texto = texto.replace(' –', '-')
    texto = texto.replace('– ', '-')
    texto = re.sub(r'\s{2,}', ' ', texto)
    return texto


def analizar_estructura_tabla_completa(tabla):
    """
    Analiza la estructura de la tabla incluyendo filas duplicadas.
    """
    if not tabla:
        return
    
    print(f"\n{'='*80}")
    print("ANÁLISIS COMPLETO DE ESTRUCTURA DE TABLA")
    print(f"{'='*80}")
    
    # Analizar encabezados
    if tabla:
        encabezados = tabla[0]
        print(f"Encabezados: {encabezados}")
    
    # Analizar filas duplicadas por horario
    horarios = defaultdict(list)
    for fila_idx, fila in enumerate(tabla[1:], start=1):  # Omitir encabezados
        if fila:
            horario = fila[0].strip() if len(fila) > 0 else ""
            if horario:
                horarios[horario].append(fila_idx)
    
    print(f"\nAnálisis de filas por horario:")
    for horario, filas in horarios.items():
        if len(filas) > 1:
            print(f"  {horario}: {len(filas)} filas duplicadas (filas {filas})")
        else:
            print(f"  {horario}: 1 fila")
    
    # Dimensiones
    num_filas = len(tabla)
    num_columnas = max(len(fila) for fila in tabla) if tabla else 0
    print(f"\nDimensiones finales: {num_filas} filas × {num_columnas} columnas")

def formatear_tabla_mejorada(tabla):
    """
    Fusiona columnas duplicadas en una tabla, fusionando inteligentemente el contenido.
    Para celdas con contenido similar, solo añade las diferencias.
    """
    if not tabla:
        return tabla
    
    # Obtener encabezados de la primera fila
    encabezados = tabla[0]
    
    # Crear un diccionario para mapear encabezados a índices de columna
    encabezado_a_indices = defaultdict(list)
    
    for idx, encabezado in enumerate(encabezados):
        encabezado_limpio = encabezado.strip()
        if encabezado_limpio:
            encabezado_a_indices[encabezado_limpio].append(idx)
        else:
            encabezado_a_indices[f"_vacio_{idx}"].append(idx)
    
    # Identificar columnas que se fusionarán
    columnas_a_mantener = []
    grupos_a_fusionar = {}  # Diccionario: columna_principal -> [columnas_secundarias]
    
    for encabezado, indices in encabezado_a_indices.items():
        if len(indices) > 1:
            # Mantener la primera columna, las demás se fusionarán en ella
            columna_principal = indices[0]
            columnas_a_mantener.append(columna_principal)
            grupos_a_fusionar[columna_principal] = indices[1:]
        else:
            columnas_a_mantener.append(indices[0])
    
    # Ordenar las columnas a mantener
    columnas_a_mantener.sort()
    
    # Crear nueva tabla formateada
    tabla_formateada = []
    
    for fila_idx, fila in enumerate(tabla):
        nueva_fila = []
        
        for col_idx in columnas_a_mantener:
            if col_idx >= len(fila):
                nueva_fila.append("")
                continue
            
            contenido_principal = fila[col_idx]
            
            # Verificar si esta columna tiene columnas para fusionar
            if col_idx in grupos_a_fusionar:
                contenidos_a_fusionar = [contenido_principal]
                
                for dup_idx in grupos_a_fusionar[col_idx]:
                    if dup_idx < len(fila):
                        contenido_dup = fila[dup_idx].strip()
                        if contenido_dup:
                            contenidos_a_fusionar.append(contenido_dup)
                
                # Si solo hay un contenido o todos son iguales, usar solo uno
                if len(set(contenidos_a_fusionar)) == 1:
                    contenido_final = contenido_principal
                else:
                    # Análisis inteligente: extraer partes únicas
                    contenido_final = fusionar_contenidos_inteligentemente(contenidos_a_fusionar)
            else:
                contenido_final = contenido_principal
            
            nueva_fila.append(contenido_final)
        
        tabla_formateada.append(nueva_fila)
    
    return tabla_formateada

def fusionar_contenidos_inteligentemente(contenidos):
    """
    Fusiona múltiples contenidos de forma inteligente.
    Si los contenidos comparten una parte común, solo se añaden las diferencias.
    """
    if not contenidos:
        return ""
    
    # Si solo hay un contenido, devolverlo
    if len(contenidos) == 1:
        return contenidos[0]
    
    # Verificar si todos los contenidos son iguales
    if len(set(contenidos)) == 1:
        return contenidos[0]
    
    # Extraer palabras y frases comunes
    palabras_contenidos = [re.findall(r'\b\w+\b', c.lower()) for c in contenidos]
    
    # Encontrar palabras comunes a todos los contenidos
    palabras_comunes = set(palabras_contenidos[0])
    for palabras in palabras_contenidos[1:]:
        palabras_comunes = palabras_comunes.intersection(set(palabras))
    
    # Si hay muchas palabras comunes, probablemente sea el mismo contenido
    if len(palabras_comunes) > 3:  # Umbral ajustable
        # Devolver el primer contenido (o el más completo)
        return max(contenidos, key=len)
    
    # Para el caso específico de horarios como "Cálc. Dif. CP:1-16 (2C)"
    # Buscar patrones comunes
    patrones_comunes = []
    for contenido in contenidos:
        # Buscar patrones como "Cálc. Dif.", "Quím. General", etc.
        patrones = re.findall(r'([A-Za-zÁÉÍÓÚáéíóúñÑ\.]+\s+[A-Za-zÁÉÍÓÚáéíóúñÑ\.]+)', contenido)
        patrones_comunes.append(set(patrones))
    
    # Encontrar patrones comunes a todos
    if patrones_comunes:
        patrones_interseccion = set.intersection(*patrones_comunes)
        if patrones_interseccion:
            diferencias = []
            for contenido in contenidos:
                for patron in patrones_interseccion:
                    if patron in contenido:
                        parte_diferente = contenido.replace(patron, "").strip()
                        if parte_diferente and parte_diferente not in diferencias:
                            diferencias.append(parte_diferente)
            
            if diferencias:
                patron_comun = next(iter(patrones_interseccion))
                if len(diferencias) == 1:
                    return f"{patron_comun} {diferencias[0]}"
                else:
                    return f"{patron_comun} ({' | '.join(diferencias)})"
    
    # Pero primero eliminar duplicados exactos
    contenidos_unicos = []
    for contenido in contenidos:
        if contenido and contenido not in contenidos_unicos:
            contenidos_unicos.append(contenido)
    
    if len(contenidos_unicos) == 1:
        return contenidos_unicos[0]
    else:
        return " | ".join(contenidos_unicos)

def formatear_tabla_completa(tabla):
    """
    Fusiona inteligentemente columnas y filas duplicadas en una tabla.
    """
    if not tabla:
        return tabla
    
    # PRIMERO: Fusionar columnas duplicadas
    tabla = formatear_tabla_mejorada(tabla)
    
    # SEGUNDO: Identificar y fusionar filas duplicadas
    # Las filas duplicadas son aquellas con la misma primera celda (horario)
    filas_fusionadas = []
    filas_por_horario = defaultdict(list)
    
    # Agrupar filas por horario (primera celda)
    for fila in tabla:
        if fila:  # Asegurar que la fila no esté vacía
            horario = fila[0].strip()
            filas_por_horario[horario].append(fila)
    
    # Fusionar filas con el mismo horario
    for horario, filas_grupo in filas_por_horario.items():
        if len(filas_grupo) == 1:
            # Solo una fila con este horario
            filas_fusionadas.append(filas_grupo[0])
        else:
            # Múltiples filas con el mismo horario, fusionarlas
            fila_fusionada = fusionar_filas_duplicadas(filas_grupo)
            filas_fusionadas.append(fila_fusionada)
    
    # Ordenar las filas por horario (opcional, pero útil para horarios)
    filas_fusionadas = ordenar_filas_por_horario(filas_fusionadas)
    
    return filas_fusionadas

def fusionar_filas_duplicadas(filas):
    """
    Fusiona múltiples filas con el mismo horario en una sola fila.
    """
    if not filas:
        return []
    
    # La primera celda (horario) es la misma para todas
    horario = filas[0][0]
    
    # Para cada columna, fusionar los contenidos de todas las filas
    num_columnas = len(filas[0])
    fila_fusionada = [horario]  # Comenzar con el horario
    
    for col_idx in range(1, num_columnas):  # Empezar desde la columna 1 (omitir horario)
        contenidos_columna = []
        
        for fila in filas:
            if col_idx < len(fila):
                contenido = fila[col_idx].strip()
                if contenido:  # Solo añadir si no está vacío
                    contenidos_columna.append(contenido)
        
        if not contenidos_columna:
            # Columna vacía en todas las filas
            contenido_final = ""
        elif len(contenidos_columna) == 1:
            # Solo un contenido en esta columna
            contenido_final = contenidos_columna[0]
        else:
            # Múltiples contenidos, fusionar inteligentemente
            contenido_final = fusionar_contenidos_inteligentemente(contenidos_columna)
        
        fila_fusionada.append(contenido_final)
    
    return fila_fusionada

def ordenar_filas_por_horario(filas):
    """
    Ordena las filas por el número de horario (primera columna).
    """
    if len(filas) <= 1:
        return filas
    
    # Extraer el número del horario (ej: "1 8:30am a 10:05am" -> 1)
    def obtener_numero_horario(fila):
        if fila and fila[0]:
            # Buscar el primer número en la cadena
            match = re.search(r'^(\d+)', fila[0].strip())
            if match:
                return int(match.group(1))
        return 0
    
    # Ordenar por número de horario
    return sorted(filas, key=obtener_numero_horario)

def procesar_documento_completo(file_path):
    """
    Procesa el documento Word, extrae tablas y las formatea completamente.
    """
    print(f"\n{'#'*80}")
    print("EXTRACCIÓN Y FORMATEO COMPLETO DE TABLAS")
    print(f"{'#'*80}")
    
    # Extraer tablas originales
    doc = Document(file_path)
    tablas_originales = []
    
    for tabla_idx, tabla in enumerate(doc.tables):
        datos_tabla = []
        for fila in tabla.rows:
            datos_fila = [limpiar_texto(celda.text.strip()) for celda in fila.cells]
            datos_tabla.append(datos_fila)
        tablas_originales.append(datos_tabla)
    
    print(f"\n{'#'*80}")
    print("TABLAS FORMATEADAS COMPLETAMENTE (columnas y filas fusionadas)")
    print(f"{'#'*80}")
    
    # Formatear cada tabla completamente
    tablas_formateadas = []
    
    for idx, tabla in enumerate(tablas_originales):
        tabla_formateada = formatear_tabla_completa(tabla)
        tablas_formateadas.append(tabla_formateada)
        
        print(f"\n--- Tabla {idx + 1} (Formateada Completa) ---")
        for fila_idx, fila in enumerate(tabla_formateada):
            print(f"Fila {fila_idx}: Columnas={len(fila)} -> {fila}")
    
    return tablas_formateadas

def guardar_tablas_como_csv(tablas_formateadas, archivo_base='tabla'):
    """
    Guarda las tablas formateadas como archivos CSV para su uso posterior.
    """
    import csv
    
    for idx, tabla in enumerate(tablas_formateadas):
        nombre_archivo = f'{archivo_base}_{idx + 1}.csv'
        
        with open(nombre_archivo, 'w', newline='', encoding='utf-8') as archivo_csv:
            escritor = csv.writer(archivo_csv)
            escritor.writerows(tabla)
        
        print(f"\nTabla {idx + 1} guardada como '{nombre_archivo}'")

if __name__ == "__main__":
    # Procesar el documento con la versión completa
    tablas_formateadas = procesar_documento_completo(FILE_PATH)
    
    # Analizar la estructura de cada tabla
    for idx, tabla in enumerate(tablas_formateadas):
        analizar_estructura_tabla_completa(tabla)
    
    # Opcional: guardar como CSV
    guardar_tablas_como_csv(tablas_formateadas)