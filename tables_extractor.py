import re
from docx import Document

def export_to_md(matrix, idx):
    filename = f"tabla_{idx+1}.md"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write("| | " + " | ".join(matrix[0][1:]) + " |\n")
        
        f.write("|" + "---|" * (len(matrix[0])) + "\n")
        
        for i in range(1, len(matrix)):
            fila = [matrix[i][0]] + matrix[i][1:]
            f.write("| " + " | ".join(fila) + " |\n")

    print(f"Tabla {idx+1} exportada a {filename}")

def extract_data_from_tables(path_to_doc):
    wordDoc = Document(path_to_doc)
    group_tables = []
    group_tables_no_dupes = []

    for table in wordDoc.tables:
    
        matrix = [["" for _ in range(len(table.columns))] for _ in range(len(table.rows))]

        for column_index, column in enumerate(table.columns):
            for cell_index, cell in enumerate(column.cells):
                formatted_text = re.sub(r'\s+', ' ', cell.text.replace('\n', ' ')).strip()
                matrix[cell_index][column_index] = formatted_text
            print()
        group_tables.append(matrix)

    group_tables_no_dupes = []

    for idx,matrix in enumerate(group_tables):
        days_set = list(dict.fromkeys(matrix[0]))
        hours_set = list(dict.fromkeys(fila[0] for fila in matrix))
        
        cell_data = {(day, hour): set() for day in days_set for hour in hours_set}
        
        for i in range(1, len(matrix)):
            hour_val = matrix[i][0]
            for j in range(1, len(matrix[0])):
                day_val = matrix[0][j]
                value = matrix[i][j]
                if value != "" and value is not None:
                    cell_data[(day_val, hour_val)].add(value)
        
        matrix_no_dupes = [["" for _ in range(len(days_set))] for _ in range(len(hours_set))]
        
        for i in range(len(hours_set)):
            matrix_no_dupes[i][0] = hours_set[i]
        for j in range(len(days_set)):
            matrix_no_dupes[0][j] = days_set[j]
        
        for i in range(1, len(hours_set)):
            for j in range(1, len(days_set)):
                key = (days_set[j], hours_set[i])
                conjunto = cell_data.get(key, set())
                matrix_no_dupes[i][j] = " / ".join(sorted(conjunto)) if conjunto else ""
        
        group_tables_no_dupes.append(matrix_no_dupes)
        return matrix_no_dupes
