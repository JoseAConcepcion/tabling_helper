def export_to_md(matrix, idx):
    filename = f"tabla_{idx+1}.md"
    with open(filename, 'w', encoding='utf-8') as f:
        f.write("| | " + " | ".join(matrix[0][1:]) + " |\n")
        
        f.write("|" + "---|" * (len(matrix[0])) + "\n")
        
        for i in range(1, len(matrix)):
            fila = [matrix[i][0]] + matrix[i][1:]
            f.write("| " + " | ".join(fila) + " |\n")

    print(f"Tabla {idx+1} exportada a {filename}")

def export_excel(self):
    pass
def export_txt(self):
    pass