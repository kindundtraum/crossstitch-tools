def expand_stitches(stitches):
    rows = len(stitches)
    cols = len(stitches[0])
    
    # Результирующая сетка дырочек на 1 строку и 1 столбец больше
    holes = [[0 for _ in range(cols + 1)] for _ in range(rows + 1)]
    
    for r in range(rows):
        for c in range(cols):
            val = stitches[r][c]
            
            if val == 1:
                # Крестик: 4 дырочки (перекрываются с соседями)
                holes[r][c] = 1
                holes[r][c+1] = 1
                holes[r+1][c] = 1
                holes[r+1][c+1] = 1
                
    return holes

def print_openscad_matrix(matrix, name="hole_pattern"):
    print(f"{name} = [")
    for row in matrix:
        row_str = ", ".join(str(x) for x in row)
        print(f"    [{row_str}],")
    print("];")

if __name__ == "__main__":
    # Матрица крестиков
    #   1 - есть крестик
    #   0 - крестика нет
    stitch_pattern = [
        [0, 0, 1, 0, 0],
        [0, 1, 1, 1, 0],
        [1, 1, 1, 1, 1],
    ]

    print("Исходный шаблон (крестики):")
    for row in stitch_pattern:
        print(row)
    
    expanded = expand_stitches(stitch_pattern)
    
    print("\n" + "="*40)
    print("Готовый массив для OpenSCAD:")
    print("="*40)
    print_openscad_matrix(expanded)