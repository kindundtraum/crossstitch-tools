/* 
  Матрица формы
  0 - пустота
  1 - обычный крестик (маленькая дырка)
  2 - подвес (большая дырка)
*/

/*
// Домик
pattern = [
    [0, 0, 0, 0, 1, 1, 0],
    [0, 0, 0, 0, 1, 1, 0],
    [0, 0, 1, 1, 1, 1, 0],
    [0, 1, 1, 1, 1, 1, 0],
    [1, 1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1]
];
*/

// Сердечко
pattern = [
    [0, 1, 1, 0, 0, 0, 0, 1, 1, 0],
    [1, 1, 1, 1, 2, 2, 1, 1, 1, 1],
    [1, 1, 1, 1, 2, 2, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    [0, 1, 1, 1, 1, 1, 1, 1, 1, 0],
    [0, 0, 1, 1, 1, 1, 1, 1, 0, 0],
    [0, 0, 0, 1, 1, 1, 1, 0, 0, 0],
    [0, 0, 0, 0, 1, 1, 0, 0, 0, 0]
];

/* 
  Параметры заготовки
*/
count = 14;            // Каунт (крестиков на дюйм)
thickness = 1;         // Толщина заготовки, мм
hole_diameter = 1.3;   // Диаметр дырочки, мм
margin = 1.5;          // Отступ от центра крестика до края, мм

// Параметры подвеса
hang_width = 3;        // Ширина подвеса (мм)

// Вычисляемые параметры
inch = 25.4;
pitch = inch / count;    // Шаг сетки, мм (1.5875)
cell_radius = pitch / 2;

rows = len(pattern);
cols = len(pattern[0]);

/* 
  Генерация заготовки
*/
module embroidery_blank_2d() {
    difference() {
        // --- ФОРМИРОВАНИЕ КРАЯ ---
        union() {
            for (r = [0 : rows - 2]) {
                for (c = [0 : cols - 2]) {
                    s = (pattern[r][c] > 0 ? 1 : 0) +
                        (pattern[r][c+1] > 0 ? 1 : 0) +
                        (pattern[r+1][c] > 0 ? 1 : 0) +
                        (pattern[r+1][c+1] > 0 ? 1 : 0);
                    
                    if (s >= 2) {
                        hull() {
                            if (pattern[r][c] > 0) add_circle(r, c);
                            if (pattern[r][c+1] > 0) add_circle(r, c+1);
                            if (pattern[r+1][c] > 0) add_circle(r+1, c);
                            if (pattern[r+1][c+1] > 0) add_circle(r+1, c+1);
                        }
                    }
                }
            }
        }
        
        // --- ВЫРЕЗАНИЕ ДЫРОЧЕК ---
        for (r = [0 : rows - 1]) {
            for (c = [0 : cols - 1]) {
                x = c * pitch;
                y = -r * pitch;
                
                if (pattern[r][c] == 1) {
                    // Обычный крестик
                    translate([x, y]) 
                        circle(d = hole_diameter, $fn=60);
                } 
                else if (pattern[r][c] == 2) {
                    // Проверяем, является ли эта двойка ВТОРОЙ в паре (справа или снизу)
                    is_second = (c > 0 && pattern[r][c-1] == 2) || (r > 0 && pattern[r-1][c] == 2);
                    
                    // Если это не вторая двойка, значит мы рисуем вырез
                    if (!is_second) {
                        // Ищем соседа справа
                        has_right = (c < cols - 1 && pattern[r][c+1] == 2);
                        // Ищем соседа снизу
                        has_bottom = (r < rows - 1 && pattern[r+1][c] == 2);
                        
                        // Сдвигаем на полшага в сторону соседа
                        shift_x = has_right ? pitch / 2 : 0;
                        shift_y = has_bottom ? -pitch / 2 : 0;
                        
                        translate([x + shift_x, y + shift_y]) 
                            single_hang_hole();
                    }
                }
            }
        }
    }
}

// Вспомогательный модуль для контура
module add_circle(r, c) {
    x = c * pitch;
    y = -r * pitch;
    translate([x, y]) 
        circle(r = cell_radius + margin, $fn=60);
}

// Модуль для вырезания подвеса
module single_hang_hole() {
    // Мы используем 2D-объекты для вырезания
    difference() {
        // Основной круг
        circle(d = hang_width, $fn=60);
    }
}

// Выдавливаем
linear_extrude(height = thickness, $fn=120)
    embroidery_blank_2d();
