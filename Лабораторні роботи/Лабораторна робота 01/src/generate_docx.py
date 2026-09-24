"""
Генерація офіційного звіту з Лабораторної роботи № 1 у форматі DOCX (Microsoft Word).
Повністю відповідає вимогам оформлення студентських звітів КрНУ:
- Формат А4, поля по 20 мм (2.0 см);
- Шрифт Times New Roman, розмір 14 пт, міжрядковий інтервал 1.5, абзацний відступ 1.25 см;
- Офіційний титульний аркуш на першій сторінці;
- Вбудовані таблиці з межами та вирівнюванням;
- Вбудовані повнорозмірні графіки (Рисунок 1, 2, 3);
- Математичні формули, лістинг програми, контрольні питання та висновки;
- Рік: 2026.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_border(cell, **kwargs):
    """
    Встановлення меж комірки таблиці:
    set_cell_border(cell, top={"sz": 4, "val": "single", "color": "000000"})
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{kwargs.get("top", "single")}" w:sz="{kwargs.get("top_sz", "4")}" w:color="{kwargs.get("color", "000000")}"/>\n'
        f'  <w:left w:val="{kwargs.get("left", "single")}" w:sz="{kwargs.get("left_sz", "4")}" w:color="{kwargs.get("color", "000000")}"/>\n'
        f'  <w:bottom w:val="{kwargs.get("bottom", "single")}" w:sz="{kwargs.get("bottom_sz", "4")}" w:color="{kwargs.get("color", "000000")}"/>\n'
        f'  <w:right w:val="{kwargs.get("right", "single")}" w:sz="{kwargs.get("right_sz", "4")}" w:color="{kwargs.get("color", "000000")}"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex="F2F2F2"):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def create_report_docx():
    doc = Document()
    
    # 1. Налаштування параметрів сторінки А4 та полів 20 мм
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(20)
    section.right_margin = Mm(20)
    
    # Стиль Normal: Times New Roman, 14 pt, 1.5 інтервал
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(14)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(4)
    normal_style.paragraph_format.space_before = Pt(0)
    
    # Шляхи до файлів
    base_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(base_dir, "..", "assets")
    output_docx = os.path.join(base_dir, "..", "Звіт_ЛР1_Буханцев.docx")
    
    # ==================== ТИТУЛЬНИЙ АРКУШ ====================
    # Шапка університету
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.line_spacing = 1.15
    p_inst.paragraph_format.first_line_indent = Cm(0)
    p_inst.paragraph_format.space_after = Pt(20)
    
    run_inst = p_inst.add_run(
        "МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ\n"
        "КРЕМЕНЧУЦЬКИЙ НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ\n"
        "ІМЕНІ МИХАЙЛА ОСТРОГРАДСЬКОГО\n"
        "НАВЧАЛЬНО-НАУКОВИЙ ІНСТИТУТ ЕЛЕКТРИЧНОЇ ІНЖЕНЕРІЇ\n"
        "ТА ІНФОРМАЦІЙНИХ ТЕХНОЛОГІЙ\n"
        "КАФЕДРА АВТОМАТИЗАЦІЇ ТА ІНФОРМАЦІЙНИХ СИСТЕМ"
    )
    run_inst.bold = True
    run_inst.font.size = Pt(12)
    
    # Предмет
    p_course = doc.add_paragraph()
    p_course.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_course.paragraph_format.line_spacing = 1.15
    p_course.paragraph_format.first_line_indent = Cm(0)
    p_course.paragraph_format.space_before = Pt(25)
    p_course.paragraph_format.space_after = Pt(25)
    
    r_sub = p_course.add_run("ОСВІТНІЙ КОМПОНЕНТ\n«ДОСЛІДЖЕННЯ ОПЕРАЦІЙ»\n\n")
    r_sub.bold = True
    r_sub.font.size = Pt(14)
    
    r_rep = p_course.add_run("ЗВІТ\n")
    r_rep.bold = True
    r_rep.font.size = Pt(22)
    
    r_lab = p_course.add_run("З ЛАБОРАТОРНОЇ РОБОТИ № 1\n")
    r_lab.bold = True
    r_lab.font.size = Pt(16)
    
    r_theme = p_course.add_run("Тема: «Двомірна задача лінійного програмування»\n")
    r_theme.font.size = Pt(14)
    
    r_var = p_course.add_run("Варіант № 4")
    r_var.bold = True
    r_var.font.size = Pt(14)
    
    # Блок виконавця та перевіряючого (вирівняно праворуч)
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.paragraph_format.line_spacing = 1.25
    p_sign.paragraph_format.first_line_indent = Cm(0)
    p_sign.paragraph_format.space_before = Pt(30)
    p_sign.paragraph_format.space_after = Pt(45)
    
    r_sign = p_sign.add_run(
        "Виконав:                \n"
        "здобувач групи КН-24-1  \n"
        "Буханцев М. В.          \n\n"
        "Перевірила:             \n"
        "доцент кафедри АІС      \n"
        "Бурдільна Є. В.         "
    )
    r_sign.font.size = Pt(13.5)
    
    # Місто і рік
    p_city = doc.add_paragraph()
    p_city.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_city.paragraph_format.first_line_indent = Cm(0)
    p_city.paragraph_format.space_after = Pt(0)
    r_city = p_city.add_run("Кременчук 2026")
    r_city.bold = True
    r_city.font.size = Pt(13)
    
    # Розрив сторінки після титулки
    doc.add_page_break()
    
    # ==================== СТОРІНКА 2: ЗМІСТ РОБОТИ ====================
    # Заголовок лабораторної
    p_h1 = doc.add_paragraph()
    p_h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_h1.paragraph_format.first_line_indent = Cm(0)
    p_h1.paragraph_format.space_before = Pt(6)
    p_h1.paragraph_format.space_after = Pt(8)
    r_h1 = p_h1.add_run("Лабораторна робота № 1")
    r_h1.bold = True
    r_h1.font.size = Pt(16)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.first_line_indent = Cm(0)
    r = p_meta.add_run("Тема: ")
    r.bold = True
    p_meta.add_run("Двомірна задача лінійного програмування.")
    
    p_meta2 = doc.add_paragraph()
    p_meta2.paragraph_format.first_line_indent = Cm(0)
    p_meta2.paragraph_format.space_after = Pt(12)
    r = p_meta2.add_run("Мета: ")
    r.bold = True
    p_meta2.add_run("набути навичок з розв’язування двомірних задач лінійного програмування графічним методом.")
    
    # Хід роботи
    p_hid = doc.add_paragraph()
    p_hid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_hid.paragraph_format.first_line_indent = Cm(0)
    p_hid.paragraph_format.space_before = Pt(8)
    p_hid.paragraph_format.space_after = Pt(10)
    r_hid = p_hid.add_run("Хід роботи")
    r_hid.bold = True
    r_hid.font.size = Pt(15)
    
    # 1. Постановка задачі
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.first_line_indent = Cm(1.25)
    p_sec1.paragraph_format.space_before = Pt(6)
    r = p_sec1.add_run("1. Постановка задачі та вихідні дані для варіанта 4")
    r.bold = True
    
    p_text1 = doc.add_paragraph()
    p_text1.paragraph_format.first_line_indent = Cm(1.25)
    p_text1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_text1.add_run(
        "Підприємство випускає столи двох моделей: А і В. Для випуску одного столу моделі А потрібно SA одиниці сировини "
        "та TA одиниці машинного часу. Для випуску одного столу моделі В потрібно SB одиниці сировини та TB одиниць машинного часу. "
        "Прибуток від реалізації одного столу моделі А складає WA грошові одиниці, столу моделі В – WB грошові одиниці. "
        "На підприємстві наявні 1700 одиниць сировини та 1600 одиниць машинного часу. Визначити, яким має бути план виробництва, "
        "щоб підприємство отримало максимальний прибуток."
    )
    
    p_text2 = doc.add_paragraph()
    p_text2.paragraph_format.first_line_indent = Cm(1.25)
    p_text2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_text2.add_run("За індивідуальним завданням варіанта 4 (номер у списку N = 4, коефіцієнт підгрупи 1 k = 1):")
    
    params = [
        "SA = 0.5 · N · k = 0.5 · 4 · 1 = 2.0 (од. сировини на стіл А);",
        "TA = N / k = 4 / 1 = 4.0 (год. машинного часу на стіл А);",
        "SB = N = 4.0 (од. сировини на стіл В);",
        "TB = 0.4 · N · k = 0.4 · 4 · 1 = 1.6 (год. машинного часу на стіл В);",
        "WA = N · k = 4 · 1 = 4.0 (грн прибутку від 1 столу А);",
        "WB = 1.2 · N · k = 1.2 · 4 · 1 = 4.8 (грн прибутку від 1 столу В)."
    ]
    for param in params:
        p_p = doc.add_paragraph()
        p_p.paragraph_format.first_line_indent = Cm(2.0)
        p_p.paragraph_format.space_after = Pt(2)
        p_p.add_run(f"• {param}")
    
    # Таблиця 1
    p_t1_title = doc.add_paragraph()
    p_t1_title.paragraph_format.first_line_indent = Cm(0)
    p_t1_title.paragraph_format.space_before = Pt(8)
    p_t1_title.paragraph_format.space_after = Pt(2)
    r = p_t1_title.add_run("Таблиця 1 – Вихідні дані задачі оптимізації виробництва (Варіант 4)")
    r.bold = True
    
    table1 = doc.add_table(rows=4, cols=4)
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER
    table1.autofit = False
    
    t1_headers = ["Показник", "Стіл моделі А (x1)", "Стіл моделі В (x2)", "Наявний запас на складі"]
    t1_data = [
        ["Витрати сировини на виріб, од.", "2.0", "4.0", "1700 од."],
        ["Витрати машинного часу, год.", "4.0", "1.6", "1600 год."],
        ["Прибуток від реалізації виробу, грош. од.", "4.0", "4.8", "—"]
    ]
    
    for c_idx, h_text in enumerate(t1_headers):
        cell = table1.cell(0, c_idx)
        cell.text = h_text
        set_cell_shading(cell, "F2F2F2")
        set_cell_border(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(12)
        
    for r_idx, row_values in enumerate(t1_data):
        for c_idx, val in enumerate(row_values):
            cell = table1.cell(r_idx + 1, c_idx)
            cell.text = val
            set_cell_border(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.runs[0].font.size = Pt(12)
            
    # Встановлення ширини стовпців
    col_widths = [Cm(7.0), Cm(3.2), Cm(3.2), Cm(3.6)]
    for row in table1.rows:
        for c_idx, w in enumerate(col_widths):
            row.cells[c_idx].width = w

    # 2. Економіко-математична модель ЗЛП
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.first_line_indent = Cm(1.25)
    p_sec2.paragraph_format.space_before = Pt(12)
    r = p_sec2.add_run("2. Економіко-математична модель ЗЛП")
    r.bold = True
    
    p_m1 = doc.add_paragraph()
    p_m1.paragraph_format.first_line_indent = Cm(1.25)
    p_m1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_m1.add_run("Нехай x1 — щоденний обсяг випуску столів моделі А (шт.), а x2 — моделі В (шт.).")
    
    p_m2 = doc.add_paragraph()
    p_m2.paragraph_format.first_line_indent = Cm(1.25)
    r = p_m2.add_run("Цільова функція (критерій ефективності):")
    r.bold = True
    
    p_cf = doc.add_paragraph()
    p_cf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cf.paragraph_format.first_line_indent = Cm(0)
    r = p_cf.add_run("W(x1, x2) = 4.0 · x1 + 4.8 · x2 → max")
    r.bold = True
    r.font.size = Pt(14)
    
    p_m3 = doc.add_paragraph()
    p_m3.paragraph_format.first_line_indent = Cm(1.25)
    r = p_m3.add_run("Система обмежень за виробничими ресурсами:")
    r.bold = True
    
    p_sys = doc.add_paragraph()
    p_sys.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sys.paragraph_format.first_line_indent = Cm(0)
    p_sys.add_run(
        "2.0 · x1 + 4.0 · x2 ≤ 1700  (обмеження за сировиною)\n"
        "4.0 · x1 + 1.6 · x2 ≤ 1600  (обмеження за машинним часом)"
    )
    
    p_m4 = doc.add_paragraph()
    p_m4.paragraph_format.first_line_indent = Cm(1.25)
    r = p_m4.add_run("Граничні умови невід'ємності змінних:")
    r.bold = True
    
    p_nn = doc.add_paragraph()
    p_nn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nn.paragraph_format.first_line_indent = Cm(0)
    p_nn.add_run("x1 ≥ 0,   x2 ≥ 0")

    # 3. Графічний розв'язок ЗЛП
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.first_line_indent = Cm(1.25)
    p_sec3.paragraph_format.space_before = Pt(12)
    r = p_sec3.add_run("3. Графічний розв'язок ЗЛП")
    r.bold = True
    
    p_gr1 = doc.add_paragraph()
    p_gr1.paragraph_format.first_line_indent = Cm(1.25)
    p_gr1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_gr1.add_run(
        "Перетворимо нерівності обмежень у рівності граничних прямих:\n"
        "1. Пряма L1 (сировина): 2 x1 + 4 x2 = 1700  ⇒  x2 = 425 - 0.5 x1. "
        "Точки перетину з осями координат: (0; 425) та (850; 0).\n"
        "2. Пряма L2 (час): 4 x1 + 1.6 x2 = 1600  ⇒  x2 = 1000 - 2.5 x1. "
        "Точки перетину з осями координат: (0; 1000) та (400; 0)."
    )
    
    p_gr2 = doc.add_paragraph()
    p_gr2.paragraph_format.first_line_indent = Cm(1.25)
    p_gr2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_gr2.add_run(
        "Перетин півплощин у першому координатному квадранті утворює область допустимих розв'язків (ОДР) "
        "у вигляді опуклого замкненого чотирикутника OABC з вершинами:\n"
        "• O(0; 0) — початок координат;\n"
        "• A(400; 0) — точка перетину прямої L2 з віссю Ox1;\n"
        "• B(x1*; x2*) — точка взаємного перетину прямих обмежень L1 та L2;\n"
        "• C(0; 425) — точка перетину прямої L1 з віссю Ox2."
    )
    
    p_gr3 = doc.add_paragraph()
    p_gr3.paragraph_format.first_line_indent = Cm(1.25)
    p_gr3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_gr3.add_run(
        "Знайдемо аналітично координати вершини B розв'язанням системи лінійних рівнянь:\n"
        "   2 x1 + 4 x2 = 1700\n"
        "   4 x1 + 1.6 x2 = 1600\n"
        "Помноживши перше рівняння на 2 та віднявши друге, отримуємо: 6.4 x2 = 1800 ⇒ x2* = 281.25.\n"
        "Підставивши у перше рівняння, маємо: 2 x1 = 1700 - 4(281.25) = 575 ⇒ x1* = 287.5."
    )
    
    p_gr4 = doc.add_paragraph()
    p_gr4.paragraph_format.first_line_indent = Cm(1.25)
    p_gr4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_gr4.add_run(
        "Вектор градієнта цільової функції: c⃗ = ∇W = (4.0; 4.8). "
        "Він перпендикулярний до ліній рівня 4 x1 + 4.8 x2 = const. "
        "Переміщуючи лінію рівня у напрямку градієнта від початку координат, встановлюємо, "
        "що крайньою точкою виходу лінії рівня з багатокутника ОДР є вершина B(287.5; 281.25)."
    )
    
    # Вставка Рисунка 1
    fig1_path = os.path.join(assets_dir, "graph_lab1.png")
    if os.path.exists(fig1_path):
        p_fig1 = doc.add_paragraph()
        p_fig1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fig1.paragraph_format.first_line_indent = Cm(0)
        p_fig1.paragraph_format.space_before = Pt(8)
        p_fig1.paragraph_format.space_after = Pt(2)
        doc.add_picture(fig1_path, width=Inches(6.0))
        
        p_cap1 = doc.add_paragraph()
        p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap1.paragraph_format.first_line_indent = Cm(0)
        p_cap1.paragraph_format.space_after = Pt(12)
        r = p_cap1.add_run("Рисунок 1 – Графічне розв’язання задачі лінійного програмування (Варіант 4)")
        r.bold = True
        r.font.size = Pt(12)

    # 4. Розрахунок цільової функції у кутових точках
    p_sec4 = doc.add_paragraph()
    p_sec4.paragraph_format.first_line_indent = Cm(1.25)
    p_sec4.paragraph_format.space_before = Pt(10)
    r = p_sec4.add_run("4. Розрахунок цільової функції у кутових точках ОДР")
    r.bold = True
    
    pts_vals = [
        "W(O) = 4(0) + 4.8(0) = 0 грош. од.;",
        "W(A) = 4(400) + 4.8(0) = 1600 грош. од.;",
        "W(C) = 4(0) + 4.8(425) = 2040 грош. од.;",
        "W(B) = 4(287.5) + 4.8(281.25) = 1150 + 1350 = 2500.00 грош. од."
    ]
    for pt in pts_vals:
        p_p = doc.add_paragraph()
        p_p.paragraph_format.first_line_indent = Cm(2.0)
        p_p.paragraph_format.space_after = Pt(2)
        p_p.add_run(f"• {pt}")
        
    p_wmax = doc.add_paragraph()
    p_wmax.paragraph_format.first_line_indent = Cm(1.25)
    p_wmax.add_run("Отже, максимальний прибуток досягається у точці B і становить ")
    r = p_wmax.add_run("W_max = 2500.00 грн.")
    r.bold = True

    # 5. Обчислення залишків на складах
    p_sec5 = doc.add_paragraph()
    p_sec5.paragraph_format.first_line_indent = Cm(1.25)
    p_sec5.paragraph_format.space_before = Pt(10)
    r = p_sec5.add_run("5. Обчислення залишків ресурсів на складах")
    r.bold = True
    
    p_left = doc.add_paragraph()
    p_left.paragraph_format.first_line_indent = Cm(1.25)
    p_left.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_left.add_run(
        "Підставимо отримані оптимальні значення x1* = 287.5 та x2* = 281.25 у ліві частини рівнянь обмежень:\n"
        "• Використання сировини: 2 · 287.5 + 4 · 281.25 = 1700 од. Залишок: 1700 - 1700 = 0 од. (100% використання);\n"
        "• Використання машинного часу: 4 · 287.5 + 1.6 · 281.25 = 1600 год. Залишок: 1600 - 1600 = 0 год. (100% використання).\n"
        "Висновок: обидва ресурси є повністю вичерпаними (дефіцитними), резервів немає."
    )

    # 6. Цілочисельний аналіз
    p_sec6 = doc.add_paragraph()
    p_sec6.paragraph_format.first_line_indent = Cm(1.25)
    p_sec6.paragraph_format.space_before = Pt(10)
    r = p_sec6.add_run("6. Цілочисельний аналіз для неподільних виробів")
    r.bold = True
    
    p_int1 = doc.add_paragraph()
    p_int1.paragraph_format.first_line_indent = Cm(1.25)
    p_int1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_int1.add_run(
        "Оскільки столи є фізично неподільною продукцією (x1, x2 ∈ ℤ), проведено аналіз цілочисельних вузлів сітки "
        "в околі вершини B(287.5; 281.25) з перевіркою виконання обмежень."
    )
    
    # Вставка Рисунка 2
    fig2_path = os.path.join(assets_dir, "graph_lab1_zoom.png")
    if os.path.exists(fig2_path):
        p_fig2 = doc.add_paragraph()
        p_fig2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fig2.paragraph_format.first_line_indent = Cm(0)
        p_fig2.paragraph_format.space_before = Pt(8)
        p_fig2.paragraph_format.space_after = Pt(2)
        doc.add_picture(fig2_path, width=Inches(5.8))
        
        p_cap2 = doc.add_paragraph()
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.paragraph_format.first_line_indent = Cm(0)
        p_cap2.paragraph_format.space_after = Pt(10)
        r = p_cap2.add_run("Рисунок 2 – Деталізований окіл точки оптимуму та цілочисельний аналіз")
        r.bold = True
        r.font.size = Pt(12)
        
    p_int2 = doc.add_paragraph()
    p_int2.paragraph_format.first_line_indent = Cm(1.25)
    p_int2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_int2.add_run(
        "Серед усіх допустимих цілочисельних планів найкращим є план (286; 282):\n"
        "• Випуск: 286 столів моделі А та 282 столи моделі В;\n"
        "• Прибуток: W = 4(286) + 4.8(282) = 2497.60 грн;\n"
        "• Використання сировини: 2(286) + 4(282) = 1700 од. (залишок 0 од., 100% завантаження);\n"
        "• Використання часу: 4(286) + 1.6(282) = 1595.2 год. (залишок 4.8 год., завантаження 99.7%)."
    )

    # 7. Зведена таблиця та рекомендації
    p_sec7 = doc.add_paragraph()
    p_sec7.paragraph_format.first_line_indent = Cm(1.25)
    p_sec7.paragraph_format.space_before = Pt(10)
    r = p_sec7.add_run("7. Зведена таблиця результатів та рекомендації виробництву")
    r.bold = True
    
    # Вставка Рисунка 3
    fig3_path = os.path.join(assets_dir, "results_summary.png")
    if os.path.exists(fig3_path):
        p_fig3 = doc.add_paragraph()
        p_fig3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fig3.paragraph_format.first_line_indent = Cm(0)
        p_fig3.paragraph_format.space_before = Pt(6)
        p_fig3.paragraph_format.space_after = Pt(2)
        doc.add_picture(fig3_path, width=Inches(6.0))
        
        p_cap3 = doc.add_paragraph()
        p_cap3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap3.paragraph_format.first_line_indent = Cm(0)
        p_cap3.paragraph_format.space_after = Pt(10)
        r = p_cap3.add_run("Рисунок 3 – Зведені показники розв’язку оптимізаційної задачі")
        r.bold = True
        r.font.size = Pt(12)
        
    p_t2_title = doc.add_paragraph()
    p_t2_title.paragraph_format.first_line_indent = Cm(0)
    p_t2_title.paragraph_format.space_before = Pt(8)
    p_t2_title.paragraph_format.space_after = Pt(2)
    r = p_t2_title.add_run("Таблиця 2 – Порівняльний аналіз отриманих планів виробництва")
    r.bold = True
    
    table2 = doc.add_table(rows=8, cols=3)
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER
    table2.autofit = False
    
    t2_headers = ["Показник", "Неперервний розв'язок (x*)", "Практичний цілочисельний план"]
    t2_data = [
        ["Столи моделі А (x1), шт.", "287.50", "286"],
        ["Столи моделі В (x2), шт.", "281.25", "282"],
        ["Максимальний прибуток (W), грн", "2500.00", "2497.60"],
        ["Використання сировини", "1700 / 1700 од. (100%)", "1700 / 1700 од. (100%)"],
        ["Залишок сировини", "0.0 од.", "0.0 од."],
        ["Використання машинного часу", "1600 / 1600 год. (100%)", "1595.2 / 1600 год. (99.7%)"],
        ["Залишок машинного часу", "0.0 год.", "4.8 год."]
    ]
    
    for c_idx, h_text in enumerate(t2_headers):
        cell = table2.cell(0, c_idx)
        cell.text = h_text
        set_cell_shading(cell, "F2F2F2")
        set_cell_border(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(12)
        
    for r_idx, row_values in enumerate(t2_data):
        for c_idx, val in enumerate(row_values):
            cell = table2.cell(r_idx + 1, c_idx)
            cell.text = val
            set_cell_border(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            if r_idx == 2: # Рядок прибутку
                p.runs[0].bold = True
            p.runs[0].font.size = Pt(12)
            
    col_widths2 = [Cm(8.0), Cm(4.5), Cm(4.5)]
    for row in table2.rows:
        for c_idx, w in enumerate(col_widths2):
            row.cells[c_idx].width = w
            
    p_rec = doc.add_paragraph()
    p_rec.paragraph_format.first_line_indent = Cm(1.25)
    p_rec.paragraph_format.space_before = Pt(8)
    p_rec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_rec.add_run(
        "Рекомендації щодо оптимізації виробництва:\n"
        "Оскільки ресурси сировини та машинного часу вичерпані повністю, подальше зростання прибутку "
        "можливе виключно шляхом залучення додаткових обсягів сировини та нарощування фонду обладнання. "
        "У практичній діяльності рекомендується затвердити щоденний план випуску: 286 столів моделі А "
        "та 282 столів моделі В, що забезпечує 2497.60 грн чистого прибутку."
    )

    # 8. Програмна реалізація мовою Python
    p_sec8 = doc.add_paragraph()
    p_sec8.paragraph_format.first_line_indent = Cm(1.25)
    p_sec8.paragraph_format.space_before = Pt(12)
    r = p_sec8.add_run("8. Програмна реалізація мовою Python (лістинг розрахунку)")
    r.bold = True
    
    code_text = (
        "import numpy as np\n"
        "from scipy.optimize import linprog\n\n"
        "N, k = 4, 1\n"
        "S_A, T_A, W_A = 0.5 * N * k, N / k, N * k             # 2.0, 4.0, 4.0\n"
        "S_B, T_B, W_B = float(N), 0.4 * N * k, 1.2 * N * k       # 4.0, 1.6, 4.8\n"
        "LIMIT_RAW, LIMIT_TIME = 1700.0, 1600.0\n\n"
        "# Чисельний розв'язок за допомогою SciPy HiGHS\n"
        "c = [-W_A, -W_B]\n"
        "A_ub = [[S_A, S_B], [T_A, T_B]]\n"
        "b_ub = [LIMIT_RAW, LIMIT_TIME]\n"
        "bounds = [(0, None), (0, None)]\n\n"
        "res_cont = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method='highs')\n"
        "res_int = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, integrality=[1, 1], method='highs')\n\n"
        "print(f'Неперервний: x1={res_cont.x[0]:.2f}, x2={res_cont.x[1]:.2f}, W={-res_cont.fun:.2f}')\n"
        "print(f'Цілочисельний: x1={res_int.x[0]:.0f}, x2={res_int.x[1]:.0f}, W={-res_int.fun:.2f}')"
    )
    p_code = doc.add_paragraph()
    p_code.paragraph_format.first_line_indent = Cm(0)
    p_code.paragraph_format.line_spacing = 1.15
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(8)
    r_code = p_code.add_run(code_text)
    r_code.font.name = "Consolas"
    r_code.font.size = Pt(9.5)

    # 9. Контрольні питання
    p_sec9 = doc.add_paragraph()
    p_sec9.paragraph_format.first_line_indent = Cm(1.25)
    p_sec9.paragraph_format.space_before = Pt(12)
    r = p_sec9.add_run("9. Відповіді на контрольні питання")
    r.bold = True
    
    questions = [
        ("1. Які задачі називають задачами лінійного програмування?",
         "Оптимізаційні задачі, в яких цільова функція є лінійною, а множина допустимих розв'язків задана системою лінійних рівностей або нерівностей."),
        ("2. Що таке цільова функція?",
         "Функція залежності оптимізованого критерію (прибутку, собівартості) від керованих змінних, екстремальне значення якої є метою дослідження."),
        ("3. Як записуються рівняння обмеження?",
         "У вигляді алгебраїчних нерівностей або рівнянь ∑ c_ij · x_j ≤ S_i, де c_ij — питомі норми витрат, а S_i — наявний запас ресурсу."),
        ("4. Які обмеження обов’язково застосовуються до задач оптимального виробництва?",
         "Обмеження за обсягами виробничих ресурсів, умови невід'ємності x_j ≥ 0 та вимоги цілочисельності змінних для неподільної продукції."),
        ("5. Який розв’язок ЗЛП називають оптимальним?",
         "План виробництва з області допустимих розв'язків, який надає цільовій функції екстремальне (найбільше або найменше) значення."),
        ("6. Надайте геометричну інтерпретацію ЗЛП.",
         "Знаходження точки опуклого багатокутника (ОДР), у якій лінія рівня цільової функції досягає екстремуму в напрямку вектора градієнта."),
        ("7. Яка точка допустимої множини розв’язку називається кутовою?",
         "Вершина багатокутника допустимих розв'язків, яка не може бути виражена як опукла комбінація двох інших точок цієї множини."),
        ("8. Поясніть алгоритм графічного методу розв’язання ЗЛП.",
         "Побудова прямих обмежень → побудова ОДР → обчислення градієнта c⃗ → переміщення перпендикулярної лінії рівня до крайньої точки ОДР → знаходження координат оптимуму.")
    ]
    
    for q_text, a_text in questions:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.first_line_indent = Cm(1.25)
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_q = p_q.add_run(q_text + " ")
        r_q.bold = True
        p_q.add_run(a_text)

    # 10. Висновки
    p_conc = doc.add_paragraph()
    p_conc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_conc.paragraph_format.first_line_indent = Cm(0)
    p_conc.paragraph_format.space_before = Pt(14)
    p_conc.paragraph_format.space_after = Pt(8)
    r_conc = p_conc.add_run("Висновки")
    r_conc.bold = True
    r_conc.font.size = Pt(16)
    
    conclusions = [
        "1. У процесі виконання роботи опановано методику побудови економіко-математичних моделей та їх розв'язання графічним методом.",
        "2. За індивідуальними даними Варіанта 4 (N=4, k=1) складено модель максимізації прибутку: W = 4 x1 + 4.8 x2 → max при обмеженнях на сировину (1700 од.) та машинний час (1600 год.).",
        "3. Графічним методом знайдено вершину ОДР B(287.50; 281.25), що забезпечує глобальний максимум прибутку 2500.00 грн при 100% завантаженні сировини та обладнання.",
        "4. Сформовано практичний цілочисельний план випуску: 286 столів моделі А та 282 столи моделі В, що дає 2497.60 грн чистого прибутку при нульовому залишку сировини.",
        "5. Числові та графічні розрахунки верифіковано в середовищі Python за допомогою бібліотек NumPy та SciPy (метод HiGHS)."
    ]
    for c_text in conclusions:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.first_line_indent = Cm(1.25)
        p_c.paragraph_format.space_after = Pt(3)
        p_c.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_c.add_run(c_text)

    doc.save(output_docx)
    print(f"Успішно збережено файл DOCX: {output_docx}")

if __name__ == "__main__":
    create_report_docx()
