"""
Генерація офіційного звіту з лабораторної роботи № 1 у форматі PDF.
Вимоги:
- Без жирного шрифту взагалі (font-weight: normal усюди);
- Уникати списків (без <ul>, <ol>, <li>, без маркерів •);
- Без довгих тире (лише звичайний дефіс -);
- Усі рисунки обов'язково підписані;
- Коефіцієнт k = 0.5, варіант 4 (N = 4);
- Формат А4, поля 20 мм;
- Шрифт Times New Roman, 13.5-14 пт, інтервал 1.4-1.5, абзац 1.25 см;
- Окремий титульний аркуш на 1 сторінку;
- Рік: 2026.
"""

import os
import base64
import subprocess

def get_b64_image(rel_path):
    abs_path = os.path.abspath(rel_path)
    with open(abs_path, 'rb') as f:
        return 'data:image/png;base64,' + base64.b64encode(f.read()).decode('utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
assets_dir = os.path.join(BASE_DIR, "..", "assets")

img1_b64 = get_b64_image(os.path.join(assets_dir, "graph_lab1.png"))
img2_b64 = get_b64_image(os.path.join(assets_dir, "graph_lab1_zoom.png"))
img3_b64 = get_b64_image(os.path.join(assets_dir, "results_summary.png"))

html_template = """<!DOCTYPE html>
<html lang="uk">
<head>
<meta charset="UTF-8">
<title>Звіт з лабораторної роботи № 1 - Буханцев М. В.</title>
<style>
  @page {
    size: A4 portrait;
    margin: 20mm 20mm 20mm 20mm;
    @bottom-right {
      content: counter(page);
      font-family: "Times New Roman", Times, serif;
      font-size: 11pt;
      font-weight: normal;
    }
  }

  @page :first {
    @bottom-right {
      content: "";
    }
  }
  
  * {
    box-sizing: border-box;
    font-weight: normal !important;
  }

  body {
    font-family: "Times New Roman", Times, serif;
    font-size: 13.5pt;
    line-height: 1.45;
    color: #000;
    margin: 0;
    padding: 0;
  }

  /* ТИТУЛЬНИЙ АРКУШ */
  .title-page {
    page-break-after: always;
    height: 252mm;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    text-align: center;
    padding: 0;
    margin: 0;
  }
  .title-header {
    text-transform: uppercase;
    font-size: 11.5pt;
    line-height: 1.35;
  }
  .title-middle {
    margin: auto 0;
  }
  .title-sub {
    font-size: 13.5pt;
    letter-spacing: 0.5px;
    margin-bottom: 16pt;
  }
  .title-main {
    font-size: 20pt;
    margin-bottom: 8pt;
  }
  .title-lab-num {
    font-size: 14pt;
    margin-bottom: 12pt;
  }
  .title-sign-block {
    width: 60%;
    margin-left: auto;
    text-align: left;
    font-size: 13.5pt;
    line-height: 1.35;
    margin-bottom: 15mm;
  }
  .title-footer {
    text-align: center;
    font-size: 13.5pt;
  }

  /* ОСНОВНИЙ ЗМІСТ */
  .content {
    margin-top: 5mm;
  }

  p {
    text-align: justify;
    text-indent: 1.25cm;
    margin: 0 0 5pt 0;
  }
  p.no-indent {
    text-indent: 0;
  }
  
  h2, h3, h4 {
    font-family: "Times New Roman", Times, serif;
    color: #000;
    page-break-after: avoid;
    font-weight: normal !important;
  }
  
  h2 {
    font-size: 14pt;
    text-align: center;
    margin: 12pt 0 8pt 0;
  }
  h3 {
    font-size: 13.5pt;
    margin: 12pt 0 5pt 0;
    text-indent: 1.25cm;
  }
  
  .table-title {
    text-indent: 0;
    margin: 10pt 0 3pt 0;
    text-align: left;
    page-break-after: avoid;
  }
  
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 4pt 0 10pt 0;
    font-size: 11.5pt;
    line-height: 1.2;
    page-break-inside: avoid;
  }
  table th, table td {
    border: 1px solid #000;
    padding: 4pt 6pt;
    text-align: center;
    font-weight: normal !important;
  }
  table th {
    background-color: #f2f2f2;
  }
  table td.text-left {
    text-align: left;
  }

  .fig-container {
    text-align: center;
    margin: 10pt 0;
    page-break-inside: avoid;
  }
  .fig-container img {
    max-width: 92%;
    max-height: 120mm;
    height: auto;
    border: 1px solid #ccc;
  }
  .fig-caption {
    font-size: 11.5pt;
    margin-top: 4pt;
    text-align: center;
    text-indent: 0;
    page-break-before: avoid;
  }
  
  .formula {
    text-align: center;
    margin: 6pt 0;
    font-size: 13.5pt;
    font-style: italic;
    text-indent: 0;
    page-break-inside: avoid;
  }

  .code-block {
    background: #f8f9fa;
    border: 1px solid #ddd;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9.5pt;
    line-height: 1.25;
    padding: 6pt 8pt;
    margin: 6pt 0 10pt 0;
    white-space: pre-wrap;
    text-indent: 0;
    page-break-inside: avoid;
  }
</style>
</head>
<body>

<!-- ТИТУЛЬНИЙ АРКУШ (СТОРІНКА 1) -->
<div class="title-page">
  <div class="title-header">
    МІНІСТЕРСТВО ОСВІТИ І НАУКИ УКРАЇНИ<br>
    КРЕМЕНЧУЦЬКИЙ НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ<br>
    ІМЕНІ МИХАЙЛА ОСТРОГРАДСЬКОГО<br>
    НАВЧАЛЬНО-НАУКОВИЙ ІНСТИТУТ ЕЛЕКТРИЧНОЇ ІНЖЕНЕРІЇ<br>
    ТА ІНФОРМАЦІЙНИХ ТЕХНОЛОГІЙ<br>
    КАФЕДРА АВТОМАТИЗАЦІЇ ТА ІНФОРМАЦІЙНИХ СИСТЕМ
  </div>

  <div class="title-middle">
    <div class="title-sub">ОСВІТНІЙ КОМПОНЕНТ<br>«ДОСЛІДЖЕННЯ ОПЕРАЦІЙ»</div>
    <div class="title-main">ЗВІТ</div>
    <div class="title-lab-num">З ЛАБОРАТОРНОЇ РОБОТИ №1</div>
  </div>

  <div>
    <div class="title-sign-block">
      <p class="no-indent">Виконав:<br>
      здобувач групи КН-24-1<br>
      Буханцев М. В.</p>
      <br>
      <p class="no-indent">Перевірила:<br>
      доцент кафедри АІС<br>
      Бурдільна Є. В.</p>
    </div>

    <div class="title-footer">
      Кременчук 2026
    </div>
  </div>
</div>

<!-- ОСНОВНИЙ ЗМІСТ (СТОРІНКИ 2+) -->
<div class="content">
  <p class="no-indent">Тема: Двомірна задача лінійного програмування.</p>
  <p class="no-indent">Мета: набути навичок з розв’язування двомірних задач лінійного програмування графічним методом.</p>
  
  <h2>Хід роботи</h2>

  <p>Підприємство випускає столи двох моделей: А і В. Для випуску одного столу моделі А потрібно SA одиниці сировини та TA одиниці машинного часу. Для випуску одного столу моделі В потрібно SB одиниці сировини та TB одиниць машинного часу. Прибуток від реалізації одного столу моделі А складає WA грошові одиниці, столу моделі В - WB грошові одиниці. На підприємстві наявні 1700 одиниць сировини та 1600 одиниць машинного часу. Визначити, яким має бути план виробництва, щоб підприємство отримало максимальний прибуток.</p>

  <p>За індивідуальним завданням варіанта 4 (номер у списку N = 4, коефіцієнт підгрупи 1 k = 0.5):</p>
  
  <p>SA = 0.5 &middot; N &middot; k = 0.5 &middot; 4 &middot; 0.5 = 1.0 (од. сировини на стіл А);</p>
  <p>TA = N / k = 4 / 0.5 = 8.0 (год. машинного часу на стіл А);</p>
  <p>SB = N = 4.0 (од. сировини на стіл В);</p>
  <p>TB = 0.4 &middot; N &middot; k = 0.4 &middot; 4 &middot; 0.5 = 0.8 (год. машинного часу на стіл В);</p>
  <p>WA = N &middot; k = 4 &middot; 0.5 = 2.0 (грн прибутку від 1 столу А);</p>
  <p>WB = 1.2 &middot; N &middot; k = 1.2 &middot; 4 &middot; 0.5 = 2.4 (грн прибутку від 1 столу В).</p>

  <div class="table-title">Таблиця 1 - Вихідні дані задачі оптимізації виробництва (Варіант 4)</div>
  <table>
    <tr>
      <th>Показник</th>
      <th>Стіл моделі А (x1)</th>
      <th>Стіл моделі В (x2)</th>
      <th>Наявний запас на складі</th>
    </tr>
    <tr>
      <td class="text-left">Витрати сировини на виріб, од.</td>
      <td>1.0</td>
      <td>4.0</td>
      <td>1700 од.</td>
    </tr>
    <tr>
      <td class="text-left">Витрати машинного часу, год.</td>
      <td>8.0</td>
      <td>0.8</td>
      <td>1600 год.</td>
    </tr>
    <tr>
      <td class="text-left">Прибуток від реалізації виробу, грош. од.</td>
      <td>2.0</td>
      <td>2.4</td>
      <td>-</td>
    </tr>
  </table>

  <p>2. Економіко-математична модель ЗЛП</p>
  <p>Нехай x1 - щоденний обсяг випуску столів моделі А (шт.), а x2 - щоденний обсяг випуску столів моделі В (шт.).</p>
  
  <p>Цільова функція (критерій ефективності):</p>
  <div class="formula">W(x1, x2) = 2.0 &middot; x1 + 2.4 &middot; x2 &rarr; max</div>
  
  <p>Система обмежень за виробничими ресурсами:</p>
  <div class="formula">
    1.0 &middot; x1 + 4.0 &middot; x2 &le; 1700 (обмеження за сировиною)<br>
    8.0 &middot; x1 + 0.8 &middot; x2 &le; 1600 (обмеження за машинним часом)
  </div>
  
  <p>Граничні умови невід'ємності змінних:</p>
  <div class="formula">x1 &ge; 0, &nbsp; x2 &ge; 0</div>

  <p>3. Графічний розв'язок ЗЛП</p>
  <p>Перетворимо нерівності обмежень у рівності граничних прямих:</p>
  <p>1. Пряма L1 (сировина): 1.0 x1 + 4.0 x2 = 1700 &rArr; x2 = 425 - 0.25 x1. Точки перетину з осями координат: (0; 425) та (1700; 0).</p>
  <p>2. Пряма L2 (час): 8.0 x1 + 0.8 x2 = 1600 &rArr; x2 = 2000 - 10 x1. Точки перетину з осями координат: (0; 2000) та (200; 0).</p>
  
  <p>Перетин півплощин у першому координатному квадранті утворює область допустимих розв'язків (ОДР) у вигляді опуклого замкненого чотирикутника OABC з вершинами: вершина O(0; 0) - початок координат; вершина A(200; 0) - точка перетину прямої L2 з віссю Ox1; вершина B(x1*; x2*) - точка взаємного перетину прямих обмежень L1 та L2; вершина C(0; 425) - точка перетину прямої L1 з віссю Ox2.</p>

  <p>Знайдемо аналітично координати вершини B розв'язанням системи лінійних рівнянь:</p>
  <div class="formula">
    1.0 x1 + 4.0 x2 = 1700<br>
    8.0 x1 + 0.8 x2 = 1600
  </div>
  <p>Помноживши перше рівняння на 8 та віднявши друге, отримуємо: 31.2 x2 = 12000 &rArr; x2* = 5000 / 13 &asymp; 384.62.</p>
  <p>Підставивши у перше рівняння, маємо: 1.0 x1 = 1700 - 4(384.615) = 2100 / 13 &asymp; 161.54.</p>

  <p>Вектор градієнта цільової функції: c = grad W = (2.0; 2.4). Він перпендикулярний до ліній рівня 2.0 x1 + 2.4 x2 = const. Переміщуючи лінію рівня у напрямку градієнта від початку координат, встановлюємо, що крайньою точкою виходу лінії рівня з багатокутника ОДР є вершина B(161.54; 384.62).</p>

  <div class="fig-container">
    <img src="__IMG1__" alt="Графічне розв'язання задачі лінійного програмування (Варіант 4)">
    <div class="fig-caption">Рисунок 1 - Графічне розв'язання задачі лінійного програмування (Варіант 4)</div>
  </div>

  <p>4. Розрахунок цільової функції у кутових точках ОДР</p>
  <p>W(O) = 2.0 &middot; 0 + 2.4 &middot; 0 = 0.00 грош. од.;</p>
  <p>W(A) = 2.0 &middot; 200 + 2.4 &middot; 0 = 400.00 грош. од.;</p>
  <p>W(C) = 2.0 &middot; 0 + 2.4 &middot; 425 = 1020.00 грош. од.;</p>
  <p>W(B) = 2.0 &middot; 161.54 + 2.4 &middot; 384.62 = 323.08 + 923.08 = 1246.15 грош. од.</p>
  <p>Отже, максимальний прибуток досягається у точці B і становить W_max = 1246.15 грн.</p>

  <p>5. Обчислення залишків ресурсів на складах</p>
  <p>Підставимо отримані оптимальні значення x1* = 161.54 та x2* = 384.62 у ліві частини рівнянь обмежень:</p>
  <p>Використання сировини: 1.0 &middot; 161.54 + 4.0 &middot; 384.62 = 1700.0 од. Залишок: 1700 - 1700 = 0.0 од. (100% використання).</p>
  <p>Використання машинного часу: 8.0 &middot; 161.54 + 0.8 &middot; 384.62 = 1600.0 год. Залишок: 1600 - 1600 = 0.0 год. (100% використання).</p>
  <p>Висновок: обидва ресурси є повністю вичерпаними (дефіцитними), резервів немає.</p>

  <p>6. Цілочисельний аналіз для неподільних виробів</p>
  <p>Оскільки столи є фізично неподільною продукцією (x1, x2 належать множині цілих чисел Z), проведено аналіз цілочисельних вузлів сітки в околі вершини B(161.54; 384.62) з перевіркою виконання обмежень.</p>

  <div class="fig-container">
    <img src="__IMG2__" alt="Деталізований окіл точки оптимуму та цілочисельний аналіз">
    <div class="fig-caption">Рисунок 2 - Деталізований окіл точки оптимуму та цілочисельний аналіз</div>
  </div>

  <p>Серед усіх допустимих цілочисельних планів найкращим є план (160; 385):</p>
  <p>Випуск: 160 столів моделі А та 385 столів моделі В;</p>
  <p>Прибуток: W = 2.0(160) + 2.4(385) = 320 + 924 = 1244.00 грн;</p>
  <p>Використання сировини: 1.0(160) + 4.0(385) = 1700 од. (залишок 0 од., 100% завантаження);</p>
  <p>Використання часу: 8.0(160) + 0.8(385) = 1588.0 год. (залишок 12.0 год., завантаження 99.25%).</p>

  <p>7. Зведена таблиця результатів та рекомендації виробництву</p>

  <div class="fig-container">
    <img src="__IMG3__" alt="Зведені показники розв'язку оптимізаційної задачі">
    <div class="fig-caption">Рисунок 3 - Зведені показники розв'язку оптимізаційної задачі</div>
  </div>

  <div class="table-title">Таблиця 2 - Порівняльний аналіз отриманих планів виробництва</div>
  <table>
    <tr>
      <th>Показник</th>
      <th>Неперервний розв'язок (x*)</th>
      <th>Практичний цілочисельний план</th>
    </tr>
    <tr>
      <td class="text-left">Столи моделі А (x1), шт.</td>
      <td>161.54</td>
      <td>160</td>
    </tr>
    <tr>
      <td class="text-left">Столи моделі В (x2), шт.</td>
      <td>384.62</td>
      <td>385</td>
    </tr>
    <tr>
      <td class="text-left">Максимальний прибуток (W), грн</td>
      <td>1246.15</td>
      <td>1244.00</td>
    </tr>
    <tr>
      <td class="text-left">Використання сировини</td>
      <td>1700 / 1700 од. (100%)</td>
      <td>1700 / 1700 од. (100%)</td>
    </tr>
    <tr>
      <td class="text-left">Залишок сировини</td>
      <td>0.0 од.</td>
      <td>0.0 од.</td>
    </tr>
    <tr>
      <td class="text-left">Використання машинного часу</td>
      <td>1600 / 1600 год. (100%)</td>
      <td>1588.0 / 1600 год. (99.25%)</td>
    </tr>
    <tr>
      <td class="text-left">Залишок машинного часу</td>
      <td>0.0 год.</td>
      <td>12.0 год.</td>
    </tr>
  </table>

  <p>Рекомендації щодо оптимізації виробництва: оскільки ресурс сировини вичерпаний на 100%, а машинний час задіяний на 99.25%, підприємству слід затвердити щоденний виробничий план у розмірі 160 столів моделі А та 385 столів моделі В. Це забезпечує отримання максимального практичного прибутку у сумі 1244.00 грн за мінімального недовантаження обладнання (лише 12 годин за зміну). Подальше збільшення прибутку можливе виключно за умови додаткового постачання сировини.</p>

  <p>8. Програмна реалізація мовою Python (лістинг розрахунку)</p>
  <div class="code-block">import numpy as np
from scipy.optimize import linprog

N, k = 4, 0.5
S_A, T_A, W_A = 0.5 * N * k, N / k, N * k           # 1.0, 8.0, 2.0
S_B, T_B, W_B = float(N), 0.4 * N * k, 1.2 * N * k     # 4.0, 0.8, 2.4
LIMIT_RAW, LIMIT_TIME = 1700.0, 1600.0

# Розв'язок задачі оптимізації за допомогою SciPy HiGHS
c = [-W_A, -W_B]
A_ub = [[S_A, S_B], [T_A, T_B]]
b_ub = [LIMIT_RAW, LIMIT_TIME]
bounds = [(0, None), (0, None)]

res_cont = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method='highs')
res_int = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, integrality=[1, 1], method='highs')

print(f'Неперервний: x1={res_cont.x[0]:.2f}, x2={res_cont.x[1]:.2f}, W={-res_cont.fun:.2f}')
print(f'Цілочисельний: x1={res_int.x[0]:.0f}, x2={res_int.x[1]:.0f}, W={-res_int.fun:.2f}')</div>

  <p>9. Відповіді на контрольні питання</p>
  <p>1. Які задачі називають задачами лінійного програмування? Відповідь: Оптимізаційні задачі, в яких цільова функція є лінійною, а множина допустимих розв'язків задана системою лінійних рівностей або нерівностей.</p>
  <p>2. Що таке цільова функція? Відповідь: Функція залежності оптимізованого критерію (прибутку, собівартості) від керованих змінних, екстремальне значення якої є метою дослідження.</p>
  <p>3. Як записуються рівняння обмеження? Відповідь: У вигляді алгебраїчних нерівностей або рівнянь sum(c_ij * x_j) <= S_i, де c_ij - питомі норми витрат, а S_i - наявний запас ресурсу.</p>
  <p>4. Які обмеження обов'язково застосовуються до задач оптимального виробництва? Відповідь: Обмеження за обсягами виробничих ресурсів, умови невід'ємності x_j >= 0 та вимоги цілочисельності змінних для неподільної продукції.</p>
  <p>5. Який розв'язок ЗЛП називають оптимальним? Відповідь: План виробництва з області допустимих розв'язків, який надає цільовій функції екстремальне (найбільше або найменше) значення.</p>
  <p>6. Надайте геометричну інтерпретацію ЗЛП. Відповідь: Знаходження точки опуклого багатокутника (ОДР), у якій лінія рівня цільової функції досягає екстремуму в напрямку вектора градієнта.</p>
  <p>7. Яка точка допустимої множини розв'язку називається кутовою? Відповідь: Вершина багатокутника допустимих розв'язків, яка не може бути виражена як опукла комбінація двох інших точок цієї множини.</p>
  <p>8. Поясніть алгоритм графічного методу розв'язання ЗЛП. Відповідь: Побудова прямих обмежень -> побудова ОДР -> обчислення градієнта c -> переміщення перпендикулярної лінії рівня до крайньої точки ОДР -> знаходження координат оптимуму.</p>

  <h2>Висновки</h2>
  <p>1. У процесі виконання роботи опановано методику побудови економіко-математичних моделей та їх розв'язання графічним методом.</p>
  <p>2. За індивідуальними даними Варіанта 4 (N = 4, k = 0.5) складено модель максимізації прибутку: W = 2.0 x1 + 2.4 x2 -> max при обмеженнях на сировину (1700 од.) та машинний час (1600 год.).</p>
  <p>3. Графічним методом знайдено вершину ОДР B(161.54; 384.62), що забезпечує глобальний максимум прибутку 1246.15 грн при повному завантаженні сировини та обладнання.</p>
  <p>4. Сформовано практичний цілочисельний план випуску: 160 столів моделі А та 385 столів моделі В, що дає 1244.00 грн чистого прибутку при нульовому залишку сировини та залишку 12.0 год. машинного часу.</p>
  <p>5. Числові та графічні розрахунки верифіковано в середовищі Python за допомогою бібліотек NumPy та SciPy (метод HiGHS).</p>
</div>

</body>
</html>
"""

html_content = html_template.replace("__IMG1__", img1_b64).replace("__IMG2__", img2_b64).replace("__IMG3__", img3_b64)

html_path = os.path.abspath("report_print.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Створено HTML для друку: {html_path}")

edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
pdf_targets = [
    os.path.abspath(os.path.join(BASE_DIR, "..", "Звіт_ЛР1_Буханцев.pdf"))
]

for pdf_target in pdf_targets:
    cmd = [
        edge_path,
        "--headless=new",
        "--disable-gpu",
        f"--print-to-pdf={pdf_target}",
        "--no-pdf-header-footer",
        html_path
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode == 0 and os.path.exists(pdf_target):
        size_kb = os.path.getsize(pdf_target) / 1024
        print(f"Успішно згенеровано PDF: {pdf_target} ({size_kb:.1f} KB)")
    else:
        print(f"Помилка генерації {pdf_target}: {res.stderr}")

if os.path.exists(html_path):
    os.remove(html_path)
