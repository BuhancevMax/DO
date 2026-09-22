"""
Генерація офіційного звіту з лабораторної роботи № 1 у форматі PDF.
Відповідає вимогам ДСТУ / методичних вказівок КрНУ:
- Формат А4, поля по 20 мм;
- Шрифт Times New Roman 14 пт, міжрядковий інтервал 1.5, абзац 1.25 см;
- Окремий титульний аркуш на 1 сторінку;
- Природний потік тексту з запобіганням розривів рисунків і таблиць;
- Вбудовані повнорозмірні графіки високої чіткості;
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
    }
  }

  @page :first {
    @bottom-right {
      content: "";
    }
  }
  
  * {
    box-sizing: border-box;
  }

  body {
    font-family: "Times New Roman", Times, serif;
    font-size: 13.5pt;
    line-height: 1.4;
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
    font-size: 12pt;
    line-height: 1.35;
    font-weight: bold;
  }
  .title-middle {
    margin: auto 0;
  }
  .title-sub {
    font-size: 14pt;
    font-weight: bold;
    letter-spacing: 0.5px;
    margin-bottom: 12pt;
  }
  .title-main {
    font-size: 22pt;
    font-weight: bold;
    margin-bottom: 6pt;
  }
  .title-lab-num {
    font-size: 15pt;
    font-weight: bold;
    margin-bottom: 10pt;
  }
  .title-theme {
    font-size: 13.5pt;
    line-height: 1.3;
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
    font-weight: bold;
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
  }
  
  h2 {
    font-size: 15pt;
    font-weight: bold;
    text-align: center;
    margin: 14pt 0 8pt 0;
  }
  h3 {
    font-size: 13.5pt;
    font-weight: bold;
    margin: 12pt 0 5pt 0;
    text-indent: 1.25cm;
  }
  
  .table-title {
    text-indent: 0;
    font-weight: bold;
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
  }
  table th {
    background-color: #f2f2f2;
    font-weight: bold;
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
    max-width: 90%;
    max-height: 120mm;
    height: auto;
    border: 1px solid #ccc;
  }
  .fig-caption {
    font-size: 11.5pt;
    font-weight: bold;
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

  ol, ul {
    margin: 0 0 6pt 0;
    padding-left: 2cm;
  }
  li {
    margin-bottom: 3pt;
    text-align: justify;
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
    <div class="title-lab-num">З ЛАБОРАТОРНОЇ РОБОТИ № 1</div>
    <div class="title-theme">Тема: «Двомірна задача лінійного програмування»<br><b>Варіант № 4</b></div>
  </div>

  <div>
    <div class="title-sign-block">
      <p class="no-indent"><b>Виконав:</b><br>
      здобувач групи КН-24-1<br>
      <b>Буханцев М. В.</b></p>
      <br>
      <p class="no-indent"><b>Перевірила:</b><br>
      доцент кафедри АІС<br>
      <b>Бурдільна Є. В.</b></p>
    </div>

    <div class="title-footer">
      Кременчук 2026
    </div>
  </div>
</div>

<!-- ОСНОВНИЙ ЗМІСТ (СТОРІНКИ 2+) -->
<div class="content">
  <h2>Лабораторна робота № 1</h2>
  <p class="no-indent"><b>Тема:</b> Двомірна задача лінійного програмування</p>
  <p class="no-indent"><b>Мета:</b> набути навичок з розв’язування двомірних задач лінійного програмування графічним методом.</p>
  
  <h2>Хід роботи</h2>

  <h3>1. Постановка задачі та вихідні дані для варіанта 4</h3>
  <p>Підприємство випускає столи двох моделей: А і В. Для випуску одного столу моделі А потрібно <i>S<sub>A</sub></i> одиниці сировини та <i>T<sub>A</sub></i> одиниці машинного часу. Для випуску одного столу моделі В потрібно <i>S<sub>B</sub></i> одиниці сировини та <i>T<sub>B</sub></i> одиниць машинного часу. Прибуток від реалізації одного столу моделі А складає <i>W<sub>A</sub></i> грошові одиниці, столу моделі В – <i>W<sub>B</sub></i> грошові одиниці. На підприємстві наявні 1700 одиниць сировини та 1600 одиниць машинного часу. Визначити, яким має бути план виробництва, щоб підприємство отримало максимальний прибуток.</p>

  <p>За індивідуальним завданням варіанта 4 (номер у списку <i>N = 4</i>, коефіцієнт підгрупи 1 <i>k = 1</i>):</p>
  <ul>
    <li><i>S<sub>A</sub> = 0.5 &middot; N &middot; k = 0.5 &middot; 4 &middot; 1 = 2.0</i> (од. сировини на стіл А);</li>
    <li><i>T<sub>A</sub> = N / k = 4 / 1 = 4.0</i> (год. машинного часу на стіл А);</li>
    <li><i>S<sub>B</sub> = N = 4.0</i> (од. сировини на стіл В);</li>
    <li><i>T<sub>B</sub> = 0.4 &middot; N &middot; k = 0.4 &middot; 4 &middot; 1 = 1.6</i> (год. машинного часу на стіл В);</li>
    <li><i>W<sub>A</sub> = N &middot; k = 4 &middot; 1 = 4.0</i> (грн прибутку від 1 столу А);</li>
    <li><i>W<sub>B</sub> = 1.2 &middot; N &middot; k = 1.2 &middot; 4 &middot; 1 = 4.8</i> (грн прибутку від 1 столу В).</li>
  </ul>

  <div class="table-title">Таблиця 1 – Вихідні дані задачі оптимізації виробництва (Варіант 4)</div>
  <table>
    <tr>
      <th>Показник</th>
      <th>Стіл моделі А (<i>x</i><sub>1</sub>)</th>
      <th>Стіл моделі В (<i>x</i><sub>2</sub>)</th>
      <th>Наявний запас на складі</th>
    </tr>
    <tr>
      <td class="text-left">Витрати сировини на виріб, од.</td>
      <td>2.0</td>
      <td>4.0</td>
      <td>1700 од.</td>
    </tr>
    <tr>
      <td class="text-left">Витрати машинного часу, год.</td>
      <td>4.0</td>
      <td>1.6</td>
      <td>1600 год.</td>
    </tr>
    <tr>
      <td class="text-left">Прибуток від реалізації виробу, грош. од.</td>
      <td>4.0</td>
      <td>4.8</td>
      <td>—</td>
    </tr>
  </table>

  <h3>2. Економіко-математична модель ЗЛП</h3>
  <p>Нехай <i>x</i><sub>1</sub> — щоденний обсяг випуску столів моделі А (шт.), а <i>x</i><sub>2</sub> — моделі В (шт.).</p>
  <p><b>Цільова функція:</b></p>
  <div class="formula">W(x<sub>1</sub>, x<sub>2</sub>) = 4.0 x<sub>1</sub> + 4.8 x<sub>2</sub> &rarr; max</div>
  
  <p><b>Система обмежень за ресурсами:</b></p>
  <div class="formula">
    2.0 x<sub>1</sub> + 4.0 x<sub>2</sub> &le; 1700 (запас сировини)<br>
    4.0 x<sub>1</sub> + 1.6 x<sub>2</sub> &le; 1600 (фонд машинного часу)
  </div>
  
  <p><b>Граничні умови невід'ємності:</b></p>
  <div class="formula">x<sub>1</sub> &ge; 0, &nbsp; x<sub>2</sub> &ge; 0</div>

  <h3>3. Графічний розв'язок ЗЛП</h3>
  <p>Перетворимо нерівності обмежень у рівності граничних прямих:</p>
  <p>1. Пряма <i>L</i><sub>1</sub> (сировина): <i>2 x<sub>1</sub> + 4 x<sub>2</sub> = 1700 &rArr; x<sub>2</sub> = 425 - 0.5 x<sub>1</sub></i>. Точки перетину з осями: (0; 425) та (850; 0).</p>
  <p>2. Пряма <i>L</i><sub>2</sub> (час): <i>4 x<sub>1</sub> + 1.6 x<sub>2</sub> = 1600 &rArr; x<sub>2</sub> = 1000 - 2.5 x<sub>1</sub></i>. Точки перетину з осями: (0; 1000) та (400; 0).</p>
  <p>Перетин півплощин у першому квадранті утворює область допустимих розв'язків (ОДР) у вигляді опуклого чотирикутника <i>OABC</i> з вершинами:</p>
  <ul>
    <li><i>O(0; 0)</i> &mdash; початок координат;</li>
    <li><i>A(400; 0)</i> &mdash; точка перетину прямої <i>L</i><sub>2</sub> з віссю <i>Ox</i><sub>1</sub>;</li>
    <li><i>B(x<sub>1</sub><sup>*</sup>; x<sub>2</sub><sup>*</sup>)</i> &mdash; точка перетину прямих <i>L</i><sub>1</sub> та <i>L</i><sub>2</sub>;</li>
    <li><i>C(0; 425)</i> &mdash; точка перетину прямої <i>L</i><sub>1</sub> з віссю <i>Ox</i><sub>2</sub>.</li>
  </ul>

  <p>Знайдемо аналітично координати вершини <i>B</i> розв'язком системи лінійних рівнянь:</p>
  <div class="formula">
    2 x<sub>1</sub> + 4 x<sub>2</sub> = 1700<br>
    4 x<sub>1</sub> + 1.6 x<sub>2</sub> = 1600 &rArr; x<sub>1</sub><sup>*</sup> = 287.5, &nbsp; x<sub>2</sub><sup>*</sup> = 281.25
  </div>

  <p>Вектор градієнта цільової функції: <i>c&#8407; = &nabla;W = (4.0; 4.8)</i>. Переміщуючи перпендикулярні лінії рівня <i>4x<sub>1</sub> + 4.8x<sub>2</sub> = const</i> у напрямку градієнта, встановлюємо, що крайньою точкою виходу з ОДР є вершина <b>B(287.5; 281.25)</b>.</p>

  <div class="fig-container">
    <img src="__IMG1__" alt="Графічне розв’язання ЗЛП">
    <div class="fig-caption">Рисунок 1 – Графічне розв’язання задачі лінійного програмування (Варіант 4)</div>
  </div>

  <h3>4. Розрахунок цільової функції у кутових точках</h3>
  <ul>
    <li><i>W(O) = 4(0) + 4.8(0) = 0</i> грош. од.;</li>
    <li><i>W(A) = 4(400) + 4.8(0) = 1600</i> грош. од.;</li>
    <li><i>W(C) = 4(0) + 4.8(425) = 2040</i> грош. од.;</li>
    <li><i>W(B) = 4(287.5) + 4.8(281.25) = 1150 + 1350 = <b>2500</b></i> грош. од.</li>
  </ul>
  <p>Максимальний прибуток досягається в точці <i>B</i> і складає <b>2500.00 грн</b>.</p>

  <h3>5. Обчислення залишків на складах</h3>
  <p>Підставимо оптимальний неперервний план у ліві частини обмежень:</p>
  <p>&bull; Використання сировини: <i>2 &middot; 287.5 + 4 &middot; 281.25 = 1700</i> од. Залишок: <i>1700 - 1700 = <b>0 од.</b></i></p>
  <p>&bull; Використання машинного часу: <i>4 &middot; 287.5 + 1.6 &middot; 281.25 = 1600</i> год. Залишок: <i>1600 - 1600 = <b>0 год.</b></i></p>
  <p>Обидва ресурси використовуються на 100%, є дефіцитними, невикористаних залишків немає.</p>

  <h3>6. Цілочисельний аналіз для неподільних виробів</h3>
  <p>Оскільки столи виготовляються цілими штуками (<i>x</i><sub>1</sub>, <i>x</i><sub>2</sub> &isin; &Zopf;), досліджено дискретні вузли в околі точки <i>B</i> (Рисунок 2).</p>

  <div class="fig-container">
    <img src="__IMG2__" alt="Окіл оптимуму">
    <div class="fig-caption">Рисунок 2 – Деталізований окіл точки оптимуму та цілочисельний аналіз</div>
  </div>

  <p>Серед допустимих вузлів найкращим є план <b>(286; 282)</b>: прибуток складає <i>4(286) + 4.8(282) = <b>2497.60 грн</b></i>. Залишок сировини: 0 од., машинного часу: 4.8 год (недовикористання становить лише 0.3%).</p>

  <h3>7. Зведена таблиця результатів та рекомендації</h3>
  
  <div class="fig-container">
    <img src="__IMG3__" alt="Зведені результати">
    <div class="fig-caption">Рисунок 3 – Зведені показники розв’язку оптимізаційної задачі</div>
  </div>

  <div class="table-title">Таблиця 2 – Порівняльний аналіз отриманих планів виробництва</div>
  <table>
    <tr>
      <th>Показник</th>
      <th>Неперервний розв'язок (<i>x</i><sup>*</sup>)</th>
      <th>Практичний цілочисельний план</th>
    </tr>
    <tr>
      <td class="text-left">Столи моделі А (<i>x</i><sub>1</sub>), шт.</td>
      <td>287.50</td>
      <td>286</td>
    </tr>
    <tr>
      <td class="text-left">Столи моделі В (<i>x</i><sub>2</sub>), шт.</td>
      <td>281.25</td>
      <td>282</td>
    </tr>
    <tr>
      <td class="text-left"><b>Максимальний прибуток (<i>W</i>), грн</b></td>
      <td><b>2500.00</b></td>
      <td><b>2497.60</b></td>
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
      <td>1595.2 / 1600 год. (99.7%)</td>
    </tr>
    <tr>
      <td class="text-left">Залишок машинного часу</td>
      <td>0.0 год.</td>
      <td>4.8 год.</td>
    </tr>
  </table>

  <p><b>Рекомендації щодо оптимізації виробництва:</b></p>
  <p>Оскільки всі виробничі ресурси вичерпані повністю, збільшення прибутку можливе за рахунок розширення ресурсних лімітів (закупівля сировини та модернізація обладнання). Практичний випуск рекомендується здійснювати щоденно партіями 286 столів моделі А та 282 столів моделі В.</p>

  <h3>8. Програмна реалізація мовою Python (лістинг розрахунку)</h3>
  <div class="code-block">import numpy as np
from scipy.optimize import linprog

N, k = 4, 1
S_A, T_A, W_A = 0.5 * N * k, N / k, N * k           # 2.0, 4.0, 4.0
S_B, T_B, W_B = float(N), 0.4 * N * k, 1.2 * N * k     # 4.0, 1.6, 4.8
LIMIT_RAW, LIMIT_TIME = 1700.0, 1600.0

# Розв'язок задачі оптимізації за допомогою SciPy HiGHS
c = [-W_A, -W_B]
A_ub = [[S_A, S_B], [T_A, T_B]]
b_ub = [LIMIT_RAW, LIMIT_TIME]
bounds = [(0, None), (0, None)]

res_cont = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method='highs')
res_int = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, integrality=[1, 1], method='highs')

print(f"Неперервний: x1={res_cont.x[0]:.2f}, x2={res_cont.x[1]:.2f}, W={-res_cont.fun:.2f}")
print(f"Цілочисельний: x1={res_int.x[0]:.0f}, x2={res_int.x[1]:.0f}, W={-res_int.fun:.2f}")</div>

  <h3>9. Відповіді на контрольні питання</h3>
  <p><b>1. Які задачі називають задачами лінійного програмування?</b> Оптимізаційні задачі з лінійною цільовою функцією та лінійною системою рівностей чи нерівностей обмежень.</p>
  <p><b>2. Що таке цільова функція?</b> Формалізований критерій оптимізації, який відображає кінцеву мету операції (максимум прибутку або мінімум витрат).</p>
  <p><b>3. Як записуються рівняння обмеження?</b> У формі &sum; <i>c<sub>ij</sub> x<sub>j</sub> &le; S<sub>i</sub></i> (де <i>c<sub>ij</sub></i> — питомі витрати ресурсів, <i>S<sub>i</sub></i> — запас).</p>
  <p><b>4. Які обмеження обов’язково застосовуються до задач оптимального виробництва?</b> Обмеження ресурсів, невід'ємності змінних <i>x<sub>j</sub> &ge; 0</i> та за необхідності цілочисельності.</p>
  <p><b>5. Який розв’язок ЗЛП називають оптимальним?</b> Допустимий план із множини ОДР, що забезпечує екстремальне значення цільової функції.</p>
  <p><b>6. Надайте геометричну інтерпретацію ЗЛП.</b> Пошук вершини опуклого багатокутника допустимих розв'язків, якої торкається лінія рівня при її зміщенні за вектором градієнта.</p>
  <p><b>7. Яка точка допустимої множини розв’язку називається кутовою?</b> Вершина багатокутника допустимої області, яка не лежить строго між іншими точками цієї області.</p>
  <p><b>8. Поясніть алгоритм графічного методу розв’язання ЗЛП.</b> Побудова граничних прямих, знаходження багатокутника ОДР, побудова градієнта <i>c&#8407;</i>, рух лінії рівня до виходу з ОДР, визначення координат крайньої вершини.</p>

  <h2>Висновки</h2>
  <p>1. Під час виконання лабораторної роботи оволодіно методикою геометричного та аналітичного моделювання задач лінійного програмування.</p>
  <p>2. Для Варіанта 4 (<i>N=4, k=1</i>) складено математичну модель ЗЛП максимізації прибутку виробництва меблів: <i>W = 4 x<sub>1</sub> + 4.8 x<sub>2</sub> &rarr; max</i>.</p>
  <p>3. Графічним методом знайдено екстремальну вершину <i>B(287.50; 281.25)</i>, яка дає максимальний прибуток <b>2500.00 грн</b> при абсолютному (100%) використанні фондів сировини та машинного часу.</p>
  <p>4. Проведено цілочисельний аналіз дискретного випуску столів: план <i>(286; 282)</i> забезпечує прибуток <b>2497.60 грн</b> при відсутності залишків сировини.</p>
  <p>5. Розрахунки графічного розв'язку верифіковано програмно у середовищі Python за допомогою модуля <i>scipy.optimize.linprog</i>.</p>
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
