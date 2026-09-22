"""
Лабораторна робота № 1 з дисципліни «Дослідження операцій»
Тема: Двомірна задача лінійного програмування
Варіант 4 (N = 4, k = 1)
Студент: Буханцев М. В., група КН-24-1
Викладач: Бурдільна Є. В.
"""

import os
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import linprog

# 1. Параметри варіанта 4 (N = 4, k = 1)
N = 4
k = 1

S_A = 0.5 * N * k      # 2.0 од. сировини на 1 стіл А
T_A = N / k            # 4.0 од. часу на 1 стіл А
S_B = float(N)         # 4.0 од. сировини на 1 стіл В
T_B = 0.4 * N * k      # 1.6 од. часу на 1 стіл В
W_A = float(N * k)     # 4.0 грн прибутку від 1 столу А
W_B = 1.2 * N * k      # 4.8 грн прибутку від 1 столу В

LIMIT_RAW = 1700.0     # запас сировини
LIMIT_TIME = 1600.0    # запас машинного часу

print("="*60)
print(f"ПАРАМЕТРИ ЗАВДАННЯ (Варіант 4: N={N}, k={k}):")
print(f"Стіл А: сировина SA={S_A}, час TA={T_A}, прибуток WA={W_A}")
print(f"Стіл В: сировина SB={S_B}, час TB={T_B}, прибуток WB={W_B}")
print(f"Запаси: сировина = {LIMIT_RAW}, машинний час = {LIMIT_TIME}")
print("="*60)

# 2. Точний аналітичний розрахунок вершин ОДР
# ОДР обмежена:
# 1) x1 >= 0, x2 >= 0
# 2) 2*x1 + 4*x2 <= 1700  => при x1=0: x2=425; при x2=0: x1=850
# 3) 4*x1 + 1.6*x2 <= 1600 => при x1=0: x2=1000; при x2=0: x1=400

# Точка перетину обмежень (1) та (2):
# 2*x1 + 4*x2 = 1700
# 4*x1 + 1.6*x2 = 1600
A_mat = np.array([[S_A, S_B], [T_A, T_B]])
b_vec = np.array([LIMIT_RAW, LIMIT_TIME])
intersect_pt = np.linalg.solve(A_mat, b_vec)
x1_opt, x2_opt = intersect_pt[0], intersect_pt[1]
w_max = W_A * x1_opt + W_B * x2_opt

vertices = {
    'O': (0.0, 0.0),
    'A': (LIMIT_TIME / T_A, 0.0),               # (400.0, 0.0)
    'B': (x1_opt, x2_opt),                      # (287.5, 281.25)
    'C': (0.0, LIMIT_RAW / S_B)                 # (0.0, 425.0)
}

print("Вершини області допустимих розв'язків (ОДР):")
for name, (vx, vy) in vertices.items():
    val = W_A * vx + W_B * vy
    print(f"  Вершина {name}({vx:.2f}, {vy:.2f}) -> W = {val:.2f} грн")

print(f"\nНеперервний оптимум: B({x1_opt:.2f}, {x2_opt:.2f}) з W = {w_max:.2f} грн")

# 3. Чисельна перевірка через scipy.optimize.linprog
c = [-W_A, -W_B]
A_ub = [[S_A, S_B], [T_A, T_B]]
b_ub = [LIMIT_RAW, LIMIT_TIME]
bounds = [(0, None), (0, None)]

res_cont = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, method='highs')
res_int = linprog(c, A_ub=A_ub, b_ub=b_ub, bounds=bounds, integrality=[1, 1], method='highs')

print(f"Scipy перевірка (неперервний): x1={res_cont.x[0]:.2f}, x2={res_cont.x[1]:.2f}, W={-res_cont.fun:.2f}")
print(f"Scipy перевірка (цілочисельний): x1={res_int.x[0]:.0f}, x2={res_int.x[1]:.0f}, W={-res_int.fun:.2f}")

# Залишки для неперервного:
raw_left_cont = LIMIT_RAW - (S_A * x1_opt + S_B * x2_opt)
time_left_cont = LIMIT_TIME - (T_A * x1_opt + T_B * x2_opt)

# Залишки для цілочисельного:
int_x1, int_x2 = int(res_int.x[0]), int(res_int.x[1])
raw_left_int = LIMIT_RAW - (S_A * int_x1 + S_B * int_x2)
time_left_int = LIMIT_TIME - (T_A * int_x1 + T_B * int_x2)
w_int = W_A * int_x1 + W_B * int_x2

print(f"Залишки (неперервний): сировина={raw_left_cont:.2f}, машинний час={time_left_cont:.2f}")
print(f"Залишки (цілочисельний): сировина={raw_left_int:.2f}, машинний час={time_left_int:.2f}")

# 4. Створення директорій для збереження графіків
target_dirs = [
    os.path.join("reports", "lab1", "assets"),
    os.path.join("labs", "lab1", "assets")
]
for d in target_dirs:
    os.makedirs(d, exist_ok=True)

# 5. Побудова основного графіка (Рисунок 1)
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
fig, ax = plt.subplots(figsize=(10, 7.5), dpi=300)

x1_vals = np.linspace(0, 900, 1000)

# Прямі обмежень
# 1) 2*x1 + 4*x2 = 1700 => x2 = 425 - 0.5*x1
L1 = (LIMIT_RAW - S_A * x1_vals) / S_B
# 2) 4*x1 + 1.6*x2 = 1600 => x2 = 1000 - 2.5*x1
L2 = (LIMIT_TIME - T_A * x1_vals) / T_B

ax.plot(x1_vals, L1, label=r'$L_1:\ 2x_1 + 4x_2 = 1700$ (Сировина)', color='#1f77b4', linewidth=2)
ax.plot(x1_vals, L2, label=r'$L_2:\ 4x_1 + 1.6x_2 = 1600$ (Машинний час)', color='#d62728', linewidth=2)

# Межі ОДР
poly_x = [0, 400, x1_opt, 0]
poly_y = [0, 0, x2_opt, 425]
ax.fill(poly_x, poly_y, color='#2ca02c', alpha=0.25, label='Область допустимих розв’язків (ОДР)')

# Вершини ОДР
for name, (vx, vy) in vertices.items():
    ax.scatter(vx, vy, color='#0b5394', s=60, zorder=5)
    val = W_A * vx + W_B * vy
    offset = (10, 10)
    if name == 'B':
        offset = (-100, 15)
    elif name == 'A':
        offset = (-20, 12)
    elif name == 'C':
        offset = (15, -5)
    elif name == 'O':
        offset = (10, 10)
    ax.annotate(f"{name}({vx:.1f}; {vy:.1f})\nW={val:.0f}", 
                xy=(vx, vy), xytext=offset, textcoords='offset points',
                fontsize=9.5, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#ffffcc', alpha=0.85, edgecolor='#b3b3b3'))

# Вектор градієнта grad W = (4, 4.8)
grad_scale = 35.0
gx, gy = W_A * grad_scale, W_B * grad_scale
ax.quiver(0, 0, gx, gy, angles='xy', scale_units='xy', scale=1, 
          color='#800080', width=0.007, label=r'Вектор градієнта $\vec{c} = (4.0;\ 4.8)$', zorder=6)
ax.annotate(r'$\vec{c} = \nabla W = (4; 4.8)$', xy=(gx, gy), xytext=(15, -5),
            textcoords='offset points', fontsize=10, fontweight='bold', color='#800080')

# Лінії рівня цільової функції W = 4*x1 + 4.8*x2 = const
w_levels = [800, 1600, 2040, 2500]
for idx, wl in enumerate(w_levels):
    line_level = (wl - W_A * x1_vals) / W_B
    is_opt = (wl == 2500)
    ls = '-' if is_opt else '--'
    lw = 2.2 if is_opt else 1.2
    c_color = '#e65100' if is_opt else '#ff9800'
    label = f'Оптимальна лінія рівня $W = {wl}$' if is_opt else (f'Лінії рівня $W = const$' if idx == 0 else None)
    ax.plot(x1_vals, line_level, linestyle=ls, linewidth=lw, color=c_color, label=label, alpha=0.9)

# Підсвітка точки максимуму
ax.scatter([x1_opt], [x2_opt], color='#d9534f', s=140, edgecolor='black', linewidth=1.5, zorder=7, 
           label=f'Точка максимуму B({x1_opt:.1f}; {x2_opt:.2f})')

ax.set_xlim(-20, 600)
ax.set_ylim(-20, 600)
ax.set_xlabel(r'Кількість столів моделі А ($x_1$), шт.', fontsize=11, fontweight='bold')
ax.set_ylabel(r'Кількість столів моделі В ($x_2$), шт.', fontsize=11, fontweight='bold')
ax.set_title('Графічне розв’язання задачі лінійного програмування (Варіант 4)\n' + 
             r'$W = 4x_1 + 4.8x_2 \to \max,\quad 2x_1 + 4x_2 \leq 1700,\quad 4x_1 + 1.6x_2 \leq 1600$', 
             fontsize=12, fontweight='bold', pad=14)

ax.legend(loc='upper right', frameon=True, framealpha=0.95, facecolor='#ffffff', edgecolor='#cccccc', fontsize=9)
ax.grid(True, linestyle=':', alpha=0.6)

plt.tight_layout()
fig1_path = os.path.join("reports", "lab1", "assets", "graph_lab1.png")
fig.savefig(fig1_path, dpi=300)
fig.savefig(os.path.join("labs", "lab1", "assets", "graph_lab1.png"), dpi=300)
plt.close(fig)
print(f"Збережено графік: {fig1_path}")

# 6. Побудова деталізованого графіка околу точки оптимуму (Рисунок 2)
fig_zoom, ax_z = plt.subplots(figsize=(9, 6.5), dpi=300)

z_x = np.linspace(270, 305, 500)
z_L1 = (LIMIT_RAW - S_A * z_x) / S_B
z_L2 = (LIMIT_TIME - T_A * z_x) / T_B

ax_z.plot(z_x, z_L1, label=r'Межа $L_1$: $2x_1 + 4x_2 = 1700$ (сировина)', color='#1f77b4', linewidth=2.2)
ax_z.plot(z_x, z_L2, label=r'Межа $L_2$: $4x_1 + 1.6x_2 = 1600$ (час)', color='#d62728', linewidth=2.2)

# Заливка ОДР у зумі
z_fill_y = np.minimum(z_L1, z_L2)
ax_z.fill_between(z_x, 260, z_fill_y, color='#2ca02c', alpha=0.2, label='Допустима область')

# Сітка цілих точок в околі
int_x_pts = np.arange(275, 301)
int_y_pts = np.arange(270, 290)
for ix in int_x_pts:
    for iy in int_y_pts:
        is_feas = (S_A * ix + S_B * iy <= LIMIT_RAW) and (T_A * ix + T_B * iy <= LIMIT_TIME)
        if is_feas:
            if ix == int_x1 and iy == int_x2:
                ax_z.scatter(ix, iy, color='#ff9900', s=120, marker='*', edgecolor='black', zorder=6,
                             label=f'Цілочисельний оптимум ({ix}; {iy}), W={w_int:.1f}')
            else:
                ax_z.scatter(ix, iy, color='#2ca02c', s=25, alpha=0.7, zorder=4)
        else:
            ax_z.scatter(ix, iy, color='#999999', s=12, alpha=0.3, zorder=3)

# Неперервний оптимум
ax_z.scatter([x1_opt], [x2_opt], color='#d9534f', s=150, marker='o', edgecolor='black', zorder=7,
             label=f'Неперервний оптимум B({x1_opt:.2f}; {x2_opt:.2f}), W={w_max:.1f}')

# Лінія рівня W = 2500
z_opt_line = (2500 - W_A * z_x) / W_B
ax_z.plot(z_x, z_opt_line, color='#e65100', linestyle='--', linewidth=1.8, label='Лінія рівня W = 2500')

# Лінія рівня для цілочисельного W = 2497.6
z_int_line = (w_int - W_A * z_x) / W_B
ax_z.plot(z_x, z_int_line, color='#ff9900', linestyle=':', linewidth=1.8, label=f'Лінія рівня W = {w_int:.1f}')

ax_z.set_xlim(275, 300)
ax_z.set_ylim(270, 290)
ax_z.set_xlabel(r'Кількість столів моделі А ($x_1$), шт.', fontsize=11, fontweight='bold')
ax_z.set_ylabel(r'Кількість столів моделі В ($x_2$), шт.', fontsize=11, fontweight='bold')
ax_z.set_title('Деталізований окіл точки оптимуму та цілочисельний аналіз\n' + 
               f'Неперервний розв’язок: ({x1_opt}; {x2_opt}) | Цілочисельний розв’язок: ({int_x1}; {int_x2})',
               fontsize=11.5, fontweight='bold', pad=12)

ax_z.legend(loc='lower left', frameon=True, framealpha=0.95, facecolor='#ffffff', edgecolor='#cccccc', fontsize=8.5)
ax_z.grid(True, linestyle=':', alpha=0.6)

plt.tight_layout()
fig2_path = os.path.join("reports", "lab1", "assets", "graph_lab1_zoom.png")
fig_zoom.savefig(fig2_path, dpi=300)
fig_zoom.savefig(os.path.join("labs", "lab1", "assets", "graph_lab1_zoom.png"), dpi=300)
plt.close(fig_zoom)
print(f"Збережено детальний графік: {fig2_path}")

# 7. Створення графічного зведення результатів (Таблиця / Інфографіка)
fig_t, ax_t = plt.subplots(figsize=(10, 6), dpi=300)
ax_t.axis('off')

table_data = [
    ["Параметр / Показник", "Стіл моделі А (x1)", "Стіл моделі В (x2)", "Обмеження / Наявність"],
    ["Витрати сировини (SA, SB)", f"{S_A:.1f} од./шт.", f"{S_B:.1f} од./шт.", f"≤ {LIMIT_RAW:.0f} од."],
    ["Машинний час (TA, TB)", f"{T_A:.1f} год./шт.", f"{T_B:.1f} год./шт.", f"≤ {LIMIT_TIME:.0f} год."],
    ["Прибуток на одиницю (WA, WB)", f"{W_A:.1f} грн", f"{W_B:.1f} грн", "Максимізація прибутку W"],
    ["Неперервний план x*", f"{x1_opt:.2f} шт.", f"{x2_opt:.2f} шт.", f"W_max = {w_max:.2f} грн"],
    ["Використання сировини (неперервний)", f"{S_A * x1_opt:.1f} од.", f"{S_B * x2_opt:.1f} од.", f"1700 / 1700 (залишок: {raw_left_cont:.1f})"],
    ["Використання часу (неперервний)", f"{T_A * x1_opt:.1f} год.", f"{T_B * x2_opt:.1f} год.", f"1600 / 1600 (залишок: {time_left_cont:.1f})"],
    ["Практичний план (цілочисельний)", f"{int_x1} шт.", f"{int_x2} шт.", f"W = {w_int:.1f} грн"],
    ["Використання сировини (цілий план)", f"{S_A * int_x1:.1f} од.", f"{S_B * int_x2:.1f} од.", f"{S_A*int_x1 + S_B*int_x2:.0f} / 1700 (залишок: {raw_left_int:.1f})"],
    ["Використання часу (цілий план)", f"{T_A * int_x1:.1f} год.", f"{T_B * int_x2:.1f} год.", f"{T_A*int_x1 + T_B*int_x2:.1f} / 1600 (залишок: {time_left_int:.1f})"]
]

table = ax_t.table(cellText=table_data, loc='center', cellLoc='center',
                    colWidths=[0.34, 0.20, 0.20, 0.26])
table.auto_set_font_size(False)
table.set_fontsize(10)
table.scale(1, 1.8)

# Стилізація таблиці
for (r, c_idx), cell in table.get_celld().items():
    if r == 0:
        cell.set_facecolor('#003366')
        cell.set_text_props(color='white', fontweight='bold')
    elif r in [4, 7]:
        cell.set_facecolor('#d9edf7')
        cell.set_text_props(fontweight='bold')
    elif r % 2 == 1:
        cell.set_facecolor('#f9f9f9')
    cell.set_edgecolor('#cccccc')

plt.title('Зведені результати розв’язання задачі оптимізації плану виробництва (ЛР № 1, Варіант 4)', 
          fontsize=12, fontweight='bold', pad=15)
plt.tight_layout()
fig3_path = os.path.join("reports", "lab1", "assets", "results_summary.png")
fig_t.savefig(fig3_path, dpi=300)
fig_t.savefig(os.path.join("labs", "lab1", "assets", "results_summary.png"), dpi=300)
plt.close(fig_t)
print(f"Збережено зведення результатів: {fig3_path}")
print("Усі розрахунки та графіки успішно згенеровано!")
