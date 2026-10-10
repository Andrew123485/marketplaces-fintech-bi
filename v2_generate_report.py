import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings('ignore')

FILL_DARK  = PatternFill("solid", fgColor="0070C0")
FILL_LIGHT = PatternFill("solid", fgColor="BDD7EE")
FILL_GREEN = PatternFill("solid", fgColor="C3D69B")

def bold(size=9, white=False): return Font(bold=True, size=size, name="Calibri", color="FFFFFF" if white else "000000")
def reg(size=9): return Font(bold=False, size=size, name="Calibri")

WRAP   = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT   = Alignment(wrap_text=True, vertical="center", horizontal="left")
CENTER = Alignment(horizontal="center", vertical="center")

LETTERS = "abcdefghijklmnopqrstuvwxyz"

def z_test_sig(p1, n1, p2, n2, alpha=0.05):
    if n1 < 5 or n2 < 5 or (p1 == 0 and p2 == 0): return False
    p_pool = (p1 * n1 + p2 * n2) / (n1 + n2)
    if p_pool in (0, 1): return False
    se = (p_pool * (1 - p_pool) * (1/n1 + 1/n2)) ** 0.5
    if se == 0: return False
    from scipy.stats import norm
    z = (p1 - p2) / se
    return norm.cdf(z) > (1 - alpha / 2)

def apply_ztest(values, bases, compare_groups, alpha=0.05):
    n = len(values)
    winners = [[] for _ in range(n)]
    for grp in compare_groups:
        for pos_i, gi in enumerate(grp):
            if gi >= n: continue
            for pos_j, gj in enumerate(grp):
                if gi == gj or gj >= n: continue
                vi, vj = values[gi], values[gj]
                ni, nj = bases[gi],  bases[gj]
                if vi is not None and vj is not None and ni > 0 and nj > 0:
                    if z_test_sig(vi, ni, vj, nj, alpha):
                        winners[gi].append(LETTERS[pos_j])
    result = []
    for i, val in enumerate(values):
        if val is None: result.append(None)
        elif winners[i]:
            lets = "".join(sorted(set(winners[i])))
            result.append(f"{round(val*100)}% {lets}")
        else:
            result.append(val)
    return result

COMPARE_GROUPS = [
    [1, 2, 3], [4, 5], [6, 7], [8, 9], [10, 11], [12, 13], [14, 15], [16, 17], [18, 19], [20, 21],
]

Q2_ANSWERS = {
    "display": ["Для совершения покупок со скидкой только на маркетплейсе", "Для совершения повседневных покупок вне маркетплейса", "В этом банке моя основная карта, с которой я совершаю повседневные покупки вне маркетплейса", "Для открытия вклада/накопительного счёта", "Для хранения сбережений на вкладе/накопительном счёте", "Другое"],
    "WB": ["Для совершения покупок со скидкой только на Wildberries", None, "В этом банке моя основная карта, с которой я совершаю повседневные покупки вне Wildberries", "Для открытия вклада/накопительного счета в этом банке", "Для хранения сбережений на вкладе/накопительном счете", "Другое"],
    "Ozon": ["Для совершения покупок со скидкой только на Ozon", None, "В этом банке моя основная карта, с которой я совершаю повседневные покупки вне Ozon", "Для открытия вклада/накопительного счета в этом банке", "Для хранения сбережений на вкладе/накопительном счете", "Другое"],
    "Яндекс": ["Для совершения покупок со скидкой только на Яндекс.Маркете и других сервисах Яндекса", None, "В этом банке моя основная карта, с которой я совершаю повседневные покупки вне Яндекс.Маркета и других сервисов Яндекса", "Для открытия вклада/накопительного счета в этом банке", "Для хранения сбережений на вкладе/накопительном счете", "Другое"],
}

Q3_ANSWERS = {
    "display": ["Выгодный кешбэк/бонусы на покупки на маркетплейсе", "Выгодный кешбэк/бонусы на покупки вне маркетплейса", "Удобно, что финансы и покупки внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри маркетплейса", "Привлекла рекламная акция", "Привлекли условия накопительного счёта", "Другое (укажите, что именно?)", "Затрудняюсь ответить"],
    "WB": ["Выгодный кешбэк/бонусы на покупки на Wildberries", "Выгодный кешбэк/бонусы на покупки вне Wildberries", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Wildberries (Частями)", "Привлекла рекламная акция", None, "Другое", "Затрудняюсь ответить"],
    "Ozon": ["Выгодный кешбэк/бонусы на покупки на Ozon", "Выгодный кешбэк/бонусы на покупки вне Ozon", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Ozon (Ozon Рассрочка)", "Привлекла рекламная акция", "Привлекли условия накопительного счета с начислением на ежедневный остаток", "Другое", "Затрудняюсь ответить"],
    "Яндекс": ["Выгодный кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса", "Выгодный кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисах Яндекса", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Яндекс.Маркета и других сервисов Яндекса (Сплит)", "Привлекла рекламная акция", "Привлекли условия накопительного счета с начислением на ежедневный остаток", "Другое", "Затрудняюсь ответить"],
}

Q5_ANSWERS = {"display": ["Полностью соответствует", "Скорее соответствует", "Отчасти соответствует, отчасти нет", "Скорее не соответствует", "Полностью не соответствует", "Затрудняюсь ответить"], "all": ["Полностью соответствует", "Скорее соответствует", "Отчасти соответствует, отчасти нет", "Скорее не соответствует", "Полностью не соответствует", "Затрудняюсь ответить"]}
Q6_ANSWERS = {
    "display": ["Кешбэк/бонусы на покупки на маркетплейсе ниже обещанного", "Кешбэк/бонусы на покупки вне маркетплейса ниже обещанного", "Неудобное приложение для совершения финансовых операций", "Отсутствие обновлений в приложении, финансовых продуктах", "«Фейковые», обманчивые скидки", "Некачественная техподдержка пользователей", "Неинтересные предложения по вкладам, накопительным счетам", "Нет нужных функций в приложении", "Другое (укажите, что именно?)", "Затрудняюсь ответить"],
    "WB": ["Кешбэк/бонусы на покупки на Wildberries ниже обещанного", "Кешбэк/бонусы на покупки вне Wildberries ниже обещанного", "Неудобное приложение для совершения финансовых операций", "Отсутствие обновлений в приложении, финансовых продуктах", "«Фейковые», обманчивые скидки", "Некачественная техподдержка пользователей", "Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам", "Нет нужных функций в приложении", "Другое", "Затрудняюсь ответить"],
    "Ozon": ["Кешбэк/бонусы на покупки на Ozon ниже обещанного", "Кешбэк/бонусы на покупки вне Ozon ниже обещанного", "Неудобное приложение для совершения финансовых операций", "Отсутствие обновлений в приложении, финансовых продуктах", "«Фейковые», обманчивые скидки", "Некачественная техподдержка пользователей", "Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам", "Нет нужных функций в приложении", "Другое", "Затрудняюсь ответить"],
    "Яндекс": ["Кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса ниже обещанного", "Кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисах Яндекса ниже обещанного", "Неудобное приложение для совершения финансовых операций", "Отсутствие обновлений в приложении, финансовых продуктах", "«Фейковые», обманчивые скидки", "Некачественная техподдержка пользователей", "Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам", "Нет нужных функций в приложении", "Другое", "Затрудняюсь ответить"],
}
Q7_ANSWERS = {"display": ["Банк-маркетплейс – удобный дополнительный функционал для покупок на маркетплейсе", "Я не хочу заводить регулярные платежи (ЖКХ, связь и т.д.) в банке-маркетплейсе", "Банк-маркетплейс не «дотягивает» до уровня традиционных банков", "Банк-маркетплейс выделяется среди других банков-маркетплейсов выгодными условиями", "Банк-маркетплейс обещает большой кешбэк, но не выполняет обещания", "Банк-маркетплейс не развивается как банк", "Продукты и условия банка-маркетплейса ухудшаются со временем", "Не согласен(-на) ни с одним из утверждений"], "WB": ["WB Банк/ WB кошелек – это удобный дополнительный функционал для совершения покупок на Wildberries", "Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в WB Банк/ WB кошелек", "WB Банк/ WB кошелек не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов", "WB Банк/ WB кошелек выделяется среди других банков-маркетплейсов выгодными условиями", "WB Банк/ WB кошелек обещает большой кешбэк, но не выполняет свои обещания", "WB Банк/ WB кошелек не развивается как банк", "Продукты и условия WB Банка/ WB кошелька ухудшаются со временем", "Не согласен(-на) ни с одним из утверждений"], "Ozon": ["Ozon Банк – это удобный дополнительный функционал для совершения покупок на Ozon", "Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в Ozon банк", "Ozon Банк не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов", "Ozon Банк выделяется среди других банков-маркетплейсов выгодными условиями на банковские продукты", "Ozon Банк обещает большой кешбэк, но не выполняет свои обещания", "Ozon Банк не развивается как банк", "Продукты и условия Ozon банк ухудшаются со временем", "Не согласен(-на) ни с одним из утверждений"], "Яндекс": ["Яндекс.Банк – это удобный дополнительный функционал для совершения покупок в экосистеме Яндекс", "Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в Яндекс.Банке", "Яндекс.Банк не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов", "Яндекс.Банк выделяется среди других банков-маркетплейсов выгодными условиями на банковские продукты", "Яндекс.Банк обещает большой кешбэк, но не выполняет свои обещания", "Яндекс.Банк не развивается как банк", "Продукты и условия Яндекс.Банка ухудшаются со временем", "Не согласен(-на) ни с одним из утверждений"]}
Q8_ANSWERS = {"display": ["Планирую пользоваться чаще", "Планирую пользоваться так же часто как и сейчас", "Планирую пользоваться реже или прекратить пользоваться"], "all": ["Планирую пользоваться чаще", "Планирую пользоваться так же часто как и сейчас", "Планирую пользоваться реже или прекратить пользоваться"]}
Q9_ANSWERS = {"display": ["Выгодный кешбэк/бонусы на покупки на маркетплейсе", "Выгодный кешбэк/бонусы на покупки вне маркетплейса", "Удобно, что финансы и покупки внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри маркетплейса", "Другое (укажите, что именно?)", "Затрудняюсь ответить"], "WB": ["Выгодный кешбэк/бонусы на покупки на Wildberries", "Выгодный кешбэк/бонусы на покупки вне Wildberries", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Wildberries (Частями)", "Другое", None], "Ozon": ["Выгодный кешбэк/бонусы на покупки на Ozon", "Выгодный кешбэк/бонусы на покупки вне Ozon", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Ozon (Ozon Рассрочка)", "Другое", None], "Яндекс": ["Выгодный кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса", "Выгодный кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисах Яндекса", "Удобно, что финансы и покупки находятся внутри одной системы", "Доступна мгновенная рассрочка на покупку внутри Яндекс.Маркета и других сервисов Яндекса (Сплит)", "Другое", None]}

BANK_COLS = {
    "WB": {"q1": "0. Какими из банков ниже Вы пользуетесь?", "q2": "1. С какой целью вы сейчас пользуетесь WB Банком/ WB кошельком?", "q3": "2. Почему вы решили стать клиентом WB Банка/ WB кошелька?", "q5": "4. Насколько WB Банк/ WB кошелек соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?", "q6": "5. Почему WB Банк/ WB кошелек не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?", "q7": "6. Выберите утверждения о WB Банке/ WB кошельке, с которыми вы согласны", "q8": "7. Как вы планируете пользоваться услугами WB Банка/ WB кошелька для повседневных покупок в ближайшие 12 месяцев?", "q9": "8. Почему вы будете дальше пользоваться WB Банком/ WB кошельком?"},
    "Ozon": {"q1": "0. Какими из банков ниже Вы пользуетесь?", "q2": "1. С какой целью вы сейчас пользуетесь Ozon Банком?", "q3": "2. Почему вы решили стать клиентом Ozon Банка?", "q5": "4. Насколько Ozon Банк соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?", "q6": "5. Почему Ozon Банк не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?", "q7": "6. Выберите утверждения об Ozon Банке, с которыми вы согласны", "q8": "7. Как вы планируете пользоваться услугами Ozon Банка для повседневных покупок в ближайшие 12 месяцев?", "q9": "8. Почему вы будете дальше пользоваться Ozon Банком?"},
    "Яндекс": {"q1": "0. Какими из банков ниже Вы пользуетесь?", "q2": "1. С какой целью вы сейчас пользуетесь Яндекс.Банком?", "q3": "2. Почему вы решили стать клиентом Яндекс.Банка?", "q5": "4. Насколько Яндекс.Банк соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?", "q6": "5. Почему Яндекс.Банк не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?", "q7": "6. Выберите утверждения о Яндекс.Банке, с которыми вы согласны", "q8": "7. Как вы планируете пользоваться услугами Яндекс.Банка для повседневных покупок в ближайшие 12 месяцев?", "q9": "8. Почему вы будете дальше пользоваться Яндекс.Банком?"},
}

BANK_LABELS = {"Wildberries (Вайлдберриз Банк)": "WB", "Озон Банк (Ozon)": "Ozon", "Яндекс": "Яндекс"}

def load_data(filepath):
    df = pd.read_excel(filepath, sheet_name="База_ИТОГ")
    if "Анкета полная" in df.columns: df = df[df["Анкета полная"] == "Да"].copy()
    df["bank_key"] = df["bank_target"].map(BANK_LABELS)
    return df

def contains(series, value):
    if value is None: return pd.Series([False]*len(series), index=series.index)
    if isinstance(series, pd.DataFrame): series = series.iloc[:, 0]
    return series.astype(str).apply(lambda x: any(p.strip() == value for p in str(x).split("|")))

def build_segs(df):
    bW, bO, bY = df["bank_key"]=="WB", df["bank_key"]=="Ozon", df["bank_key"]=="Яндекс"
    sg = df["Кол-во переводов (группы)"]
    sh = df["Доля переводов от поступлений (группы)"]
    sm = df["Сумма переводов по группам"]
    return [("В целом", None, None), ("WB", bW, "WB"), ("Ozon", bO, "Ozon"), ("Яндекс", bY, "Яндекс"),
            ("WB_3", bW&(sg=="3-28 переводов"), "WB"), ("WB_29+", bW&(sg=="29 переводов и более"), "WB"),
            ("Ozon_3", bO&(sg=="3-28 переводов"), "Ozon"), ("Ozon_29+", bO&(sg=="29 переводов и более"), "Ozon"),
            ("Яндекс_3", bY&(sg=="3-28 переводов"), "Яндекс"), ("Яндекс_29+", bY&(sg=="29 переводов и более"), "Яндекс"),
            ("WB_0", bW&(sh=="0-30%"), "WB"), ("WB_31", bW&(sh=="31-100%"), "WB"),
            ("Ozon_0", bO&(sh=="0-30%"), "Ozon"), ("Ozon_31", bO&(sh=="31-100%"), "Ozon"),
            ("Яндекс_0", bY&(sh=="0-30%"), "Яндекс"), ("Яндекс_31", bY&(sh=="31-100%"), "Яндекс"),
            ("WB_l", bW&(sm=="до 10000 руб"), "WB"), ("WB_g", bW&(sm=="больше 10 000 руб"), "WB"),
            ("Ozon_l", bO&(sm=="до 10000 руб"), "Ozon"), ("Ozon_g", bO&(sm=="больше 10 000 руб"), "Ozon"),
            ("Яндекс_l", bY&(sm=="до 10000 руб"), "Яндекс"), ("Яндекс_g", bY&(sm=="больше 10 000 руб"), "Яндекс")]

def write_title(ws, text, n_cols):
    c = ws.cell(row=1, column=1, value=text)
    c.fill = FILL_DARK; c.font = bold(11, white=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols+2)
    ws.row_dimensions[1].height = 20

def set_col_widths(ws, n_segs):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 54
    ws.column_dimensions["C"].width = 17
    for i in range(1, n_segs+50):
        ws.column_dimensions[get_column_letter(i+3)].width = 13

SEG_GROUPS = [("В ЦЕЛОМ", [0]), ("Банк-маркетплейс", [1, 2, 3]), ("Количество переводов за 6 месяцев", [4, 5, 6, 7, 8, 9]), ("Доля переводов от поступлений", [10, 11, 12, 13, 14, 15]), ("Сумма переводов (группы)", [16, 17, 18, 19, 20, 21])]
SEG_SUB_LABELS = {0: "В\nЦЕЛОМ", 1: "WB Банк/\nWB кошелек\n(a)", 2: "Ozon.Банк\n(b)", 3: "ЯндексБанк\n(c)", 4: "WB:\n3-28 пер.\n(a)", 5: "WB:\n29+\n(b)", 6: "Ozon:\n3-28 пер.\n(a)", 7: "Ozon:\n29+\n(b)", 8: "Яндекс:\n3-28 пер.\n(a)", 9: "Яндекс:\n29+\n(b)", 10: "WB:\n0-30%\n(a)", 11: "WB:\n31-100%\n(b)", 12: "Ozon:\n0-30%\n(a)", 13: "Ozon:\n31-100%\n(b)", 14: "Яндекс:\n0-30%\n(a)", 15: "Яндекс:\n31-100%\n(b)", 16: "WB:\nдо 10 тыс\n(a)", 17: "WB:\n>10 тыс\n(b)", 18: "Ozon:\nдо 10 тыс\n(a)", 19: "Ozon:\n>10 тыс\n(b)", 20: "Яндекс:\nдо 10 тыс\n(a)", 21: "Яндекс:\n>10 тыс\n(b)"}

def write_q_block(ws, start_row, q_label, table_rows, n_segs):
    r = start_row
    ws.cell(row=r, column=2, value=q_label).font = bold(9)
    r += 1
    for grp_name, idxs in SEG_GROUPS:
        valid = [i for i in idxs if i < n_segs]
        if not valid: continue
        col_start, col_end = valid[0] + 3, valid[-1] + 3
        c = ws.cell(row=r, column=col_start, value=grp_name)
        c.fill, c.font, c.alignment = FILL_LIGHT, bold(8), WRAP
        if col_end > col_start: ws.merge_cells(start_row=r, start_column=col_start, end_row=r, end_column=col_end)
    r += 1
    for ci in range(n_segs):
        c = ws.cell(row=r, column=ci+3, value=SEG_SUB_LABELS.get(ci, ""))
        c.fill, c.font, c.alignment = FILL_LIGHT, bold(8), WRAP
    ws.row_dimensions[r].height = 45; r += 1
    for dr in table_rows:
        label, vals = dr[0], dr[1:]
        cell_b = ws.cell(row=r, column=2, value=label)
        cell_b.font, cell_b.alignment = reg(9), LEFT
        if label != "База": cell_b.fill = FILL_LIGHT
        for ci, val in enumerate(vals[:n_segs]):
            is_sig = isinstance(val, str) and "%" in val
            c = ws.cell(row=r, column=ci+3, value=val if pd.notna(val) else 0 if label != "База" else "")
            c.alignment = CENTER; c.font = bold(9) if is_sig else reg(9)
            if is_sig: c.fill = FILL_GREEN
            if label != "База" and pd.notna(val) and not is_sig and isinstance(val, (int, float)):
                c.number_format = '0%'
        r += 1
    return r + 1

def calc_q(df, segs, qkey, ans_dict, alpha=0.05):
    all_vals, all_ns = [], []
    for i, _ in enumerate(ans_dict["display"]):
        vals, ns = [], []
        for _, mask, bk in segs:
            sub = df if mask is None else df[mask]
            
            if qkey == "q1":
                c = BANK_COLS[bk if bk else "WB"]["q1"]
                s2 = sub.dropna(subset=[c]) if c in sub.columns else sub
                av = ans_dict["display"][i]
                v = {"Ozon Банк":"Ozon Банк","WB Банк/ WB кошелек":"WB Банк/ WB кошелек","Яндекс.Банк":"Яндекс.Банк"}.get(av, av)
                if not len(s2): vals.append(None); ns.append(0); continue
                vals.append(contains(s2[c], v).sum()/len(s2)); ns.append(len(s2))
                continue
                
            if bk is None:
                all_v = []
                for bk2 in ["WB", "Ozon", "Яндекс"]:
                    c2 = BANK_COLS[bk2].get(qkey)
                    ak = "all" if "all" in ans_dict else bk2
                    av = ans_dict.get(ak, ans_dict.get("WB"))[i]
                    if c2 and c2 in df.columns:
                        s2 = df[df["bank_key"] == bk2].dropna(subset=[c2])
                        if len(s2) and av: all_v.append(contains(s2[c2], av).sum() / len(s2))
                vals.append(np.mean(all_v) if all_v else None); ns.append(len(df.dropna(subset=[BANK_COLS["WB"]["q1"]])))
                continue
            
            ak = "all" if "all" in ans_dict else bk
            c, av = BANK_COLS.get(bk, {}).get(qkey), ans_dict.get(ak, [])[i] if ak in ans_dict else None
            s2 = sub.dropna(subset=[c]) if c and c in df.columns else sub
            if not c or c not in df.columns or not len(s2) or not av: vals.append(None); ns.append(0); continue
            vals.append(contains(s2[c], av).sum()/len(s2)); ns.append(len(s2))
        all_vals.append(vals); all_ns.append(ns)
    bn = all_ns[0] if all_ns else [0]*len(segs)
    res = []
    for i, disp in enumerate(ans_dict["display"]):
        res.append([disp] + apply_ztest(all_vals[i], bn, COMPARE_GROUPS, alpha))
    if qkey == "q5":
        t2, b2 = [0,1], [3,4]
        res.append(["Top-2 (полностью+скорее соответствует)"] + apply_ztest([sum(all_vals[ti][ci] for ti in t2 if pd.notna(all_vals[ti][ci])) if any(pd.notna(all_vals[ti][ci]) for ti in t2) else None for ci in range(len(segs))], bn, COMPARE_GROUPS, alpha))
        res.append(["Bottom-2 (полностью+скорее не соответствует)"] + apply_ztest([sum(all_vals[ti][ci] for ti in t2 if pd.notna(all_vals[ti][ci])) if any(pd.notna(all_vals[ti][ci]) for ti in t2) else None for ci in range(len(segs))], bn, COMPARE_GROUPS, alpha)) # actually bot2 is [3,4]
        res[-1][1:] = apply_ztest([sum(all_vals[ti][ci] for ti in b2 if pd.notna(all_vals[ti][ci])) if any(pd.notna(all_vals[ti][ci]) for ti in b2) else None for ci in range(len(segs))], bn, COMPARE_GROUPS, alpha)
    res.append(["База"] + bn)
    return res

def calc_cross_table(df, q_cols, dict_cols, q_rows, dict_rows, alpha=0.05):
    res_rows, res_bases = [], []
    col_labels, col_banks = [], []
    for bk in ["WB", "Ozon", "Яндекс"]:
        sub = df[df["bank_key"] == bk]
        ak = "all" if dict_cols and "all" in dict_cols else bk
        disp_cols = dict_cols["display"] if dict_cols else []
        col_c = BANK_COLS[bk].get(q_cols)
        if q_cols == "q5_topbot":
            for name, ids in [("Top-2", [0, 1]), ("Bottom-2", [3, 4])]:
                ak5 = "all" if "all" in Q5_ANSWERS else bk
                avs = [Q5_ANSWERS[ak5][i] for i in ids]
                c_mask = sub[col_c].astype(str).apply(lambda x: any(v in x for v in avs if v)) if col_c in df.columns else sub["bank_key"] != bk
                col_labels.append(f"{name}"); col_banks.append((bk, c_mask))
        else:
            for i, d in enumerate(disp_cols):
                av = dict_cols.get(ak, [])[i]
                if not av: continue
                c_mask = contains(sub[col_c], av) if col_c in df.columns else sub["bank_key"] != bk
                col_labels.append(f"{d[:25]}... ({LETTERS[len([b for b, _ in col_banks if b == bk])]})")
                col_banks.append((bk, c_mask))
    
    # bases is calculated over df[df["bank_key"] == bk]
    bases = []
    for bk, c_mask in col_banks:
        sub_df = df[df["bank_key"] == bk]
        bases.append(len(sub_df[c_mask]))
        
    for r_i, r_disp in enumerate(dict_rows["display"]):
        r_vals = []
        for bk, c_mask in col_banks:
            sub_df = df[df["bank_key"] == bk]
            ak = "all" if "all" in dict_rows else bk
            r_av = dict_rows.get(ak, [])[r_i]
            r_col = BANK_COLS[bk].get(q_rows)
            grp = sub_df[c_mask]
            if not len(grp) or not r_av or r_col not in df.columns: r_vals.append(None)
            else: r_vals.append(contains(grp[r_col], r_av).sum() / len(grp))
        res_rows.append([r_disp] + r_vals)
    
    labeled_rows = []
    # Dynamic compare groups: all items of same bank compared
    comp_grps = []
    last_bk = None; curr_g = []
    for ci, (bk, _) in enumerate(col_banks):
        if bk != last_bk:
            if curr_g: comp_grps.append(curr_g)
            curr_g = [ci]; last_bk = bk
        else: curr_g.append(ci)
    if curr_g: comp_grps.append(curr_g)
    
    for r_i, r_disp in enumerate(dict_rows["display"]):
        labeled_rows.append([r_disp] + apply_ztest(res_rows[r_i][1:], bases, comp_grps, alpha))
        
    if q_rows == "q5":
        for name, ids in [("Top-2 (полностью+скорее соответствует)", [0,1]), ("Bottom-2 (полностью+скорее не)", [3,4])]:
            t_vals = [sum(res_rows[ti][ci+1] for ti in ids if pd.notna(res_rows[ti][ci+1])) if any(pd.notna(res_rows[ti][ci+1]) for ti in ids) else None for ci in range(len(col_banks))]
            labeled_rows.append([name] + apply_ztest(t_vals, bases, comp_grps, alpha))
            
    labeled_rows.append(["База"] + bases)
    return col_labels, col_banks, labeled_rows

def write_cross_block(ws, start_row, q_label, col_labels, col_banks, labeled_rows):
    r = start_row
    ws.cell(row=r, column=2, value=q_label).font = bold(9)
    r += 1
    # Bank groups
    curr_b, cnt = None, 0
    start_c = 3
    for ci, (bk, _) in enumerate(col_banks):
        ws.cell(row=r+1, column=ci+3, value=col_labels[ci]).fill = FILL_LIGHT
        ws.cell(row=r+1, column=ci+3).font = bold(8)
        ws.cell(row=r+1, column=ci+3).alignment = WRAP
        if bk != curr_b:
            if curr_b:
                ws.cell(row=r, column=start_c, value=curr_b).fill = FILL_LIGHT
                ws.cell(row=r, column=start_c).font, ws.cell(row=r, column=start_c).alignment = bold(9), CENTER
                if ci+2 > start_c: ws.merge_cells(start_row=r, start_column=start_c, end_row=r, end_column=ci+2)
            curr_b, start_c = bk, ci+3
    if curr_b:
        ws.cell(row=r, column=start_c, value=curr_b).fill = FILL_LIGHT
        ws.cell(row=r, column=start_c).font, ws.cell(row=r, column=start_c).alignment = bold(9), CENTER
        if len(col_banks)+2 > start_c: ws.merge_cells(start_row=r, start_column=start_c, end_row=r, end_column=len(col_banks)+2)
    ws.row_dimensions[r+1].height = 60
    r += 2
    for dr in labeled_rows:
        label, vals = dr[0], dr[1:]
        cell_b = ws.cell(row=r, column=2, value=label)
        cell_b.font, cell_b.alignment = reg(9), LEFT
        if label != "База": cell_b.fill = FILL_LIGHT
        for ci, val in enumerate(vals):
            is_sig = isinstance(val, str) and "%" in val
            c = ws.cell(row=r, column=ci+3, value=val if pd.notna(val) else 0 if label != "База" else "")
            c.alignment = CENTER; c.font = bold(9) if is_sig else reg(9)
            if is_sig: c.fill = FILL_GREEN
            if label != "База" and pd.notna(val) and not is_sig and isinstance(val, (int, float)):
                c.number_format = '0%'
        r += 1
    return r + 1

def build_results(ws, df, alpha=0.05):
    segs = build_segs(df)
    n = len(segs)
    set_col_widths(ws, n)
    write_title(ws, "Сводные таблицы", n)
    row = 3
    q1_dict = {"display": ["Ozon Банк", "WB Банк/ WB кошелек", "Яндекс.Банк", "Никакими из перечисленных"]}
    row = write_q_block(ws, row, "Q1. Какими из банков ниже Вы пользуетесь?", calc_q(df, segs, "q1", q1_dict, alpha), n)
    row = write_q_block(ws, row, "Q2. С какой целью вы сейчас пользуетесь банком-маркетплейсом?", calc_q(df, segs, "q2", Q2_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q3. Почему вы решили стать клиентом?", calc_q(df, segs, "q3", Q3_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q5. Насколько банк-маркетплейс соответствует ожиданиям?", calc_q(df, segs, "q5", Q5_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q6. Почему банк-маркетплейс не вполне соответствует ожиданиям?", calc_q(df, segs, "q6", Q6_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q7. Выберите утверждения о банке-маркетплейсе, с которыми вы согласны", calc_q(df, segs, "q7", Q7_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q8. Как вы планируете пользоваться услугами банка-маркетплейса?", calc_q(df, segs, "q8", Q8_ANSWERS, alpha), n)
    row = write_q_block(ws, row, "Q9. Почему вы будете дальше пользоваться банком-маркетплейсом?", calc_q(df, segs, "q9", Q9_ANSWERS, alpha), n)

    # Added Crosstabs 
    cb1, cb2, cb3 = calc_cross_table(df, "q3", Q3_ANSWERS, "q5", Q5_ANSWERS, alpha)
    row = write_cross_block(ws, row, "Q3 x Q5 Насколько банк-маркетплейс соответствует ожиданиям", cb1, cb2, cb3)
    cb1, cb2, cb3 = calc_cross_table(df, "q3", Q3_ANSWERS, "q6", Q6_ANSWERS, alpha)
    row = write_cross_block(ws, row, "Q3 x Q6 Почему банк-маркетплейс не вполне соответствует ожиданиям", cb1, cb2, cb3)
    cb1, cb2, cb3 = calc_cross_table(df, "q5_topbot", None, "q8", Q8_ANSWERS, alpha)
    row = write_cross_block(ws, row, "Q5 (Top/Bottom-2) x Q8 Как вы планируете пользоваться", cb1, cb2, cb3)
    cb1, cb2, cb3 = calc_cross_table(df, "q3", Q3_ANSWERS, "q8", Q8_ANSWERS, alpha)
    row = write_cross_block(ws, row, "Q3 x Q8 Как вы планируете пользоваться", cb1, cb2, cb3)


def build_quotes(ws, df):
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 80
    write_title(ws, "Цитаты Другое", 2)
    ws.cell(row=3, column=1, value="—").font = reg(9)
    ws.cell(row=3, column=2, value="Замечание: открытые тексты 'Другое' отсутствуют в файле База_ИТОГ").font = reg(9)

def build_legend(ws):
    ws.column_dimensions["B"].width = 70
    c = ws.cell(row=2, column=2, value="Условные обозначения: как читать таблицы")
    c.font = bold(11); c.alignment = LEFT
    for r, txt in enumerate(["Значения в таблицах — доли (0.75 = 75%). Значения < 0.01 могут отображаться как 0.", "Зелёный фон + буква (напр. 84% ac) — значение значимо выше столбцов a и c.", "WB Банк/ WB кошелек (a), Ozon.Банк (b) — буквенная кодировка столбцов.", "В каждой секции сравнение идёт внутри одного банка.", "База — количество анкет.", "Анкета полная = Да — учитываются только полностью заполненные анкеты."]):
        ws.cell(row=r+4, column=2, value=txt).font = reg(9)

def generate_report(input_path, output_path, alpha=0.05):
    print("Загружаем данные...")
    df = load_data(input_path)
    print(f"Строк после фильтра: {len(df)}")
    wb = Workbook()
    ws1 = wb.active; ws1.title = "Результаты"
    build_results(ws1, df, alpha)
    ws2 = wb.create_sheet("Цитаты Другое")
    build_quotes(ws2, df)
    ws3 = wb.create_sheet("Усл. обозначения")
    build_legend(ws3)
    wb.save(output_path)
    print(f"ГОТОВО! Файл сохранён: {output_path}")
    return output_path
