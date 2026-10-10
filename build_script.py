
# Writes the full generate_report.py in parts
parts = []

parts.append('''
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import warnings
warnings.filterwarnings("ignore")

FILL_DARK  = PatternFill("solid", fgColor="0070C0")
FILL_LIGHT = PatternFill("solid", fgColor="BDD7EE")
FILL_GREEN = PatternFill("solid", fgColor="C3D69B")
FILL_NONE  = PatternFill(fill_type=None)
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def B(sz=9, wh=False): return Font(bold=True, size=sz, name="Calibri", color="FFFFFF" if wh else "000000")
def R(sz=9):           return Font(bold=False, size=sz, name="Calibri")
WRAP   = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT   = Alignment(wrap_text=True, vertical="center", horizontal="left")
CENTER = Alignment(horizontal="center", vertical="center")

LETTERS = "abcdefghijklmnopqrstuvwxyz"
BANK_LABELS = {"Wildberries (Вайлдберриз Банк)":"WB","Озон Банк (Ozon)":"Ozon","Яндекс":"Яндекс"}

def z_test(p1,n1,p2,n2,alpha=0.05):
    if min(n1,n2)<5 or (p1==0 and p2==0): return False
    pp=(p1*n1+p2*n2)/(n1+n2)
    if pp in (0,1): return False
    se=(pp*(1-pp)*(1/n1+1/n2))**0.5
    if se==0: return False
    from scipy.stats import norm
    return norm.cdf((p1-p2)/se)>(1-alpha/2)

def apply_z(vals, bases, grps, alpha=0.05):
    n=len(vals); W=[[] for _ in range(n)]
    for g in grps:
        for pi,gi in enumerate(g):
            if gi>=n: continue
            for pj,gj in enumerate(g):
                if gi==gj or gj>=n: continue
                vi,vj,ni,nj=vals[gi],vals[gj],bases[gi],bases[gj]
                if all(x is not None for x in [vi,vj]) and ni>0 and nj>0:
                    if z_test(vi,ni,vj,nj,alpha): W[gi].append(LETTERS[pj])
    res=[]
    for i,v in enumerate(vals):
        if v is None: res.append("-")
        elif W[i]: res.append(f"{round(v*100)}% {''.join(sorted(set(W[i])))}")
        else: res.append(v)
    return res

CG_MAIN = [[1,2,3],[4,5],[6,7],[8,9],[10,11],[12,13],[14,15],[16,17],[18,19],[20,21]]

BANK_COLS = {
    "WB":{
        "q0":"0. Какими из банков ниже Вы пользуетесь?",
        "q1":"1. С какой целью вы сейчас пользуетесь WB Банком/ WB кошельком?",
        "q2":"2. Почему вы решили стать клиентом WB Банка/ WB кошелька?",
        "q2cat":"Объед. Категория",
        "q4":"4. Насколько WB Банк/ WB кошелек соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?",
        "q5":"5. Почему WB Банк/ WB кошелек не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?",
        "q6":"6. Выберите утверждения о WB Банке/ WB кошельке, с которыми вы согласны",
        "q7":"7. Как вы планируете пользоваться услугами WB Банка/ WB кошелька для повседневных покупок в ближайшие 12 месяцев?",
        "q8":"8. Почему вы будете дальше пользоваться WB Банком/ WB кошельком?",
    },
    "Ozon":{
        "q0":"0. Какими из банков ниже Вы пользуетесь?",
        "q1":"1. С какой целью вы сейчас пользуетесь Ozon Банком?",
        "q2":"2. Почему вы решили стать клиентом Ozon Банка?",
        "q2cat":"Объед. Категория2",
        "q4":"4. Насколько Ozon Банк соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?",
        "q5":"5. Почему Ozon Банк не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?",
        "q6":"6. Выберите утверждения об Ozon Банке, с которыми вы согласны",
        "q7":"7. Как вы планируете пользоваться услугами Ozon Банка для повседневных покупок в ближайшие 12 месяцев?",
        "q8":"8. Почему вы будете дальше пользоваться Ozon Банком?",
    },
    "Яндекс":{
        "q0":"0. Какими из банков ниже Вы пользуетесь?",
        "q1":"1. С какой целью вы сейчас пользуетесь Яндекс.Банком?",
        "q2":"2. Почему вы решили стать клиентом Яндекс.Банка?",
        "q2cat":"Объед.Категория",
        "q4":"4. Насколько Яндекс.Банк соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?",
        "q5":"5. Почему Яндекс.Банк не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?",
        "q6":"6. Выберите утверждения о Яндекс.Банке, с которыми вы согласны",
        "q7":"7. Как вы планируете пользоваться услугами Яндекс.Банка для повседневных покупок в ближайшие 12 месяцев?",
        "q8":"8. Почему вы будете дальше пользоваться Яндекс.Банком?",
    },
}

# Answers per bank
Q0_DISPLAY = ["Ozon Банк","WB Банк/ WB кошелек","Яндекс.Банк","Никакими из перечисленных"]

Q1_ANS = {
    "WB":["Для совершения покупок со скидкой только на Wildberries","Для открытия вклада/накопительного счета в этом банке","Для хранения сбережений на вкладе/накопительном счете","Другое"],
    "Ozon":["Для совершения покупок со скидкой только на Ozon","Для открытия вклада/накопительного счета в этом банке","Для хранения сбережений на вкладе/накопительном счете","Другое"],
    "Яндекс":["Для совершения покупок со скидкой только на Яндекс.Маркете и других сервисах Яндекса","Для открытия вклада/накопительного счета в этом банке","Для хранения сбережений на вкладе/накопительном счете","Другое"],
}
Q1_DISPLAY = ["Для совершения покупок со скидкой только на маркетплейсе","Для открытия вклада/накопительного счёта в банке","Для хранения сбережений на вкладе/накопительном счёте","Другое"]

Q2_ANS = {
    "WB":["Выгодный кешбэк/бонусы на покупки на Wildberries","Выгодный кешбэк/бонусы на покупки вне Wildberries","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Wildberries (Частями)","Привлекла рекламная акция",None,"Другое","Затрудняюсь ответить"],
    "Ozon":["Выгодный кешбэк/бонусы на покупки на Ozon","Выгодный кешбэк/бонусы на покупки вне Ozon","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Ozon (Ozon Рассрочка)","Привлекла рекламная акция","Привлекли условия накопительного счета с начислением на ежедневный остаток","Другое","Затрудняюсь ответить"],
    "Яндекс":["Выгодный кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса","Выгодный кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисов Яндекса","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Яндекс.Маркета и других сервисов Яндекса (Сплит)","Привлекла рекламная акция","Привлекли условия накопительного счета с начислением на ежедневный остаток","Другое","Затрудняюсь ответить"],
}
Q2_DISPLAY = ["Выгодный кешбэк/бонусы на покупки на маркетплейсе","Выгодный кешбэк/бонусы на покупки вне маркетплейса","Удобно, что финансы и покупки внутри одной системы","Доступна мгновенная рассрочка на покупку внутри маркетплейса","Привлекла рекламная акция","Привлекли условия накопительного счёта с начислением на ежедневный остаток","Другое (укажите, что именно?)","Затрудняюсь ответить"]

Q4_ANS = ["Полностью соответствует","Скорее соответствует","Отчасти соответствует, отчасти нет","Скорее не соответствует","Полностью не соответствует","Затрудняюсь ответить"]

Q5_ANS = {
    "WB":["Кешбэк/бонусы на покупки на Wildberries ниже обещанного","Кешбэк/бонусы на покупки вне Wildberries ниже обещанного","Неудобное приложение для совершения финансовых операций","Отсутствие обновлений в приложении, финансовых продуктах","«Фейковые», обманчивые скидки","Некачественная техподдержка пользователей","Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам","Нет нужных функций в приложении","Другое","Затрудняюсь ответить"],
    "Ozon":["Кешбэк/бонусы на покупки на Ozon ниже обещанного","Кешбэк/бонусы на покупки вне Ozon ниже обещанного","Неудобное приложение для совершения финансовых операций","Отсутствие обновлений в приложении, финансовых продуктах","«Фейковые», обманчивые скидки","Некачественная техподдержка пользователей","Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам","Нет нужных функций в приложении","Другое","Затрудняюсь ответить"],
    "Яндекс":["Кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса ниже обещанного","Кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисов Яндекса ниже обещанного","Неудобное приложение для совершения финансовых операций","Отсутствие обновлений в приложении, финансовых продуктах","«Фейковые», обманчивые скидки","Некачественная техподдержка пользователей","Неинтересные предложения по вкладам, накопительным счетам, кредитным картам и другим продуктам","Нет нужных функций в приложении","Другое","Затрудняюсь ответить"],
}
Q5_DISPLAY = ["Кешбэк/бонусы на покупки на маркетплейсе ниже обещанного","Кешбэк/бонусы на покупки вне маркетплейса ниже обещанного","Неудобное приложение для совершения финансовых операций","Отсутствие обновлений в приложении, финансовых продуктах","«Фейковые», обманчивые скидки","Некачественная техподдержка пользователей","Неинтересные предложения по вкладам, накопительным счетам","Нет нужных функций в приложении","Другое (укажите, что именно?)","Затрудняюсь ответить"]

Q6_ANS = {
    "WB":["WB Банк/ WB кошелек – это удобный дополнительный функционал для совершения покупок на Wildberries","Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в WB Банк/ WB кошелек","WB Банк/ WB кошелек не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов","WB Банк/ WB кошелек выделяется среди других банков-маркетплейсов выгодными условиями","WB Банк/ WB кошелек обещает большой кешбэк, но не выполняет свои обещания","WB Банк/ WB кошелек не развивается как банк","Продукты и условия WB Банка/ WB кошелька ухудшаются со временем","Не согласен(-на) ни с одним из утверждений"],
    "Ozon":["Ozon Банк – это удобный дополнительный функционал для совершения покупок на Ozon","Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в Ozon банк","Ozon Банк не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов","Ozon Банк выделяется среди других банков-маркетплейсов выгодными условиями на банковские продукты","Ozon Банк обещает большой кешбэк, но не выполняет свои обещания","Ozon Банк не развивается как банк","Продукты и условия Ozon банк ухудшаются со временем","Не согласен(-на) ни с одним из утверждений"],
    "Яндекс":["Яндекс.Банк – это удобный дополнительный функционал для совершения покупок в экосистеме Яндекс","Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в Яндекс.Банке","Яндекс.Банк не «дотягивает» до уровня традиционных банков из-за ограниченного количества финансовых продуктов","Яндекс.Банк выделяется среди других банков-маркетплейсов выгодными условиями на банковские продукты","Яндекс.Банк обещает большой кешбэк, но не выполняет свои обещания","Яндекс.Банк не развивается как банк","Продукты и условия Яндекс.Банка ухудшаются со временем","Не согласен(-на) ни с одним из утверждений"],
}
Q6_DISPLAY = ["Банк-маркетплейс  – это удобный дополнительный функционал для совершения покупок","Я не хочу заводить регулярные платежи (ЖКХ, мобильная связь и т.д.) в банке-маркетплейсе","Банк-маркетплейс не «дотягивает» до уровня традиционных банков","Банк-маркетплейс выделяется среди других банков-маркетплейсов выгодными условиями","Банк-маркетплейс обещает большой кешбэк, но не выполняет свои обещания","Банк-маркетплейс не развивается как банк","Продукты и условия банка-маркетплейса ухудшаются со временем","Не согласен(-на) ни с одним из утверждений"]

Q7_ANS = ["Планирую пользоваться чаще","Планирую пользоваться так же часто как и сейчас","Планирую пользоваться реже или прекратить пользоваться"]

Q8_ANS = {
    "WB":["Выгодный кешбэк/бонусы на покупки на Wildberries","Выгодный кешбэк/бонусы на покупки вне Wildberries","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Wildberries (Частями)","Другое"],
    "Ozon":["Выгодный кешбэк/бонусы на покупки на Ozon","Выгодный кешбэк/бонусы на покупки вне Ozon","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Ozon (Ozon Рассрочка)","Другое"],
    "Яндекс":["Выгодный кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса","Выгодный кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисов Яндекса","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Яндекс.Маркета и других сервисов Яндекса (Сплит)","Другое"],
}
Q8_DISPLAY = ["Выгодный кешбэк/бонусы на покупки на маркетплейсе","Выгодный кешбэк/бонусы на покупки вне маркетплейса","Удобно, что финансы и покупки внутри одной системы","Доступна мгновенная рассрочка на покупку внутри маркетплейса","Другое (укажите, что именно?)"]

Q2_CAT_DISPLAY = ["Выгодный кешбэк/бонусы  на покупки на маркетплейсе","Выгодный кешбэк/бонусы на покупки вне маркетплейса","Удобно, что финансы и покупки внутри одной системы","Доступна мгновенная рассрочка на покупку внутри маркетплейса","Акция/Реклама"]
Q2_CAT_ANS = {
    "WB":["Выгодный кешбэк/бонусы на покупки на Wildberries","Выгодный кешбэк/бонусы на покупки вне Wildberries","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Wildberries (Частями)","Привлекла рекламная акция"],
    "Ozon":["Выгодный кешбэк/бонусы на покупки на Ozon","Выгодный кешбэк/бонусы на покупки вне Ozon","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Ozon (Ozon Рассрочка)","Привлекла рекламная акция"],
    "Яндекс":["Выгодный кешбэк/бонусы на покупки на Яндекс.Маркете и других сервисах Яндекса","Выгодный кешбэк/бонусы на покупки вне Яндекс.Маркета и других сервисов Яндекса","Удобно, что финансы и покупки находятся внутри одной системы","Доступна мгновенная рассрочка на покупку внутри Яндекс.Маркета и других сервисов Яндекса (Сплит)","Привлекла рекламная акция"],
}

def load_df(fp):
    df = pd.read_excel(fp, sheet_name="База_ИТОГ")
    if "Анкета полная" in df.columns: df=df[df["Анкета полная"]=="Да"].copy()
    df["bank_key"]=df["bank_target"].map(BANK_LABELS)
    return df

def hits(col, val):
    if val is None: return pd.Series([False]*len(col), index=col.index)
    return col.fillna("").astype(str).apply(lambda x: any(p.strip()==val for p in x.split("|")) if x else False)

def build_segs(df):
    bW=df["bank_key"]=="WB"; bO=df["bank_key"]=="Ozon"; bY=df["bank_key"]=="Яндекс"
    sg=df["Кол-во переводов (группы)"]; sh=df["Доля переводов от поступлений (группы)"]; sm=df["Сумма переводов по группам"]
    return [("В целом",None,None),("WB",bW,"WB"),("Ozon",bO,"Ozon"),("Яндекс",bY,"Яндекс"),
            ("WB_3",bW&(sg=="3-28 переводов"),"WB"),("WB_29",bW&(sg=="29 переводов и более"),"WB"),
            ("Oz_3",bO&(sg=="3-28 переводов"),"Ozon"),("Oz_29",bO&(sg=="29 переводов и более"),"Ozon"),
            ("Yn_3",bY&(sg=="3-28 переводов"),"Яндекс"),("Yn_29",bY&(sg=="29 переводов и более"),"Яндекс"),
            ("WB_030",bW&(sh=="0-30%"),"WB"),("WB_31",bW&(sh=="31-100%"),"WB"),
            ("Oz_030",bO&(sh=="0-30%"),"Ozon"),("Oz_31",bO&(sh=="31-100%"),"Ozon"),
            ("Yn_030",bY&(sh=="0-30%"),"Яндекс"),("Yn_31",bY&(sh=="31-100%"),"Яндекс"),
            ("WB_l",bW&(sm=="до 10000 руб"),"WB"),("WB_g",bW&(sm=="больше 10 000 руб"),"WB"),
            ("Oz_l",bO&(sm=="до 10000 руб"),"Ozon"),("Oz_g",bO&(sm=="больше 10 000 руб"),"Ozon"),
            ("Yn_l",bY&(sm=="до 10000 руб"),"Яндекс"),("Yn_g",bY&(sm=="больше 10 000 руб"),"Яндекс")]

SGLBL={0:"В\\nЦЕЛОМ",1:"WB Банк/\\nWB кошелек\\n\\n (a)",2:"Ozon.Банк\\n\\n (b)",3:"ЯндексБанк\\n\\n (c)",
       4:"WB:\\n3-28 пер.,% (a)",5:"WB:\\n29+,% (b)",6:"Ozon:\\n3-28,% (a)",7:"Ozon:\\n29+,% (b)",
       8:"Яндекс:\\n3-28,% (a)",9:"Яндекс:\\n29+,% (b)",
       10:"WB:\\n0-30% (a)",11:"WB:\\n31-100% (b)",12:"Ozon:\\n0-30% (a)",13:"Ozon:\\n31-100% (b)",
       14:"Яндекс:\\n0-30% (a)",15:"Яндекс:\\n31-100% (b)",
       16:"WB:\\nдо 10 тыс. руб (a)",17:"WB:\\n> 10 тыс. руб (b)",
       18:"Ozon:\\nдо 10 тыс. руб (a)",19:"Ozon:\\n> 10 тыс. руб (b)",
       20:"Яндекс:\\nдо 10 тыс. руб (a)",21:"Яндекс:\\n> 10 тыс. руб (b)"}

SGGRPS=[("В ЦЕЛОМ",[0]),("Банк-маркетплейс",[1,2,3]),
        ("Количество переводов за 6 месяцев",[4,5,6,7,8,9]),
        ("Доля переводов от поступлений (группы)",[10,11,12,13,14,15]),
        ("Сумма переводов (группы)",[16,17,18,19,20,21])]
''')

parts.append('''
def calc_row_generic(df, segs, bk_col_map, ans_map, disp_list, ans_idx, pre_filter=None, alpha=0.05):
    """Calculate one answer row across all segs. bk_col_map: bank->col. ans_map: bank->list of answers."""
    vals, ns = [], []
    for _, mask, bk in segs:
        sub = df if mask is None else df[mask]
        if pre_filter is not None:
            sub = sub[pre_filter] if mask is None else sub[pre_filter[sub.index]]
        if bk is None:
            # "В целом": weighted mean across all banks
            all_v = []
            for bk2 in ["WB","Ozon","Яндекс"]:
                col2 = bk_col_map.get(bk2)
                av2 = (ans_map.get(bk2) or [])[ans_idx] if bk2 in ans_map else None
                if col2 and col2 in df.columns:
                    s2 = df[df["bank_key"]==bk2].dropna(subset=[col2])
                    if len(s2) and av2: all_v.append((hits(s2[col2],av2).sum(), len(s2)))
            if all_v:
                tot_h=sum(h for h,_ in all_v); tot_n=sum(n for _,n in all_v)
                vals.append(tot_h/tot_n if tot_n else None); ns.append(tot_n)
            else: vals.append(None); ns.append(0)
            continue
        col = bk_col_map.get(bk)
        av  = (ans_map.get(bk) or [])[ans_idx] if bk in ans_map else None
        if not col or col not in df.columns: vals.append(None); ns.append(0); continue
        s2 = sub.dropna(subset=[col])
        n = len(s2)
        if n==0: vals.append(None); ns.append(0); continue
        if av is None: vals.append(None); ns.append(n); continue
        vals.append(hits(s2[col], av).sum()/n); ns.append(n)
    return vals, ns

def calc_table(df, segs, bk_col_map, ans_map, disp_list, cg=None, alpha=0.05,
               pre_filter=None, add_top2=False, top2_idx=(0,1), bot2_idx=(3,4)):
    if cg is None: cg=CG_MAIN
    all_v, all_n = [], []
    for i in range(len(disp_list)):
        v,n = calc_row_generic(df,segs,bk_col_map,ans_map,disp_list,i,pre_filter,alpha)
        all_v.append(v); all_n.append(n)
    bn = all_n[0] if all_n else [0]*len(segs)
    rows = []
    for i,d in enumerate(disp_list):
        rows.append([d]+apply_z(all_v[i],bn,cg,alpha))
    if add_top2:
        def combo(idxs):
            return [sum(all_v[ti][ci] for ti in idxs if all_v[ti][ci] is not None)
                    if any(all_v[ti][ci] is not None for ti in idxs) else None
                    for ci in range(len(segs))]
        rows.append(["Top-2 (полностью соответсвует+скорее соответствует)"]+apply_z(combo(top2_idx),bn,cg,alpha))
        rows.append(["Bottom-2  (полностью не соответствует+скорее не соответствует)"]+apply_z(combo(bot2_idx),bn,cg,alpha))
    rows.append(["База"]+bn)
    return rows

def calc_q0(df, segs, alpha=0.05):
    """Q0 - which banks do you use"""
    col="0. Какими из банков ниже Вы пользуетесь?"
    vals_all=[]; ns_all=[]
    for ans in Q0_DISPLAY:
        vals,ns=[],[]
        for _,mask,bk in segs:
            sub=df if mask is None else df[mask]
            s2=sub.dropna(subset=[col]) if col in df.columns else sub
            n=len(s2)
            if n==0: vals.append(None); ns.append(0); continue
            vals.append(hits(s2[col],ans).sum()/n); ns.append(n)
        vals_all.append(vals); ns_all.append(ns)
    bn=ns_all[0]
    rows=[]
    for i,d in enumerate(Q0_DISPLAY):
        rows.append([d]+apply_z(vals_all[i],bn,CG_MAIN,alpha))
    rows.append(["База"]+bn)
    return rows

def write_title_row(ws, text, ncols):
    c=ws.cell(1,1,text); c.fill=FILL_DARK; c.font=B(11,True); c.alignment=LEFT
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncols+2)
    ws.row_dimensions[1].height=20

def set_widths(ws, n):
    ws.column_dimensions["A"].width=3
    ws.column_dimensions["B"].width=52
    ws.column_dimensions["C"].width=16
    for i in range(1,n): ws.column_dimensions[get_column_letter(i+3)].width=13

def write_header(ws, row, n, extra_label=None):
    # Row with group labels
    r=row
    for gname, idxs in SGGRPS:
        valid=[i for i in idxs if i<n]
        if not valid: continue
        cs,ce=valid[0]+3,valid[-1]+3
        c=ws.cell(r,cs,gname); c.fill=FILL_LIGHT; c.font=B(8); c.alignment=WRAP
        if ce>cs: ws.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=ce)
    r+=1
    # Row with column sub-labels
    for ci in range(n):
        lbl=SGLBL.get(ci,"")
        c=ws.cell(r,ci+3,lbl); c.fill=FILL_LIGHT; c.font=B(8); c.alignment=WRAP
    ws.row_dimensions[r].height=50
    return r+1

def write_q_block(ws, row, title, table_rows, n, qnote=None):
    # Title
    c=ws.cell(row,2,title); c.font=B(9); c.alignment=LEFT
    if qnote:
        ws.cell(row,3,qnote).font=R(8)
    row+=1
    row=write_header(ws,row,n)
    for dr in table_rows:
        lbl=dr[0]; vals=dr[1:]
        is_base=(lbl=="База")
        cb=ws.cell(row,2,lbl); cb.font=R(9); cb.alignment=LEFT
        if not is_base: cb.fill=FILL_LIGHT
        for ci,v in enumerate(vals[:n]):
            is_sig=isinstance(v,str) and "%" in v
            disp=v if is_sig else (int(v) if is_base and isinstance(v,(int,float,np.integer)) else v)
            if not is_base and isinstance(v,(int,float,np.floating)) and not is_sig:
                pct=round(v*100); disp=f"{pct}%" if pct>0 else "-"
            if disp is None: disp="-"
            cell=ws.cell(row,ci+3,disp); cell.alignment=CENTER
            cell.font=B(9) if is_sig else R(9)
            if is_sig: cell.fill=FILL_GREEN
        row+=1
    return row+1
''')

parts.append('''
def build_results(ws, df, alpha=0.05):
    segs=build_segs(df); n=len(segs)
    set_widths(ws,n); write_title_row(ws,"Сводные таблицы",n)
    row=3

    # Q0
    row=write_q_block(ws,row,"Q1.  Какими из банков ниже Вы пользуетесь?",
                      calc_q0(df,segs,alpha),n)

    # Q1
    row=write_q_block(ws,row,"Q2.  С какой целью вы сейчас пользуетесь банком-маркетплейсом?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q1"] for bk in["WB","Ozon","Яндекс"]},
                          Q1_ANS, Q1_DISPLAY, alpha=alpha),n)

    # Q2
    row=write_q_block(ws,row,"Q3.  Почему вы решили стать клиентом банка-маркетплейса?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q2"] for bk in["WB","Ozon","Яндекс"]},
                          Q2_ANS, Q2_DISPLAY, alpha=alpha),n)

    # Q2 укрупненная
    row=write_q_block(ws,row,"Q3.  Почему вы решили стать клиентом банка-маркетплейса? (укрупненная категория)",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q2"] for bk in["WB","Ozon","Яндекс"]},
                          Q2_CAT_ANS, Q2_CAT_DISPLAY, alpha=alpha),n)

    # Q4 - satisfaction
    row=write_q_block(ws,row,"Q4.  Насколько банк-маркетплейс соответствует ожиданиям, которые у вас были, когда вы стали клиентом банка?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q4"] for bk in["WB","Ozon","Яндекс"]},
                          {bk:Q4_ANS for bk in["WB","Ozon","Яндекс"]}, Q4_ANS,
                          alpha=alpha, add_top2=True, top2_idx=(0,1), bot2_idx=(3,4)),n)

    # Q5 - reasons for dissatisfaction. Base = those who answered q4 as partial/not-satisfy
    def make_q5_prefilter(bk):
        col=BANK_COLS[bk]["q4"]
        if col not in df.columns: return pd.Series([False]*len(df),index=df.index)
        return df["bank_key"].eq(bk) & df[col].isin(["Скорее не соответствует","Полностью не соответствует","Отчасти соответствует, отчасти нет"])
    q5_mask = make_q5_prefilter("WB")|make_q5_prefilter("Ozon")|make_q5_prefilter("Яндекс")
    row=write_q_block(ws,row,"Q5.  Почему банк-маркетплейс не вполне соответствует ожиданиям, которые были у вас, когда вы стали клиентом банка?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q5"] for bk in["WB","Ozon","Яндекс"]},
                          Q5_ANS, Q5_DISPLAY, alpha=alpha, pre_filter=q5_mask),n)

    # Q6 - statements
    row=write_q_block(ws,row,"Q7.  Выберите утверждения о банке-маркетплейсе, с которыми вы согласны",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q6"] for bk in["WB","Ozon","Яндекс"]},
                          Q6_ANS, Q6_DISPLAY, alpha=alpha),n)

    # Q7 - planning
    row=write_q_block(ws,row,"Q8.  Как вы планируете пользоваться услугами банка-маркетплейса для повседневных покупок в ближайшие 12 месяцев?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q7"] for bk in["WB","Ozon","Яндекс"]},
                          {bk:Q7_ANS for bk in["WB","Ozon","Яндекс"]},Q7_ANS,alpha=alpha),n)

    # Q8 - why continue (base = planning to continue from q7)
    def make_q8_prefilter(bk):
        col=BANK_COLS[bk]["q7"]
        if col not in df.columns: return pd.Series([False]*len(df),index=df.index)
        return df["bank_key"].eq(bk) & df[col].isin(["Планирую пользоваться чаще","Планирую пользоваться так же часто как и сейчас"])
    q8_mask=make_q8_prefilter("WB")|make_q8_prefilter("Ozon")|make_q8_prefilter("Яндекс")
    row=write_q_block(ws,row,"Q9.  Почему вы будете дальше пользоваться банком-маркетплейсом?",
                      calc_table(df,segs,
                          {bk:BANK_COLS[bk]["q8"] for bk in["WB","Ozon","Яндекс"]},
                          Q8_ANS, Q8_DISPLAY, alpha=alpha, pre_filter=q8_mask),n)

    row+=1
    ws.cell(row,2,"База: все опрошенные").font=R(8)
    row+=2

    # ====== CROSSTABS ======
    # Helper: build cross-segs for one bank with Q2 answers as columns
    def build_cross_q2_segs(df, bk):
        sub=df[df["bank_key"]==bk]
        col=BANK_COLS[bk]["q2"]
        segs_out=[]
        for i,av in enumerate(Q2_ANS[bk]):
            if av is None: continue
            mask=hits(sub[col] if col in sub.columns else pd.Series(dtype=str), av)
            lab=Q2_DISPLAY[i][:30]
            ltr=LETTERS[len(segs_out)]
            segs_out.append((lab,ltr,sub[mask]))
        return segs_out

    def write_cross_header_q2_q4(ws, row, banks_segs):
        # Row: bank names spanning their columns
        c=ws.cell(row,2," "); c.font=B(9); c.fill=FILL_LIGHT
        col=3
        for bk,bk_label,bk_segs in banks_segs:
            start=col
            for i in range(len(bk_segs)): col+=1
            if col-1>start:
                ws.merge_cells(start_row=row,start_column=start,end_row=row,end_column=col-1)
            c=ws.cell(row,start,bk_label); c.fill=FILL_LIGHT; c.font=B(9); c.alignment=CENTER
        row+=1
        # Row: Q2 answer labels + letters
        c=ws.cell(row,2,ans_row_lbl if "ans_row_lbl" in dir() else ""); c.font=B(8); c.fill=FILL_LIGHT
        col=3
        for bk,bk_label,bk_segs in banks_segs:
            for lab,ltr,_ in bk_segs:
                c=ws.cell(row,col,f"{lab[:25]} ({ltr})"); c.fill=FILL_LIGHT; c.font=B(8); c.alignment=WRAP
                col+=1
        ws.row_dimensions[row].height=50
        return row+1

    def write_cross_block_q2_qX(ws, row, title, df, bk_segs_map, qX_col_bk, ans_X, disp_X, alpha=0.05, add_top2=False, top2i=(0,1), bot2i=(3,4)):
        ws.cell(row,2,title).font=B(9); row+=1
        all_bk_segs=[(bk, {"WB":"WB Банк/ WB кошелек","Ozon":"Ozon.Банк","Яндекс":"ЯндексБанк"}[bk], bk_segs_map[bk]) for bk in ["WB","Ozon","Яндекс"]]
        row=write_cross_header_q2_q4(ws,row,all_bk_segs)
        total_cols=sum(len(segs) for _,_,segs in all_bk_segs)
        all_col_segs=[]; all_bk_tags=[]
        for bk,_,bk_segs in all_bk_segs:
            for _,ltr,sub in bk_segs: all_col_segs.append((bk,sub,ltr))
        bases=[len(sub) for _,sub,_ in all_col_segs]
        
        def same_bk_grps():
            grps=[]; cur_bk=None; g=[]
            for i,(bk,_,_) in enumerate(all_col_segs):
                if bk!=cur_bk:
                    if g: grps.append(g)
                    g=[i]; cur_bk=bk
                else: g.append(i)
            if g: grps.append(g)
            return grps
        cg=same_bk_grps()
        
        all_vals=[]
        for ri,d in enumerate(disp_X):
            rvals=[]
            for bk,sub,_ in all_col_segs:
                col2=qX_col_bk.get(bk); av=(ans_X.get(bk) or [])[ri] if bk in ans_X else None
                if not col2 or col2 not in df.columns or not len(sub) or not av:
                    rvals.append(None); continue
                rvals.append(hits(sub[col2] if col2 in sub.columns else pd.Series(dtype=str),av).sum()/len(sub))
            all_vals.append(rvals)
        
        def fmt_row(rvals):
            labeled=apply_z(rvals,bases,cg,alpha)
            fmtd=[]
            for v in labeled:
                if isinstance(v,str) and "%" in v: fmtd.append(v)
                elif isinstance(v,(int,float,np.floating)) and v is not None:
                    p=round(v*100); fmtd.append(f"{p}%" if p>0 else "-")
                else: fmtd.append("-")
            return fmtd
        
        def write_data_row(ws, row, label, vals):
            cb=ws.cell(row,2,label); cb.font=R(9); cb.alignment=LEFT; cb.fill=FILL_LIGHT
            for ci,v in enumerate(vals):
                is_sig=isinstance(v,str) and "%" in v
                c=ws.cell(row,ci+3,v); c.alignment=CENTER; c.font=B(9) if is_sig else R(9)
                if is_sig: c.fill=FILL_GREEN
            return row+1
        
        for ri,d in enumerate(disp_X):
            row=write_data_row(ws,row,d,fmt_row(all_vals[ri]))
        if add_top2:
            def combo2(idxs):
                return [sum(all_vals[ti][ci] for ti in idxs if all_vals[ti][ci] is not None)
                        if any(all_vals[ti][ci] is not None for ti in idxs) else None
                        for ci in range(total_cols)]
            row=write_data_row(ws,row,"Top-2 (полностью соответсвует+скорее соответствует)",fmt_row(combo2(top2i)))
            row=write_data_row(ws,row,"Bottom-2  (полностью не соответствует+скорее не соответствует)",fmt_row(combo2(bot2i)))
        # Base row
        cb=ws.cell(row,2,"База"); cb.font=R(9); cb.alignment=LEFT
        for ci,b in enumerate(bases): ws.cell(row,ci+3,b).alignment=CENTER
        return row+2

    # Build Q2 segs per bank
    q2_segs={bk:build_cross_q2_segs(df,bk) for bk in ["WB","Ozon","Яндекс"]}

    # Cross: Q3xQ5 (Q2 answer segs x Q4 satisfaction)
    row=write_cross_block_q2_qX(ws,row,
        "Q3  Почему вы решили стать клиентом банка-маркетплейса?хQ4.  Насколько банк-маркетплейс соответствует ожиданиям?",
        df,q2_segs,{bk:BANK_COLS[bk]["q4"] for bk in["WB","Ozon","Яндекс"]},
        {bk:Q4_ANS for bk in["WB","Ozon","Яндекс"]},Q4_ANS,alpha,
        add_top2=True,top2i=(0,1),bot2i=(3,4))

    # Cross: Q3xQ6 (Q2 answer segs x Q5 dissatisfy reasons)
    row=write_cross_block_q2_qX(ws,row,
        "Q3  Почему вы решили стать клиентом банка-маркетплейса?хQ5.  Почему банк-маркетплейс не вполне соответствует ожиданиям?",
        df,q2_segs,{bk:BANK_COLS[bk]["q5"] for bk in["WB","Ozon","Яндекс"]},
        Q5_ANS,Q5_DISPLAY,alpha)

    # Cross: Q4topbot x Q8 (top2/bot2 from Q4 x planning Q7)
    # Build Q4 top/bot segs
    def build_q4_topbot_segs(df,bk):
        col=BANK_COLS[bk]["q4"]
        sub=df[df["bank_key"]==bk]
        top2=sub[sub[col].isin(["Полностью соответствует","Скорее соответствует"])] if col in df.columns else sub.iloc[0:0]
        bot2=sub[sub[col].isin(["Скорее не соответствует","Полностью не соответствует"])] if col in df.columns else sub.iloc[0:0]
        return [("Top-2","a",top2),("Bottom-2","b",bot2)]
    q4_segs={bk:build_q4_topbot_segs(df,bk) for bk in["WB","Ozon","Яндекс"]}
    row=write_cross_block_q2_qX(ws,row,
        "Q5 Насколько банк-маркетплейс соответствует ожиданиям...хQ8.  Как вы планируете пользоваться услугами банка-маркетплейса?",
        df,q4_segs,{bk:BANK_COLS[bk]["q7"] for bk in["WB","Ozon","Яндекс"]},
        {bk:Q7_ANS for bk in["WB","Ozon","Яндекс"]},Q7_ANS,alpha)

    # Cross: Q3xQ8
    row=write_cross_block_q2_qX(ws,row,
        "Q3.  Почему вы решили стать клиентом банка-маркетплейса?хQ8.  Как вы планируете пользоваться услугами банка-маркетплейса?",
        df,q2_segs,{bk:BANK_COLS[bk]["q7"] for bk in["WB","Ozon","Яндекс"]},
        {bk:Q7_ANS for bk in["WB","Ozon","Яндекс"]},Q7_ANS,alpha)

def build_quotes(ws, df):
    ws.column_dimensions["A"].width=25; ws.column_dimensions["B"].width=80
    from openpyxl.styles import PatternFill,Font
    FILL_DARK2=PatternFill("solid",fgColor="0070C0")
    c=ws.cell(1,1,"Цитаты Другое"); c.fill=FILL_DARK2; c.font=B(11,True)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=2)
    row=3
    # Q5 "другое" - why continue (Q9 open text from q8 col + 3.Уточните)
    sections=[
        ("Q3. Почему вы решили стать клиентом банка-маркетплейса? — ответ «Другое»",
         [("WB","WB Банк/ WB кошелек",BANK_COLS["WB"]["q2"],"Другое"),
          ("Ozon","Ozon.Банк",BANK_COLS["Ozon"]["q2"],"Другое"),
          ("Яндекс","Яндекс.Банк",BANK_COLS["Яндекс"]["q2"],"Другое")]),
        ("Q9. Почему вы будете дальше пользоваться банком-маркетплейсом? — ответ «Другое»",
         [("WB","WB Банк/ WB кошелек",BANK_COLS["WB"]["q8"],"Другое"),
          ("Ozon","Ozon.Банк",BANK_COLS["Ozon"]["q8"],"Другое"),
          ("Яндекс","Яндекс.Банк",BANK_COLS["Яндекс"]["q8"],"Другое")]),
    ]
    for sec_title, bk_list in sections:
        c=ws.cell(row,1,sec_title); c.font=B(9); c.fill=FILL_LIGHT
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2); row+=1
        for bk,bk_label,col,filter_val in bk_list:
            sub=df[df["bank_key"]==bk]
            if col not in df.columns: continue
            has_other=sub[sub[col].astype(str).str.contains(filter_val,na=False)]
            c=ws.cell(row,1,bk_label); c.font=B(9); c.fill=FILL_LIGHT; row+=1
    # Q9 open texts from "8. Почему..." with open text col
    # Actual open answer is in a separate column - unfortunately not in this dataset
    # We use the "3. Уточните" column for Q3 reklama
    uточн_col="3. Уточните, какая именно реклама вас привлекла?"
    if uточн_col in df.columns:
        c=ws.cell(row,1,"Q3. Уточните, какая именно реклама вас привлекла?"); c.font=B(9); c.fill=FILL_LIGHT
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=2); row+=1
        for _,r2 in df.dropna(subset=[uточн_col]).iterrows():
            bk=r2.get("bank_key","")
            txt=str(r2[uточн_col]).strip()
            if txt and txt!="nan":
                ws.cell(row,1,bk).font=R(9); ws.cell(row,2,txt).font=R(9); row+=1

def build_legend(ws):
    ws.column_dimensions["B"].width=80
    c=ws.cell(2,2,"Условные обозначения: как читать таблицы"); c.font=B(11); c.alignment=LEFT
    items=[
        "Доля — число от 0 до 100% (уже пересчитана). Если ячейка пуста или '-', значит вопрос не задавался или нет ответивших.",
        "Зелёная ячейка + буква(ы) — значение статистически значимо выше, чем у столбцов с указанными буквами (Z-тест, p<0.05).",
        "WB Банк/WB кошелек (a), Ozon.Банк (b), ЯндексБанк (c) — буквенные коды столбцов в таблице.",
        "Внутри групп «Количество переводов», «Доля переводов», «Сумма переводов» сравнение (a vs b) ведётся внутри одного банка.",
        "База — количество человек, для которых задан вопрос.",
        "Анкета полная = Да — использованы только полностью заполненные анкеты.",
    ]
    for i,t in enumerate(items): ws.cell(4+i,2,t).font=R(9)

def generate_report(input_path, output_path, alpha=0.05):
    print("Загружаем данные...")
    df=load_df(input_path)
    print(f"Строк: {len(df)}")
    wb=Workbook()
    ws1=wb.active; ws1.title="Результаты"
    build_results(ws1,df,alpha)
    ws2=wb.create_sheet("Цитаты Другое"); build_quotes(ws2,df)
    ws3=wb.create_sheet("Усл. обозначения"); build_legend(ws3)
    wb.save(output_path)
    print(f"ГОТОВО: {output_path}")
    return output_path
''')

code = "\n".join(parts)
with open('generate_report.py','w',encoding='utf-8') as f:
    f.write(code)
print("Done. Lines:", len(code.splitlines()))
