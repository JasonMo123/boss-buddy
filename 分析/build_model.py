"""產生「HINO 省油分成與導流收入試算表」。

用法：
    python3 -I 分析/derive_metrics.py   # 先產生 數據推算_每車.csv
    python3 -I 分析/build_model.py      # 需要 pycel（pip install pycel）來驗證公式並排序敏感度

流程：先依假設順序建表並重算，再讀出各假設的擺幅，依影響大小排序敏感度表後重建一次。
所有計算皆為 Excel 公式；只有「數據推算」頁的每車數據與「假設」頁的輸入值是寫入的數字。
"""

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent
OUT = HERE / "HINO省油分成試算.xlsx"
CSV = HERE / "數據推算_每車.csv"

FONT = "Arial"
BLUE, BLACK, GREEN = "0000FF", "000000", "008000"
FILL = {
    "數據": PatternFill("solid", fgColor="DDEBF7"),
    "公開": PatternFill("solid", fgColor="E2EFDA"),
    "假設": PatternFill("solid", fgColor="FCE4D6"),
    "head": PatternFill("solid", fgColor="1F3864"),
    "sub": PatternFill("solid", fgColor="D9D9D9"),
    "key": PatternFill("solid", fgColor="FFF2CC"),
}
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
NT = '#,##0;(#,##0);"-"'
LIT = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
DEC = '#,##0.0;(#,##0.0);"-"'
TYPES = [300, 500, 700]

DIESEL_URL = "https://cna.com.tw/news/ahel/202610030137.aspx"
FIN_URL = "https://www.cna.com.tw/news/afe/202112010362.aspx"
SALES_URL = "https://cars.tvbs.com.tw/car-news/300940"

# 代碼, 項目, 單位, 保守, 基準, 樂觀, 來源類型, 依據, 驗證方式, 列入敏感度, 格式
ASSUMPTIONS = [
    ("N", "試算車輛數", "台", 100, 100, 100, "假設",
     "試算單位：每 100 台加入車輛", "—", False, LIT),
    ("MIX300", "車型組成：300 系列", "%", .50, .50, .50, "數據",
     "比賽資料 20 台中 300 系列 10 台", "實際加入名單", False, PCT),
    ("MIX500", "車型組成：500 系列", "%", .25, .25, .25, "數據",
     "比賽資料 20 台中 500 系列 5 台", "實際加入名單", False, PCT),
    ("MIX700", "車型組成：700 系列", "%", .25, .25, .25, "數據",
     "比賽資料 20 台中 700 系列 5 台", "實際加入名單", False, PCT),
    ("DIESEL", "柴油價格", "元/L", 26, 29.9, 33, "公開",
     f"中油超級柴油牌價 29.9 元/L（2026/10/5–10/11，{DIESEL_URL}）；2026 年初最低約 25.7 元。"
     "保守／樂觀取約 ±10%", "結算時改用當月平均牌價", True, DEC),
    ("IDLE_RED", "怠速可減比例", "%", .15, .30, .45, "假設",
     "冷凍車停靠必須怠速，只有『提早到場等候、排隊等裝卸』這類等待型怠速能減；"
     "做法是調整出車時間", "試辦車隊：調整前後的怠速時數（方法 A 直接量測）", True, PCT),
    ("DRIVE_IMP", "行駛油耗差距可改善比例", "%", .10, .25, .40, "假設",
     "與同車型前 25% 車輛的油耗差距，一部分來自路線與載重，無法靠駕駛習慣改善，只取其中一部分",
     "試辦：建議前後的車隊 L/100km，扣除同類車同期變化（方法 B）", True, PCT),
    ("ADOPT", "建議採納率", "%", .40, .60, .80, "假設",
     "老闆實際採納並執行建議的比例", "LINE 採納按鈕紀錄", True, PCT),
    ("SHARE", "省油分成比例", "%", .20, .20, .20, "假設",
     "決策參數：討論結論為先以 20% 試辦（見『情境與敏感度』頁的分成比例表）",
     "試辦期依採用率調整", False, PCT),
    ("COST_AI", "每台每年 AI 與雲端成本", "元/台/年", "AI:AI_TOTAL", "AI:AI_TOTAL", "AI:AI_TOTAL", "假設",
     "由『AI服務成本』頁計算：LLM 推論＋數據＋平台分攤＋LINE。保守＝高用量",
     "原型實測每次調查的 token 數；試辦期的雲端帳單", True, NT),
    ("COST_SUP", "每台每年客服成本", "元/台/年", "AI:SUP_YR", "AI:SUP_YR", "AI:SUP_YR", "假設",
     "由『AI服務成本』頁計算：客服人力成本 ÷ 每人負責車數。和 AI 成本分開列",
     "試辦期每 100 台的客服工時", True, NT),
    ("ALERT_CONV", "保修警示轉成增量原廠進廠比例", "%", .30, .50, .70, "假設",
     "P 碼警示中，轉成『原本不會進 HINO 原廠』的進廠比例", "HINO 服務廠：警示車輛的進廠紀錄", True, PCT),
    ("TICKET", "每次預測保修單價", "元", 20000, 30000, 50000, "假設",
     "找不到台灣 DPF 清洗／維修的公開行情；海外資料為清洗約百至數百美元、更換數千美元",
     "HINO 服務廠工單平均金額", True, NT),
    ("SVC_M", "保修毛利率", "%", .30, .35, .40, "假設",
     "售後零件與工資的毛利", "和泰財報／HINO 內部數字", True, PCT),
    ("REPL", "自然換車率", "%/年", .06, .08, .10, "假設",
     "約對應 10～16 年的換車週期", "交通部統計：營業貨車車齡分布", True, PCT),
    ("UPLIFT", "換車時選回 HINO 的提升", "百分點", .05, .10, .15, "假設",
     "因 AI 顧問綁定，換車時改選或續選 HINO 的機率提升", "試辦客群與一般客群的回購率比較", True, PCT),
    ("PRICE", "新車平均售價", "元", 2_500_000, 3_000_000, 3_500_000, "假設",
     "找不到台灣公開價目；200 系列 3.49 噸約 110 萬起（2021），中大型車更高", "經銷商報價", True, NT),
    ("VEH_M", "新車毛利率", "%", .04, .06, .08, "假設",
     "總代理與經銷合計毛利", "和泰財報／HINO 內部數字", True, PCT),
    ("FIN_PEN", "新車融資滲透率", "%", .40, .50, .60, "假設",
     f"新車採用分期的比例；和勁企業專營重車分期，重車分期市場一年逾 350 億元（{FIN_URL}）",
     "和潤／和勁內部數字", True, PCT),
    ("FIN_LT", "每件融資案存續期間利差收入", "元", 80000, 100000, 120000, "假設",
     "每件分期案在整個貸款期間的淨利差", "和潤／和勁內部數字", True, NT),
    ("SCALE", "推廣規模（加入車輛數）", "台", 5000, 5000, 5000, "假設",
     f"推廣目標。參考：HINO 年銷量約 6,350 台（{SALES_URL}）", "—", False, LIT),
]
CODES = [a[0] for a in ASSUMPTIONS]
AROW = {code: 7 + k for k, code in enumerate(CODES)}  # 假設頁每個代碼所在列

AI_SHEET = "AI服務成本"
AI_URL = "https://cloud.google.com/gemini-enterprise-agent-platform/generative-ai/pricing"
BQ_URL = "https://cloud.google.com/bigquery/pricing"
LINE_URL = "https://actgsys.com/zh/blog/line-bot-development-cost-guide-2026"

# 代碼, 模型, 輸入, 快取輸入, 輸出（美元／百萬 tokens）, 說明
MODELS = [
    ("PRO", "Gemini 3.1 Pro Preview（標準）", 2.0, 0.20, 12.0,
     "≤200K tokens 區間。調查員、反駁者、老闆難題、實驗設計"),
    ("PRO_B", "Gemini 3.1 Pro Preview（批次）", 1.0, 0.10, 6.0,
     "批次／Flex 為標準價一半；快取價按同比例推估。每週探索不急，用批次"),
    ("FLASH", "Gemini 3.8 Flash", 1.5, 0.15, 7.5,
     "2027/1/1 起的正式價（2026 年底前為半價，上線時已結束）。老闆對話、週報"),
    ("LITE", "Gemini 3.5 Flash-Lite", 0.30, 0.03, 2.50, "異常訊號初步分類"),
]
# 代碼, 項目, 單位, 值, 來源類型, 依據
AI_PARAMS = [
    ("FX", "匯率", "元/美元", 32, "假設", "試算用"),
    ("WEEKS", "每月週數", "週/月", 4.3, "假設", "52 週 ÷ 12 月"),
    ("FLEET", "平均車隊規模", "台/車隊", 5, "假設", "頭家方案以 1–10 台為主；車隊規模的影響見第 6 節"),
    ("DATA", "數據儲存與查詢", "美元/台/月", 0.10, "公開",
     "比賽數據每台每天約 1,400 筆，一年約 0.2 GB；BigQuery 查詢 6.25 美元/TiB、"
     f"每月前 1 TiB 免費，含每日監測查詢仍遠低於此值（{BQ_URL}）"),
    ("PLAT", "雲端平台固定成本", "美元/月", 2000, "假設",
     "資料串流、Cloud Run、監控、備援等不隨車數增加的部分，由推廣規模分攤"),
    ("SCALE", "推廣規模", "台", f"=假設!$E${AROW['SCALE']}", "假設", "引用『假設』頁"),
    ("LINE_MSG", "LINE 推播則數", "則/車隊/月", 10, "假設", "主動建議與週報；回覆老闆提問不計費"),
    ("LINE_PRICE", "LINE 每則推播單價", "元/則", 0.33, "公開",
     f"中用量方案 2026/11 起約 1,000 元／3,000 則（{LINE_URL}）"),
    ("SUP_SAL", "客服人力成本", "元/人/月", 60000, "假設", "月薪含勞健保與辦公分攤"),
]
# 代碼, 項目, 單位, 保守（高用量）, 基準, 樂觀（低用量）, 依據
USAGE = [
    ("SIG", "異常訊號數", "則/台/月", 40, 20, 10, "固定規則監測觸發、送 LLM 初步分類的訊號"),
    ("INV", "深入調查次數", "次/台/月", 2, 1, 0.5, "初步分類後值得深入的訊號；每次含調查員與反駁者"),
    ("EXPL", "每週開放探索", "次/週", 2, 1, 0.5, "不靠規則、讓 LLM 自己找新異常（例：DPF 故障碼與怠速的關聯）"),
    ("QA", "老闆提問數", "題/車隊/月", 60, 30, 15, "LINE 文字或語音提問"),
    ("ESC", "提問升級 Pro 比例", "%", .30, .20, .10, "需要數學推論、路線或派車最佳化的問題"),
    ("EXP", "實驗設計", "次/車隊/月", 4, 2, 1, "換車、換班表等消融實驗"),
    ("SUP_LOAD", "每位客服負責車數", "台/人", 500, 1000, 2000,
     "日常問題由 AI 回答；客服處理帳務、分成爭議、車機與 LINE 綁定"),
]
# 工作, 計費基礎, 模型, 每次輸入（千 tokens）, 快取比例, 每次輸出（千 tokens）, 每月次數公式（用 USAGE／AI_PARAMS 代碼）
TASKS = [
    ("異常訊號初步分類", "每台", "LITE", 3, 0, 0.3, "SIG"),
    ("調查員（約 12 輪工具呼叫）", "每台", "PRO", 200, .6, 20, "INV"),
    ("反駁者（試著推翻調查結論）", "每台", "PRO", 60, .3, 8, "INV"),
    ("每週開放探索", "每車隊", "PRO_B", 500, .3, 50, "EXPL*WEEKS"),
    ("週報", "每車隊", "FLASH", 30, .3, 3, "WEEKS"),
    ("老闆提問（Flash 回答）", "每車隊", "FLASH", 72, .5, 4.8, "QA*(1-ESC)"),
    ("老闆提問（升級 Pro）", "每車隊", "PRO", 72, .5, 6, "QA*ESC"),
    ("實驗設計", "每車隊", "PRO", 100, .3, 10, "EXP"),
]


def font(color=BLACK, bold=False, size=10, italic=False):
    return Font(name=FONT, color=color, bold=bold, size=size, italic=italic)


def header_row(ws, row, values, start_col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = font("FFFFFF", bold=True)
        c.fill = FILL["head"]
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = font(bold=True, size=14)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = font(italic=True, color="595959")


def apply_default_font(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None and (c.font is None or c.font.name != FONT):
                    c.font = Font(name=FONT, size=c.font.size or 10, bold=c.font.bold,
                                  italic=c.font.italic, color=c.font.color)


# ---------------------------------------------------------------- 說明
def sheet_readme(wb):
    ws = wb.active
    ws.title = "說明"
    title(ws, "HINO 省油分成與導流收入試算表", "和泰 AI 黑客松｜商用車頭家智慧夥伴｜商業模式試算")
    rows = [
        ("用途", "估算 AI 車隊顧問每加入 100 台車，和泰每年可得的毛利（省油分成＋保修、新車、融資導流），"
                "以及老闆實際省下的錢；並找出哪個假設對結果影響最大、最該優先驗證。"),
        ("怎麼用", "1. 到『假設』頁 B4 切換情境（保守／基準／樂觀）。"),
        ("", "2. 只改『假設』頁 D:F 欄與『AI服務成本』頁的藍字數字，其他頁會自動更新。"
             "『假設』頁的兩項服務成本是綠字，由『AI服務成本』頁算出。"),
        ("", "3. 看『情境與敏感度』頁：情境比較、各假設的影響排序、分成比例的取捨。"),
        ("頁面", "假設：所有輸入值｜數據推算：比賽數據推出的每車指標與車型彙總｜"
                "AI服務成本：模型單價 × 各項工作的 token 用量，算出每台每年的 AI 與雲端成本，客服另列｜"
                "計算：每一欄是一組輸入（情境或單一假設變動），列是計算步驟｜情境與敏感度：結果彙整"),
        ("顏色：底色", "藍底＝比賽數據推算｜綠底＝公開資料（附連結與日期）｜橘底＝我們的假設（附驗證方式）"),
        ("顏色：字色", "藍字＝寫入的數字（可修改）｜黑字＝公式｜綠字＝引用其他頁的公式"),
        ("數據推算方式", "腳本：分析/derive_metrics.py（從原始官方 Excel 重算，可重跑驗證）。"
                      "用油＝累積油耗每日差值；排除相鄰兩筆里程或油耗跳動超過 5 的異常（例：總里程歸零）。"),
        ("", "日均量用所有出勤日；油耗比率（L/100km）只用行駛 ≥30 km 的有效日。"),
        ("", "年營運天數＝主要抽樣日（≥5 台車有資料的 22 天）中該車出現的比例 × 365。"
             "前提：沒資料代表沒出車，而不是車機斷線。"),
        ("", "可省油＝怠速用油 × 怠速可減比例 ＋ 行駛用油中超出同車型前 25% 的部分 × 可改善比例，再乘採納率。"),
        ("", "保修警示率＝車型內出現 P 開頭故障碼的車輛比例（U 開頭多為通訊類，不計）。"
             "資料只涵蓋約 1/3 的日子，實際一年的警示率可能更高。"),
        ("限制", "有 1 台 500 系列沒有任何一天行駛 ≥30 km，未列入數據推算；車型組成仍依原始 10:5:5。"),
        ("", "每車只有 11～23 天資料，車型內只有 4～10 台車，所有數據推算值都是量級參考，上線後需以連續資料重估。"),
        ("", "分成以年為單位估算；實際每條建議只收 12 個月，穩定狀態需靠持續找到新建議維持。"),
        ("", "車險（第二階段）、碳排報告、報價助手等未列入。"),
    ]
    r = 4
    for k, v in rows:
        ws.cell(row=r, column=1, value=k).font = font(bold=True)
        c = ws.cell(row=r, column=2, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 120


# ---------------------------------------------------------------- 假設
def sheet_assumptions(wb, ai_refs):
    ws = wb.create_sheet("假設")
    title(ws, "假設與輸入", "只修改 D:F 欄的藍字；B4 切換情境。保守／樂觀是從和泰角度看的不利／有利值。")
    ws["A4"] = "目前情境"
    ws["A4"].font = font(bold=True)
    ws["B4"] = "基準"
    ws["B4"].font = font(BLUE, bold=True)
    ws["B4"].fill = FILL["key"]
    ws["C4"] = '=IF(B4="保守",1,IF(B4="樂觀",3,2))'
    ws["D4"] = "← 情境代碼（公式）"
    ws["D4"].font = font(italic=True, color="595959")
    dv = DataValidation(type="list", formula1='"保守,基準,樂觀"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B4")

    header_row(ws, 6, ["代碼", "項目", "單位", "保守", "基準", "樂觀", "來源類型", "依據／說明", "驗證方式"])
    rows = {}
    for i, (code, name, unit, lo, base, hi, kind, basis, check, _sens, fmt) in enumerate(ASSUMPTIONS):
        r = 7 + i
        rows[code] = r
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=unit)
        for j, val in enumerate((lo, base, hi)):
            linked = isinstance(val, str) and val.startswith("AI:")
            if linked:
                val = f"={ai_refs[val[3:]][j]}"
            c = ws.cell(row=r, column=4 + j, value=val)
            c.font = font(GREEN if linked else BLUE)
            c.fill = FILL[kind]
            c.number_format = fmt
        ws.cell(row=r, column=7, value=kind).fill = FILL[kind]
        ws.cell(row=r, column=8, value=basis).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=9, value=check).alignment = Alignment(wrap_text=True, vertical="top")
        for col in range(1, 10):
            ws.cell(row=r, column=col).border = THIN
    last = 7 + len(ASSUMPTIONS) - 1
    chk = last + 2
    ws.cell(row=chk, column=2, value="檢查：車型組成合計").font = font(bold=True)
    for j in range(3):
        col = L(4 + j)
        c = ws.cell(row=chk, column=4 + j,
                    value=f"=IF(ABS({col}{rows['MIX300']}+{col}{rows['MIX500']}+{col}{rows['MIX700']}-1)<0.0001,\"OK\",\"≠100%\")")
        c.alignment = Alignment(horizontal="center")
    widths = [12, 30, 10, 13, 13, 13, 10, 70, 45]
    for i, w in enumerate(widths):
        ws.column_dimensions[L(i + 1)].width = w
    ws.freeze_panes = "D7"
    assert rows == AROW
    return rows


# ---------------------------------------------------------------- 數據推算
def sheet_data(wb, df):
    ws = wb.create_sheet("數據推算")
    title(ws, "數據推算：比賽數據 → 每車指標與車型彙總",
          "每車數據由 分析/derive_metrics.py 從原始官方 Excel 產生（藍底藍字＝寫入值）；車型彙總為公式。")

    vstart = 20
    cols = list(df.columns)
    header_row(ws, vstart - 1, cols)
    for i, rec in enumerate(df.itertuples(index=False)):
        r = vstart + i
        for j, val in enumerate(rec):
            c = ws.cell(row=r, column=j + 1, value=val)
            c.font = font(BLUE)
            c.fill = FILL["數據"]
            name = cols[j]
            if "比例" in name or "佔比" in name:
                c.number_format = PCT
            elif isinstance(val, float):
                c.number_format = DEC
    vend = vstart + len(df) - 1
    notes = {
        "抽樣日出現比例": "主要抽樣日（≥5 台車有資料）中這台車有資料的比例，近似出勤率",
        "日均用油_L": "所有出勤日的平均；用油＝累積油耗（can.engine.totalFuelUsed）每日差值",
        "日均怠速用油_L": "carStatus＝2（怠速）期間的用油，所有出勤日平均",
        "行駛油耗_L每百公里": "只用有效日（≥30 km）：(總用油－怠速用油)／總里程 × 100",
        "同車型前25%行駛油耗": "同車型行駛油耗的第 25 百分位數",
        "超出前25%比例": "1 －（前25% ÷ 本車），小於 0 取 0",
        "日均超出前25%用油_L": "日均行駛用油 × 超出前25%比例：理論上可改善的上限",
        "有P碼": "曾出現 P 開頭故障碼＝1（U 開頭多為通訊類，不計）",
        "有DPF碼": "曾出現 P2002（DPF 效率低）或 P244B（DPF 壓差）＝1",
    }
    for j, name in enumerate(cols):
        if name in notes:
            ws.cell(row=vstart - 1, column=j + 1).comment = Comment(notes[name], "derive_metrics.py")

    def col_of(name):
        return L(cols.index(name) + 1)

    type_rng = f"${col_of('車型')}${vstart}:${col_of('車型')}${vend}"

    def rng(name):
        c = col_of(name)
        return f"${c}${vstart}:${c}${vend}"

    # 車型彙總
    ws["A4"] = "車型彙總（公式）"
    ws["A4"].font = font(bold=True, size=12)
    header_row(ws, 5, ["指標", "單位"] + [f"{t} 系列" for t in TYPES] + ["說明"])
    summary = [
        ("CARS", "納入車輛數", "台", lambda t: f"=COUNTIF({type_rng},{t})", LIT, ""),
        ("OPS", "出勤比例", "%", lambda t: f"=AVERAGEIFS({rng('抽樣日出現比例')},{type_rng},{t})", PCT,
         "車型內平均"),
        ("DAYS", "年營運天數", "天", None, LIT, "＝出勤比例 × 365"),
        ("DAILY_L", "每台日均用油", "L/天", lambda t: f"=AVERAGEIFS({rng('日均用油_L')},{type_rng},{t})", DEC, ""),
        ("IDLE_L", "每台日均怠速用油", "L/天", lambda t: f"=AVERAGEIFS({rng('日均怠速用油_L')},{type_rng},{t})", DEC, ""),
        ("IDLE_SHARE", "怠速用油佔比", "%", None, PCT, "＝怠速用油 ÷ 總用油"),
        ("GAP_L", "每台日均超出前 25% 用油", "L/天",
         lambda t: f"=AVERAGEIFS({rng('日均超出前25%用油_L')},{type_rng},{t})", DEC, "行駛面可改善的上限"),
        ("ALERT", "保修警示率（P 碼）", "%/年", lambda t: f"=AVERAGEIFS({rng('有P碼')},{type_rng},{t})", PCT,
         "資料只涵蓋約 1/3 的日子，可能低估"),
        ("DPF", "DPF 碼車輛比例", "%", lambda t: f"=AVERAGEIFS({rng('有DPF碼')},{type_rng},{t})", PCT, "參考"),
        ("YEAR_L", "每台年用油", "L/年", None, LIT, "＝日均用油 × 年營運天數"),
    ]
    refs = {}
    for i, (key, name, unit, fn, fmt, note) in enumerate(summary):
        r = 6 + i
        refs[key] = r
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=unit)
        ws.cell(row=r, column=6, value=note).font = font(italic=True, color="595959")
        for j, t in enumerate(TYPES):
            col = L(3 + j)
            if key == "DAYS":
                f = f"={col}{refs['OPS']}*365"
            elif key == "IDLE_SHARE":
                f = f"=IFERROR({col}{refs['IDLE_L']}/{col}{refs['DAILY_L']},0)"
            elif key == "YEAR_L":
                f = f"={col}{refs['DAILY_L']}*{col}{refs['DAYS']}"
            else:
                f = fn(t)
            c = ws.cell(row=r, column=3 + j, value=f)
            c.number_format = fmt
            c.border = THIN
    ws.column_dimensions["A"].width = 24
    for j in range(2, len(cols) + 1):
        ws.column_dimensions[L(j)].width = 13
    ws.column_dimensions["F"].width = 30
    ws.freeze_panes = ws.cell(row=vstart, column=4)
    # 回傳：{key: {300: "數據推算!$C$9", ...}}
    return {k: {t: f"數據推算!${L(3 + j)}${r}" for j, t in enumerate(TYPES)} for k, r in refs.items()}


# ---------------------------------------------------------------- AI服務成本
def sheet_ai(wb):
    ws = wb.create_sheet(AI_SHEET)
    title(ws, "AI 服務成本：模型單價 × 各項工作的 token 用量",
          "只改藍字。D:F 欄依『假設』頁的情境順序排列：保守＝高用量、樂觀＝低用量。客服成本另列，不混入 AI 成本。")

    def sec(r, text):
        ws.cell(row=r, column=1, value=text).font = font(bold=True, size=12)

    # 1. 模型單價
    sec(4, f"1. 模型單價（美元／百萬 tokens，Vertex AI 官方牌價，2026/10/12 查詢：{AI_URL}）")
    header_row(ws, 5, ["代碼", "模型", "輸入", "快取輸入", "輸出", "說明"])
    m0 = 6
    for k, (code, name, pin, pc, pout, note) in enumerate(MODELS):
        r = m0 + k
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        for j, v in enumerate((pin, pc, pout)):
            c = ws.cell(row=r, column=3 + j, value=v)
            c.font = font(BLUE)
            c.fill = FILL["公開"]
            c.number_format = "0.00"
        ws.cell(row=r, column=6, value=note)
    m1 = m0 + len(MODELS) - 1
    price = lambda col, model: f"INDEX(${col}${m0}:${col}${m1},MATCH({model},$A${m0}:$A${m1},0))"

    # 2. 通用參數
    p0 = m1 + 3
    sec(p0 - 1, "2. 通用參數")
    header_row(ws, p0, ["代碼", "項目", "單位", "值", "來源類型", "依據"])
    prow = {}
    for k, (code, name, unit, v, kind, basis) in enumerate(AI_PARAMS):
        r = p0 + 1 + k
        prow[code] = r
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=unit)
        c = ws.cell(row=r, column=4, value=v)
        c.font = font(GREEN if isinstance(v, str) else BLUE)
        c.fill = FILL[kind]
        c.number_format = "#,##0.00" if isinstance(v, float) else "#,##0"
        ws.cell(row=r, column=5, value=kind).fill = FILL[kind]
        ws.cell(row=r, column=6, value=basis)
    P = lambda code: f"$D${prow[code]}"

    # 3. 用量（依情境）
    u0 = p0 + len(AI_PARAMS) + 3
    sec(u0 - 1, "3. 用量（依情境）")
    header_row(ws, u0, ["代碼", "項目", "單位", "保守（高用量）", "基準", "樂觀（低用量）", "依據"])
    urow = {}
    for k, (code, name, unit, hi_use, base, lo_use, basis) in enumerate(USAGE):
        r = u0 + 1 + k
        urow[code] = r
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=unit)
        for j, v in enumerate((hi_use, base, lo_use)):
            c = ws.cell(row=r, column=4 + j, value=v)
            c.font = font(BLUE)
            c.fill = FILL["假設"]
            c.number_format = PCT if unit == "%" else '#,##0.0;(#,##0.0);"-"'
        ws.cell(row=r, column=7, value=basis)

    def freq(expr, col):
        """把次數公式裡的代碼換成儲存格：用量代碼取該情境欄，通用參數取 D 欄。"""
        return "=" + re.sub(r"[A-Z_]+", lambda m: f"{col}${urow[m[0]]}" if m[0] in urow else P(m[0]), expr)

    # 4. 各項工作
    t0 = u0 + len(USAGE) + 3
    sec(t0 - 1, "4. 各項工作的 token 用量與每月成本（美元）")
    header_row(ws, t0, ["工作", "計費基礎", "模型", "每次輸入（千 tokens）", "快取比例", "每次輸出（千 tokens）",
                        "每次成本（美元）", "每月次數：保守", "基準", "樂觀", "每月成本：保守", "基準", "樂觀"])
    tr0 = t0 + 1
    for k, (name, basis, model, tin, cached, tout, fexpr) in enumerate(TASKS):
        r = tr0 + k
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=basis)
        ws.cell(row=r, column=3, value=model).font = font(BLUE)
        for j, (v, fmt) in enumerate(((tin, "#,##0.0"), (cached, "0%"), (tout, "#,##0.0"))):
            c = ws.cell(row=r, column=4 + j, value=v)
            c.font = font(BLUE)
            c.fill = FILL["假設"]
            c.number_format = fmt
        ws.cell(row=r, column=7, value=(
            f"=(D{r}*(1-E{r})*{price('C', f'C{r}')}+D{r}*E{r}*{price('D', f'C{r}')}"
            f"+F{r}*{price('E', f'C{r}')})/1000")).number_format = "0.0000"
        for j, ucol in enumerate("DEF"):
            ws.cell(row=r, column=8 + j, value=freq(fexpr, ucol)).number_format = "#,##0.0"
            ws.cell(row=r, column=11 + j, value=f"=$G{r}*{L(8 + j)}{r}").number_format = "#,##0.00"
    tr1 = tr0 + len(TASKS) - 1

    # 5. 每台每年成本
    s0 = tr1 + 3
    sec(s0 - 1, "5. 每台每年服務成本（元）")
    header_row(ws, s0, ["項目", "", "單位", "保守（高用量）", "基準", "樂觀（低用量）", "說明"])
    srow = {}
    lines = [
        ("LLM_TRUCK", "LLM：每台車的工作", "美元/台/月",
         lambda c, mc: f'=SUMIF($B${tr0}:$B${tr1},"每台",{mc}{tr0}:{mc}{tr1})', "#,##0.00", "訊號分類、調查員、反駁者"),
        ("LLM_FLEET", "LLM：每個車隊的工作", "美元/車隊/月",
         lambda c, mc: f'=SUMIF($B${tr0}:$B${tr1},"每車隊",{mc}{tr0}:{mc}{tr1})', "#,##0.00",
         "探索、週報、老闆提問、實驗設計"),
        ("LLM_YR", "LLM 推論", "元/台/年",
         lambda c, mc: f"=({c}{{LLM_TRUCK}}+{c}{{LLM_FLEET}}/{P('FLEET')})*12*{P('FX')}", NT,
         "每車隊的工作由平均車隊規模分攤"),
        ("DATA_YR", "數據儲存與查詢", "元/台/年", lambda c, mc: f"={P('DATA')}*12*{P('FX')}", NT, ""),
        ("PLAT_YR", "平台固定成本分攤", "元/台/年",
         lambda c, mc: f"=IFERROR({P('PLAT')}*12*{P('FX')}/{P('SCALE')},0)", NT, "推廣初期車少時會高很多"),
        ("LINE_YR", "LINE 推播", "元/台/年",
         lambda c, mc: f"={P('LINE_MSG')}*{P('LINE_PRICE')}*12/{P('FLEET')}", NT, ""),
        ("AI_TOTAL", "AI 與雲端成本合計", "元/台/年",
         lambda c, mc: f"={c}{{LLM_YR}}+{c}{{DATA_YR}}+{c}{{PLAT_YR}}+{c}{{LINE_YR}}", NT, "→『假設』頁 COST_AI"),
        ("SUP_YR", "客服成本", "元/台/年",
         lambda c, mc: f"=IFERROR({P('SUP_SAL')}*12/{c}${urow['SUP_LOAD']},0)", NT, "→『假設』頁 COST_SUP"),
        ("TOTAL", "服務成本合計", "元/台/年", lambda c, mc: f"={c}{{AI_TOTAL}}+{c}{{SUP_YR}}", NT, ""),
        ("LLM_PCT", "LLM 佔服務成本比例", "%", lambda c, mc: f"=IFERROR({c}{{LLM_YR}}/{c}{{TOTAL}},0)", PCT, ""),
    ]
    for k, (key, name, unit, fn, fmt, note) in enumerate(lines):
        srow[key] = s0 + 1 + k
    for key, name, unit, fn, fmt, note in lines:
        r = srow[key]
        bold = key in ("AI_TOTAL", "SUP_YR", "TOTAL")
        ws.cell(row=r, column=1, value=name).font = font(bold=bold)
        ws.cell(row=r, column=3, value=unit)
        ws.cell(row=r, column=7, value=note).font = font(italic=True, color="595959")
        for j, (c, mc) in enumerate(zip("DEF", "KLM")):
            cell = ws.cell(row=r, column=4 + j, value=fn(c, mc).format(**srow))
            cell.number_format = fmt
            cell.border = THIN
            if bold:
                cell.font = font(bold=True)
            if key in ("AI_TOTAL", "SUP_YR"):
                cell.fill = FILL["key"]

    # 6. 車隊規模的影響
    f0 = srow["LLM_PCT"] + 3
    sec(f0 - 1, "6. 車隊規模的影響：每台每年 AI 與雲端成本（元）")
    header_row(ws, f0, ["車隊規模（台）", "", "", "保守（高用量）", "基準", "樂觀（低用量）", "說明"])
    for k, n in enumerate((1, 3, 5, 10, 20, 50)):
        r = f0 + 1 + k
        c = ws.cell(row=r, column=1, value=n)
        c.font = font(BLUE)
        for j, col in enumerate("DEF"):
            ws.cell(row=r, column=4 + j, value=(
                f"=({col}${srow['LLM_TRUCK']}+{col}${srow['LLM_FLEET']}/$A{r})*12*{P('FX')}"
                f"+{col}${srow['DATA_YR']}+{col}${srow['PLAT_YR']}"
                f"+{P('LINE_MSG')}*{P('LINE_PRICE')}*12/$A{r}")).number_format = NT
    ws.cell(row=f0 + 1, column=7, value="車隊越小，探索與對話等每車隊的工作攤在越少台車上").font = \
        font(italic=True, color="595959")

    widths = [30, 30, 12, 16, 14, 16, 60, 14, 10, 10, 14, 10, 10]
    for k, w in enumerate(widths):
        ws.column_dimensions[L(k + 1)].width = w
    return {key: {j: f"{AI_SHEET}!${col}${srow[key]}" for j, col in enumerate("DEF")}
            for key in ("AI_TOTAL", "SUP_YR")}


# ---------------------------------------------------------------- 計算
def sheet_calc(wb, arow, data, sens_order):
    ws = wb.create_sheet("計算")
    title(ws, "計算（每一欄是一組輸入，列是計算步驟）",
          "第 4–6 列決定該欄用哪個情境、覆寫哪個假設；第 8 列起的輸入值會依此自動帶入。")

    variants = [("目前情境", "=假設!$C$4", None, None)]
    variants += [(n, i, None, None) for i, n in ((1, "保守"), (2, "基準"), (3, "樂觀"))]
    for code in sens_order:
        r = arow[code]
        variants.append((f"{code}｜保守值", 2, code, f"=假設!$D${r}"))
        variants.append((f"{code}｜樂觀值", 2, code, f"=假設!$F${r}"))
    for s in (.15, .20, .25, .30):
        variants.append((f"分成 {int(s * 100)}%", 2, "SHARE", s))

    ws["A3"], ws["A4"], ws["A5"], ws["A6"] = "欄位", "情境代碼（1保守 2基準 3樂觀）", "覆寫假設代碼", "覆寫值"
    for a in ("A3", "A4", "A5", "A6"):
        ws[a].font = font(bold=True)
    vcol = {}
    for k, (name, lvl, code, val) in enumerate(variants):
        c = 3 + k
        vcol[name] = L(c)
        h = ws.cell(row=3, column=c, value=name)
        h.font = font("FFFFFF", bold=True)
        h.fill = FILL["head"]
        h.alignment = Alignment(horizontal="center", wrap_text=True)
        x = ws.cell(row=4, column=c, value=lvl)
        x.font = font(GREEN if isinstance(lvl, str) else BLUE)
        ws.cell(row=5, column=c, value=code).font = font(BLUE)
        v = ws.cell(row=6, column=c, value=val)
        v.font = font(GREEN if isinstance(val, str) else BLUE)
    ncols = len(variants)

    row = 8
    ws.cell(row=row, column=1, value="輸入值").font = font(bold=True)
    ws.cell(row=row, column=1).fill = FILL["sub"]
    row += 1
    irow = {}
    fmt_of = {a[0]: a[10] for a in ASSUMPTIONS}
    for code in CODES:
        ar = arow[code]
        irow[code] = row
        ws.cell(row=row, column=1, value=code)
        ws.cell(row=row, column=2, value=ASSUMPTIONS[CODES.index(code)][1])
        for k in range(ncols):
            col = L(3 + k)
            f = f'=IF({col}$5="{code}",{col}$6,INDEX(假設!$D${ar}:$F${ar},1,{col}$4))'
            c = ws.cell(row=row, column=3 + k, value=f)
            c.font = font(GREEN)
            c.number_format = fmt_of[code]
        row += 1

    crow = {}

    def add(key, label, unit, builder, fmt=NT, bold=False, fill=None):
        nonlocal row
        crow[key] = row
        ws.cell(row=row, column=1, value=label).font = font(bold=bold)
        ws.cell(row=row, column=2, value=unit)
        for k in range(ncols):
            col = L(3 + k)
            c = ws.cell(row=row, column=3 + k, value=builder(col))
            c.number_format = fmt
            c.border = THIN
            if bold:
                c.font = font(bold=True)
            if fill:
                c.fill = fill
        row += 1

    def section(name):
        nonlocal row
        row += 1
        ws.cell(row=row, column=1, value=name).font = font(bold=True)
        ws.cell(row=row, column=1).fill = FILL["sub"]
        row += 1

    i = lambda code, col: f"{col}{irow[code]}"
    r = lambda key, col: f"{col}{crow[key]}"

    for t in TYPES:
        section(f"{t} 系列（每台）")
        add(f"N{t}", f"{t} 車數", "台", lambda c, t=t: f"={i('N', c)}*{i(f'MIX{t}', c)}", DEC)
        add(f"FUEL{t}", f"{t} 每台年用油", "L/年", lambda c, t=t: f"={data['YEAR_L'][t]}", LIT)
        add(f"SAVE{t}", f"{t} 每台年可省油", "L/年",
            lambda c, t=t: (f"={data['DAYS'][t]}*({data['IDLE_L'][t]}*{i('IDLE_RED', c)}"
                            f"+{data['GAP_L'][t]}*{i('DRIVE_IMP', c)})*{i('ADOPT', c)}"), DEC)

    section("車隊合計")
    add("FUEL", "車隊年用油", "L/年",
        lambda c: "=" + "+".join(f"{r(f'N{t}', c)}*{r(f'FUEL{t}', c)}" for t in TYPES), LIT)
    add("FUEL_NT", "車隊年油錢", "元/年", lambda c: f"={r('FUEL', c)}*{i('DIESEL', c)}")
    add("SAVE", "車隊年省油", "L/年",
        lambda c: "=" + "+".join(f"{r(f'N{t}', c)}*{r(f'SAVE{t}', c)}" for t in TYPES), LIT)
    add("SAVE_PCT", "省油幅度", "%", lambda c: f"=IFERROR({r('SAVE', c)}/{r('FUEL', c)},0)", PCT)
    add("SAVE_NT", "省下油錢", "元/年", lambda c: f"={r('SAVE', c)}*{i('DIESEL', c)}")

    section("和泰收入（毛利）")
    add("GS", "① 省油分成", "元/年", lambda c: f"={r('SAVE_NT', c)}*{i('SHARE', c)}")
    add("ALERTS", "預測保修警示件數", "件/年",
        lambda c: "=" + "+".join(f"{r(f'N{t}', c)}*{data['ALERT'][t]}" for t in TYPES), DEC)
    add("JOBS", "增量原廠進廠件數", "件/年", lambda c: f"={r('ALERTS', c)}*{i('ALERT_CONV', c)}", DEC)
    add("SVC", "② 保修導流毛利", "元/年",
        lambda c: f"={r('JOBS', c)}*{i('TICKET', c)}*{i('SVC_M', c)}")
    add("UNITS", "增量新車台數", "台/年", lambda c: f"={i('N', c)}*{i('REPL', c)}*{i('UPLIFT', c)}", '0.00')
    add("VEH", "③ 新車導流毛利", "元/年", lambda c: f"={r('UNITS', c)}*{i('PRICE', c)}*{i('VEH_M', c)}")
    add("FIN", "④ 融資利差", "元/年", lambda c: f"={r('UNITS', c)}*{i('FIN_PEN', c)}*{i('FIN_LT', c)}")
    add("GP", "毛利合計", "元/年",
        lambda c: f"={r('GS', c)}+{r('SVC', c)}+{r('VEH', c)}+{r('FIN', c)}", bold=True)
    add("COSTS_AI", "AI 與雲端成本", "元/年", lambda c: f"={i('N', c)}*{i('COST_AI', c)}")
    add("COSTS_SUP", "客服成本", "元/年", lambda c: f"={i('N', c)}*{i('COST_SUP', c)}")
    add("COSTS", "服務成本合計", "元/年", lambda c: f"={r('COSTS_AI', c)}+{r('COSTS_SUP', c)}")
    add("NET", "淨貢獻", "元/年", lambda c: f"={r('GP', c)}-{r('COSTS', c)}", bold=True, fill=FILL["key"])
    add("NET_PER", "每台每年淨貢獻", "元/台/年", lambda c: f"=IFERROR({r('NET', c)}/{i('N', c)},0)",
        bold=True, fill=FILL["key"])
    add("GS_SHARE", "分成佔毛利比例", "%", lambda c: f"=IFERROR({r('GS', c)}/{r('GP', c)},0)", PCT)
    add("NET_SCALE", "推廣規模下年淨貢獻", "元/年", lambda c: f"={r('NET_PER', c)}*{i('SCALE', c)}", bold=True)

    section("老闆")
    add("BOSS", "老闆年淨省（扣分成）", "元/年", lambda c: f"={r('SAVE_NT', c)}-{r('GS', c)}")
    add("BOSS_PER", "老闆每台每年淨省", "元/台/年", lambda c: f"=IFERROR({r('BOSS', c)}/{i('N', c)},0)", bold=True)
    add("BOSS_PER_M", "老闆每台每月淨省", "元/台/月", lambda c: f"={r('BOSS_PER', c)}/12")
    add("BOSS_700", "老闆每台 700 系列每年淨省", "元/台/年",
        lambda c: f"={r('SAVE700', c)}*{i('DIESEL', c)}*(1-{i('SHARE', c)})")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 10
    for k in range(ncols):
        ws.column_dimensions[L(3 + k)].width = 15
    ws.freeze_panes = "C7"
    return vcol, crow


# ---------------------------------------------------------------- 情境與敏感度
def sheet_results(wb, vcol, crow, sens_order):
    ws = wb.create_sheet("情境與敏感度")
    title(ws, "情境比較、敏感度與分成比例", "全部為引用『計算』頁的公式。")

    ws["A4"] = "1. 情境比較（每 100 台車／年）"
    ws["A4"].font = font(bold=True, size=12)
    header_row(ws, 5, ["指標", "單位", "保守", "基準", "樂觀", "目前情境"])
    lines = [
        ("SAVE_PCT", "省油幅度", "%", PCT), ("SAVE_NT", "省下油錢", "元/年", NT),
        ("GS", "① 省油分成", "元/年", NT), ("SVC", "② 保修導流毛利", "元/年", NT),
        ("VEH", "③ 新車導流毛利", "元/年", NT), ("FIN", "④ 融資利差", "元/年", NT),
        ("GP", "毛利合計", "元/年", NT), ("COSTS_AI", "AI 與雲端成本", "元/年", NT),
        ("COSTS_SUP", "客服成本", "元/年", NT), ("COSTS", "服務成本合計", "元/年", NT),
        ("NET", "淨貢獻", "元/年", NT), ("NET_PER", "每台每年淨貢獻", "元/台/年", NT),
        ("GS_SHARE", "分成佔毛利比例", "%", PCT), ("NET_SCALE", "推廣規模下年淨貢獻", "元/年", NT),
        ("BOSS_PER", "老闆每台每年淨省", "元/台/年", NT), ("BOSS_PER_M", "老闆每台每月淨省", "元/台/月", NT),
        ("BOSS_700", "老闆每台 700 系列每年淨省", "元/台/年", NT),
    ]
    for k, (key, name, unit, fmt) in enumerate(lines):
        r = 6 + k
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=unit)
        for j, v in enumerate(("保守", "基準", "樂觀", "目前情境")):
            c = ws.cell(row=r, column=3 + j, value=f"=計算!{vcol[v]}{crow[key]}")
            c.font = font(GREEN, bold=key in ("NET", "NET_PER"))
            c.number_format = fmt
            c.border = THIN
            if key in ("NET", "NET_PER"):
                c.fill = FILL["key"]

    top = 6 + len(lines) + 2
    ws.cell(row=top, column=1, value="2. 敏感度：其他假設固定在基準值，單獨把一個假設換成保守值／樂觀值").font = font(bold=True, size=12)
    ws.cell(row=top + 1, column=1, value="依擺幅由大到小排序：越上面的假設越該優先驗證。結果指標＝每台每年淨貢獻（元）").font = font(italic=True, color="595959")
    header_row(ws, top + 2, ["假設", "代碼", "保守值時", "基準", "樂觀值時", "擺幅"])
    name_of = {a[0]: a[1] for a in ASSUMPTIONS}
    first = top + 3
    for k, code in enumerate(sens_order):
        r = first + k
        ws.cell(row=r, column=1, value=name_of[code])
        ws.cell(row=r, column=2, value=code)
        lo = ws.cell(row=r, column=3, value=f"=計算!{vcol[f'{code}｜保守值']}{crow['NET_PER']}")
        base = ws.cell(row=r, column=4, value=f"=計算!{vcol['基準']}{crow['NET_PER']}")
        hi = ws.cell(row=r, column=5, value=f"=計算!{vcol[f'{code}｜樂觀值']}{crow['NET_PER']}")
        sw = ws.cell(row=r, column=6, value=f"=ABS(E{r}-C{r})")
        for c in (lo, base, hi):
            c.font = font(GREEN)
            c.number_format = NT
        sw.number_format = NT
        sw.font = font(bold=True)
    last = first + len(sens_order) - 1

    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "各假設對每台淨貢獻的擺幅（元）"
    chart.y_axis.title = "擺幅（元/台/年）"
    chart.add_data(Reference(ws, min_col=6, min_row=top + 2, max_row=last), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    chart.x_axis.scaling.orientation = "maxMin"
    chart.legend = None
    chart.height, chart.width = 9, 16
    ws.add_chart(chart, f"H{top}")

    st = last + 3
    ws.cell(row=st, column=1, value="3. 分成比例的取捨（其他假設為基準值）").font = font(bold=True, size=12)
    header_row(ws, st + 1, ["分成比例", "和泰每台每年淨貢獻", "老闆每台每年淨省", "老闆每台每月淨省", "分成佔毛利比例"])
    for k, s in enumerate((15, 20, 25, 30)):
        r = st + 2 + k
        col = vcol[f"分成 {s}%"]
        ws.cell(row=r, column=1, value=f"=計算!{col}{irow_share_ref}")
        for j, key in enumerate(("NET_PER", "BOSS_PER", "BOSS_PER_M", "GS_SHARE")):
            c = ws.cell(row=r, column=2 + j, value=f"=計算!{col}{crow[key]}")
            c.font = font(GREEN)
            c.number_format = PCT if key == "GS_SHARE" else NT
        ws.cell(row=r, column=1).number_format = '0%'
        ws.cell(row=r, column=1).font = font(GREEN)

    ws.column_dimensions["A"].width = 30
    for col in "BCDEF":
        ws.column_dimensions[col].width = 18
    return first, last


irow_share_ref = None


def build(sens_order):
    global irow_share_ref
    df = pd.read_csv(CSV, encoding="utf-8-sig")
    wb = Workbook()
    sheet_readme(wb)
    ai_refs = sheet_ai(wb)
    arow = sheet_assumptions(wb, ai_refs)
    data = sheet_data(wb, df)
    vcol, crow = sheet_calc(wb, arow, data, sens_order)
    irow_share_ref = 9 + CODES.index("SHARE")
    sheet_results(wb, vcol, crow, sens_order)
    wb._sheets = [wb[n] for n in ("說明", "假設", "數據推算", AI_SHEET, "計算", "情境與敏感度")]
    apply_default_font(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    return vcol, crow


def evaluate():
    """用 pycel 計算所有公式，回傳 {"工作表!儲存格": 值}；任何錯誤值都會中止。"""
    from pycel import ExcelCompiler
    xl = ExcelCompiler(filename=str(OUT))
    wb = load_workbook(OUT)
    values, errors, n = {}, [], 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n += 1
                    addr = f"{ws.title}!{c.coordinate}"
                    try:
                        v = xl.evaluate(addr)
                    except Exception as e:  # noqa: BLE001
                        v = f"#ERR {e}"
                    if isinstance(v, str) and v.startswith("#"):
                        errors.append((addr, v))
                    values[addr] = v
    print(f"formulas={n} errors={len(errors)}")
    for e in errors[:20]:
        print("  ", e)
    if errors:
        sys.exit("formula errors")
    return values


def main():
    order = [a[0] for a in ASSUMPTIONS if a[9]]
    vcol, crow = build(order)
    vals = evaluate()
    swing = {}
    for code in order:
        lo = vals[f"計算!{vcol[f'{code}｜保守值']}{crow['NET_PER']}"]
        hi = vals[f"計算!{vcol[f'{code}｜樂觀值']}{crow['NET_PER']}"]
        swing[code] = abs(hi - lo)
    order = sorted(order, key=lambda c: -swing[c])
    build(order)
    evaluate()


if __name__ == "__main__":
    main()
