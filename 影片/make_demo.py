"""產生「頭家夥伴」Demo 影片（英文配音、中英字幕）。

用法（在 HINO競賽/ 底下）：
    <venv>/bin/python -I 影片/make_demo.py                    # 完整影片
    <venv>/bin/python -I 影片/make_demo.py --stills 20,45,90  # 只輸出指定秒數的畫面，檢查版面

需要：Pillow、numpy、ffmpeg、kokoro-onnx、soundfile。
語音：Kokoro-82M（Apache 2.0，本機執行），模型放在 影片/_models/：
  kokoro-v1.0.onnx、voices-v1.0.bin、espeak-ng-data/（從 espeakng_loader 複製；
  espeak-ng 的路徑上限約 160 字元，放在 venv 深層路徑會讀不到）
背景音樂與音效：soundtrack.py 用 numpy 合成，不用外部音檔。
輸出：影片/頭家夥伴_Demo.mp4（字幕已燒進畫面）、影片/頭家夥伴_Demo.srt（中英字幕檔）

數字來源：
  - 每車油錢、怠速、油耗：分析/數據推算_每車.csv（分析/derive_metrics.py 從官方原始數據產生）
  - 停靠點怠速、故障碼日期、油量讀數：官方原始數據（分析/_cache/output.parquet）直接查詢
  - 結算金額：分析/HINO省油分成試算.xlsx 基準情境（車隊可省 426 L/月 × 採納率 60%）
  - 實驗結果與信賴區間：情境模擬，畫面上標「示意」
"""

import argparse
import sys
import glob
import hashlib
import re
import subprocess
from functools import lru_cache
from pathlib import Path

import numpy as np
import soundfile as sf
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
import soundtrack  # noqa: E402

HERE = Path(__file__).resolve().parent
BUILD = HERE / "_build"
OUT = HERE / "頭家夥伴_Demo.mp4"
SRT = HERE / "頭家夥伴_Demo.srt"

W, H, FPS = 1920, 1080, 30
MODELS = HERE / "_models"
VOICE, SPEED = "af_heart", 1.0  # Kokoro 語音與語速
SR = soundtrack.SR
GAP = 0.45  # 兩句旁白之間的停頓（秒）

# ---------------------------------------------------------------- 樣式
BG = (11, 18, 32)
PANEL_BG = (17, 26, 43)
CARD = (27, 39, 61)
BORDER = (48, 63, 88)
TEXT = (236, 242, 248)
MUTED = (148, 163, 184)
DIM = (90, 106, 130)
RED = (248, 113, 113)
AMBER = (251, 191, 36)
GREEN = (74, 222, 128)
BLUE = (96, 165, 250)
TEAL = (45, 212, 191)
CHAT_BG = (196, 210, 230)
LINE_GREEN = (6, 199, 85)
BOSS_BUBBLE = (140, 230, 96)

PHONE = (100, 95, 600, 925)
SCREEN = (112, 107, 588, 913)
CHAT_TOP = 175
PANEL = (680, 95, 1840, 925)
CARD_X, CARD_W, CARD_GAP = 700, 1120, 16


def _pingfang():
    hits = glob.glob("/System/Library/AssetsV2/*/*/AssetData/PingFang.ttc")
    return hits[0] if hits else "/System/Library/Fonts/STHeiti Medium.ttc"


PF = _pingfang()
WEIGHT = {"r": 2, "m": 6, "s": 10, "l": 14}  # PingFang TC 的字重索引


@lru_cache(maxsize=None)
def F(size, w="r"):
    try:
        return ImageFont.truetype(PF, size, index=WEIGHT[w])
    except OSError:
        return ImageFont.truetype(PF, size)


# ---------------------------------------------------------------- 旁白與字幕
# 代碼: (中文字幕, 英文字幕, 給 TTS 唸的英文；None＝同英文字幕)
LINE = {
    "title": ("這是頭家夥伴：一個會自己看數據、主動找問題的車隊 AI 顧問。",
              "Meet Boss Buddy: a fleet AI advisor that reads the data and finds problems on its own.", None),
    "fleet": ("王老闆有 5 台 HINO 貨車，每個月油錢大約 19 萬元。",
              "Mr. Wang runs five HINO trucks. His fuel bill is about 190,000 NT dollars a month.",
              "Mr. Wang runs five Hino trucks. His fuel bill is about one hundred and ninety thousand "
              "N.T. dollars a month."),
    "nightly": ("他沒空看報表。AI 每天晚上自己檢查車隊數據。",
                "He has no time for reports. So every night, the AI checks his fleet data by itself.", None),
    "rule": ("固定規則先篩出異常：ABC-6776 每天怠速 6.1 小時，同型車平均只有 2.3 小時。",
             "Simple rules flag it first: truck 6776 idles 6.1 hours a day. Similar trucks average 2.3.",
             "Simple rules flag it first. Truck six seven seven six idles six point one hours a day. "
             "Similar trucks average two point three."),
    "invest": ("調查員 AI 接著查：怠速集中在萬華一帶同一個停靠點，每次 5 到 8 小時沒熄火。",
               "An investigator agent digs in: the idling happens at one stop in Wanhua, "
               "five to eight hours at a time, engine running.", None),
    "push1": ("換算每個月多花約 3,000 元，AI 主動傳 LINE 給王老闆。",
              "That's about 3,000 dollars a month. The AI messages Mr. Wang on LINE.",
              "That's about three thousand dollars a month. So the AI messages Mr. Wang on the LINE app."),
    "push_back": ("王老闆回：那是冷凍車，停著也要開冷凍機，不能熄火。",
                  "Mr. Wang pushes back: it's a refrigerated truck. The engine has to stay on.", None),
    "constraint": ("AI 把「冷凍機必須運轉」當成限制條件，重新拆解這段怠速。",
                   "The AI takes that as a hard constraint, and breaks the idle time down again.", None),
    "experiment": ("數據分不出多少時間是在排隊等卸貨，所以 AI 設計了兩週的對照實驗。",
                   "The data can't tell how much of it is just waiting in line to unload. "
                   "So the AI designs a two-week experiment.", None),
    "options": ("它也算給老闆看：如果冷凍機能插電待機，每月最多可省約 100 公升。",
                "It also runs the numbers: plugging the reefer into shore power could save up to "
                "100 liters a month.",
                "It also runs the numbers. Plugging the reefer into shore power could save up to "
                "one hundred liters a month."),
    "explore": ("每週，AI 會不帶預設規則，自己在數據裡找新問題。",
                "Every week, the AI also explores the data with no preset rules, looking for new problems.",
                None),
    "theft": ("它先懷疑 ABC-5310 被偷油：油量讀數 1 分鐘從 55% 掉到 25%。",
              "It first suspects fuel theft on truck 5310: the gauge fell from 55 to 25 percent in a minute.",
              "It first suspects fuel theft on truck five three one zero. The gauge fell from fifty five "
              "to twenty five percent in a minute."),
    "skeptic": ("但反駁者發現讀數馬上跳回，同期只用了 2 公升油。是感測器跳動，不打擾老闆。",
                "But a skeptic agent sees it jump right back, with only 2 liters burned. "
                "Sensor noise, so no false alarm.",
                "But a skeptic agent sees it jump right back, with only two liters burned. "
                "Sensor noise, so no false alarm."),
    "dpf": ("真正的發現在 ABC-7610：濾煙器（DPF）故障碼 3 月開始出現，11 月越來越頻繁。",
            "The real finding is truck 7610: its diesel particulate filter codes started in March, "
            "and are getting more frequent in November.",
            "The real finding is truck seven six one zero. Its diesel particulate filter codes started "
            "in March, and are getting more frequent in November."),
    "service": ("AI 建議兩週內進 HINO 服務廠，在拋錨之前處理。",
                "The AI suggests a HINO service visit within two weeks, before it breaks down on the road.",
                "The AI suggests a Hino service visit within two weeks, before it breaks down on the road."),
    "result": ("兩週後實驗結果（示意）：晚一小時到場，每次少怠速將近 2 小時。",
               "Two weeks later, the result, illustrative: arriving an hour later cuts nearly "
               "two hours of idling per stop.", None),
    "billing": ("月底結算以匿名同型車為對照，扣掉季節與路況影響，並按保守下限計費。",
                "At month end, savings are measured against anonymous peer trucks, "
                "and billed at the conservative lower bound.", None),
    "net": ("王老闆這個月淨省約 6,500 元。沒省到，就不收錢。",
            "Mr. Wang keeps about 6,500 dollars this month. No savings, no fee.",
            "Mr. Wang keeps about six thousand five hundred dollars this month. No savings, no fee."),
    "close": ("頭家夥伴：自己找問題、聽得懂限制，省下來才收費。",
              "Boss Buddy finds problems on its own, respects real constraints, and only gets paid "
              "when you save.", None),
    "next": ("下一階段，同一套引擎會用在行車安全。",
             "Next, the same engine moves on to driving safety.", None),
}

# 場景: (代碼, 中文標題, 英文標題, 旁白, 前置秒數, 結尾秒數)
SCENES = [
    ("title", None, None, ["title"], 1.2, 1.0),
    ("fleet", "王老闆的車隊", "Mr. Wang's fleet", ["fleet", "nightly"], 0.8, 0.6),
    ("detect", "① AI 自己發現異常", "AI spots an anomaly", ["rule", "invest", "push1"], 0.4, 1.6),
    ("push", "② 老闆反駁，AI 重新計算", "Owner pushes back", ["push_back", "constraint", "experiment", "options"],
     0.6, 1.6),
    ("explore", "③ 每週探索：自己找新問題", "Weekly open exploration", ["explore", "theft", "skeptic", "dpf", "service"],
     0.4, 1.8),
    ("settle", "④ 月底結算：省下才收費", "Paid only from savings", ["result", "billing", "net"], 0.6, 1.8),
    ("close", None, None, ["close", "next"], 0.6, 2.4),
]


@lru_cache(maxsize=1)
def kokoro():
    from kokoro_onnx import Kokoro
    from kokoro_onnx.config import EspeakConfig
    import espeakng_loader
    cfg = EspeakConfig(data_path=str(MODELS / "espeak-ng-data"), lib_path=espeakng_loader.get_library_path())
    return Kokoro(str(MODELS / "kokoro-v1.0.onnx"), str(MODELS / "voices-v1.0.bin"), espeak_config=cfg)


def synth(lid, text):
    """Kokoro 合成一句旁白，回傳 SR 取樣率的 float32；依（語音、語速、文字）快取。"""
    text = text.replace("LINE", "Line")  # 全大寫會被逐字母唸
    h = hashlib.md5(f"kokoro|{VOICE}|{SPEED}|{text}".encode()).hexdigest()[:10]
    wav = BUILD / f"{lid}_{h}.wav"
    if not wav.exists():
        pcm, sr = kokoro().create(text, voice=VOICE, speed=SPEED, lang="en-us")
        sf.write(str(wav), pcm, sr, subtype="FLOAT")
    pcm, sr = sf.read(str(wav), dtype="float32")
    keep = np.nonzero(np.abs(pcm) > 0.01)[0]  # 修掉頭尾靜音
    pcm = pcm[max(0, keep[0] - sr // 50): keep[-1] + sr // 10]
    if sr != SR:  # 頻域補零重取樣（24 kHz → 48 kHz）
        n = int(len(pcm) * SR / sr)
        pcm = np.fft.irfft(np.fft.rfft(pcm), n) * n / len(pcm)
    return (pcm * (0.1 / np.sqrt(np.mean(pcm ** 2)))).astype(np.float32)  # 每句 RMS 一致


def build_timeline():
    BUILD.mkdir(exist_ok=True)
    T, SC, clips, t = {}, {}, [], 0.0
    for key, _zh, _en, lines, lead, tail in SCENES:
        s0 = t
        t += lead
        for lid in lines:
            zh, en, say = LINE[lid]
            pcm = synth(lid, say or en)
            d = len(pcm) / SR
            T[lid] = (t, t + d)
            clips.append((t, pcm))
            t += d + GAP
        t += tail
        SC[key] = (s0, t)
    narr = np.zeros(int(t * SR) + SR, np.float32)
    for st, pcm in clips:
        i = int(st * SR)
        narr[i:i + len(pcm)] = pcm
    return T, SC, t, narr


T, SC, TOTAL = {}, {}, 0.0


def at(lid, frac=0.0, off=0.0):
    s, e = T[lid]
    return s + (e - s) * frac + off


# ---------------------------------------------------------------- 繪圖工具
def clamp(x, lo=0.0, hi=1.0):
    return max(lo, min(hi, x))


def ease(t, t0, d=0.5):
    p = clamp((t - t0) / d)
    return 1 - (1 - p) ** 3


def q(p, steps=20):
    return round(clamp(p) * steps) / steps


def blit(fr, spr, x, y, a=1.0):
    if a <= 0.01:
        return
    if a < 0.99:
        spr = spr.copy()
        spr.putalpha(spr.getchannel("A").point(lambda v: int(v * a)))
    fr.alpha_composite(spr, (int(x), int(y)))


TOK = re.compile(r"[A-Za-z0-9$][A-Za-z0-9.,%:/\-+~$'’]*|\s+|.")
NO_HEAD = set("，。、：；）」％！？")


def wrap(text, font, maxw):
    out = []
    for para in text.split("\n"):
        cur = ""
        for tok in TOK.findall(para):
            if not cur or font.getlength(cur + tok) <= maxw or tok in NO_HEAD:
                cur += tok
            else:
                out.append(cur.rstrip())
                cur = tok.lstrip()
        out.append(cur)
    return out


def check(d, x, y, s, color, width=4):
    d.line([(x, y + s * .55), (x + s * .38, y + s * .9), (x + s, y + s * .1)], fill=color, width=width,
           joint="curve")


def cross(d, x, y, s, color, width=4):
    d.line([(x, y), (x + s, y + s)], fill=color, width=width)
    d.line([(x + s, y), (x, y + s)], fill=color, width=width)


def chip(d, x, y, text, color, size=18, fill=None):
    f = F(size, "m")
    w = f.getlength(text) + 24
    d.rounded_rectangle((x, y, x + w, y + size + 14), 8, outline=color, width=2, fill=fill)
    d.text((x + 12, y + 5), text, font=f, fill=color)
    return w


def card(h, title, en, accent=BLUE, tag=None):
    img = Image.new("RGBA", (CARD_W, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((0, 0, CARD_W - 1, h - 1), 16, fill=CARD, outline=BORDER, width=2)
    d.rectangle((0, 18, 6, h - 18), fill=accent)
    d.text((28, 14), title, font=F(27, "s"), fill=TEXT)
    tw = F(27, "s").getlength(title)
    d.text((28 + tw + 14, 22), en, font=F(18), fill=MUTED)
    if tag:
        f = F(18, "m")
        chip(d, CARD_W - 28 - f.getlength(tag[0]) - 24, 16, tag[0], tag[1], 18)
    return img, d


# ---------------------------------------------------------------- 固定版面
def make_static():
    img = Image.new("RGBA", (W, H), BG + (255,))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(PHONE, 48, fill=(8, 8, 10), outline=(70, 78, 92), width=3)
    d.rounded_rectangle(SCREEN, 38, fill=CHAT_BG)
    d.rounded_rectangle((SCREEN[0], SCREEN[1], SCREEN[2], CHAT_TOP + 20), 38, fill=(250, 251, 253))
    d.rectangle((SCREEN[0], CHAT_TOP - 10, SCREEN[2], CHAT_TOP), fill=(250, 251, 253))
    d.line((SCREEN[0], CHAT_TOP, SCREEN[2], CHAT_TOP), fill=(214, 220, 230), width=2)
    d.ellipse((132, 122, 176, 166), fill=LINE_GREEN)
    d.text((154, 144), "夥", font=F(24, "s"), fill="white", anchor="mm")
    d.text((190, 120), "頭家夥伴", font=F(23, "s"), fill=(20, 24, 32))
    d.text((190, 148), "LINE 官方帳號 · HINO", font=F(15), fill=(110, 118, 130))
    d.rounded_rectangle(PANEL, 22, fill=PANEL_BG, outline=BORDER, width=2)
    return img


def chat_mask():
    m = Image.new("L", (SCREEN[2] - SCREEN[0], SCREEN[3] - CHAT_TOP - 1), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, -60, m.width - 1, m.height - 1), 38, fill=255)
    return m


# ---------------------------------------------------------------- 手機對話
VW = SCREEN[2] - SCREEN[0]
VH = SCREEN[3] - CHAT_TOP - 1
MSGS = []


def build_msgs():
    MSGS.clear()

    def add(kind, t0, text="", buttons=(), press=None):
        MSGS.append(dict(id=len(MSGS), kind=kind, t0=t0, text=text, buttons=list(buttons), press=press))

    add("date", at("rule", .3), "11/14（五）凌晨 02:10")
    add("bot", at("push1", .45),
        "【怠速提醒】ABC-6776\n最近每天怠速 6.1 小時，同型車平均 2.3 小時。\n"
        "大多在萬華一帶同一個停靠點，每次 5～8 小時沒熄火，估計每月多花約 3,000 元。\n要我幫你找改善方法嗎？",
        ["好，幫我看", "這台是正常的"], (1, at("push_back", 0, -0.2)))
    add("boss", at("push_back", .35), "那台是冷凍車啦，停著也要開冷凍機，不能熄火。")
    add("bot", at("experiment", .55),
        "了解，冷凍機一定要開。換個方向試試：\n① 這個點可能有一部分是在排隊等卸貨。下週照舊、再下週晚 1 小時到場，我來比較怠速。\n"
        "② 如果收貨方有插座，冷凍機改插電待機，每月最多可省約 100 公升。",
        ["開始實驗", "我再想想"], (0, at("options", 1, -0.3)))
    add("bot", at("options", 1, 0.9), "好，實驗 11/17 開始，12/1 告訴你結果。")
    add("date", at("service", 0, -1.6), "11/30（日）21:00")
    add("bot", at("service", .1),
        "【保養提醒】ABC-7610\n濾煙器（DPF）可能正在堵塞：11 月相關故障碼出現 15 次，而且越來越頻繁。\n"
        "建議兩週內進 HINO 服務廠檢查。及早處理通常比堵死後更換省錢，也少停駛幾天。",
        ["預約保養", "稍後提醒"], (0, at("service", 1, 0.2)))
    add("date", at("result", 0, -0.6), "12/1（一）08:00")
    add("bot", at("result", .25),
        "【實驗結果・示意】ABC-6776\n晚 1 小時到場的那週，每次停靠少怠速 1.8 小時。\n建議之後固定 9 點後到場。",
        ["採納", "不採納"], (0, at("result", 1, 0.1)))
    add("bot", at("net", 0, -0.4),
        "【11 月結算】\n估計省下 256 公升（7,654 元）\n只按保守下限 190 公升計費：\n5,681 元 × 20% ＝ 1,136 元\n"
        "你本月淨省 6,518 元")


BOT_W, BOSS_W = 330, 300
BF = 20
BLH = 30


@lru_cache(maxsize=None)
def msg_sprite(mid, pressed):
    m = MSGS[mid]
    f, fb = F(BF), F(BF, "s")
    if m["kind"] == "date":
        tw = F(16).getlength(m["text"]) + 28
        img = Image.new("RGBA", (VW, 40), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle(((VW - tw) / 2, 6, (VW + tw) / 2, 34), 14, fill=(120, 136, 160, 200))
        d.text((VW / 2, 20), m["text"], font=F(16), fill="white", anchor="mm")
        return img
    if m["kind"] == "boss":
        lines = wrap(m["text"], f, BOSS_W - 32)
        bw = min(BOSS_W, max(f.getlength(s) for s in lines) + 32)
        bh = len(lines) * BLH + 24
        img = Image.new("RGBA", (VW, bh + 8), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        x0 = VW - 18 - bw
        d.rounded_rectangle((x0, 0, VW - 18, bh), 18, fill=BOSS_BUBBLE)
        for i, s in enumerate(lines):
            d.text((x0 + 16, 12 + i * BLH), s, font=f, fill=(18, 22, 28))
        return img
    # bot
    lines = []
    for k, para in enumerate(m["text"].split("\n")):
        bold = para.startswith("【")
        lines += [(s, bold) for s in wrap(para, fb if bold else f, BOT_W - 32)]
    bw = min(BOT_W, max((fb if b else f).getlength(s) for s, b in lines) + 32)
    if m["buttons"]:
        bw = BOT_W
    bh = len(lines) * BLH + 24
    nb = len(m["buttons"])
    img = Image.new("RGBA", (VW, bh + nb * 50 + 10), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    x0 = 66
    d.ellipse((16, 0, 54, 38), fill=LINE_GREEN)
    d.text((35, 19), "夥", font=F(19, "s"), fill="white", anchor="mm")
    d.rounded_rectangle((x0, 0, x0 + bw, bh + (nb * 50 if nb else 0)), 18, fill="white")
    for i, (s, b) in enumerate(lines):
        d.text((x0 + 16, 12 + i * BLH), s, font=fb if b else f, fill=(18, 22, 28))
    for j, label in enumerate(m["buttons"]):
        y = bh + j * 50
        d.line((x0, y, x0 + bw, y), fill=(226, 230, 236), width=2)
        on = pressed == j
        if on:
            d.rectangle((x0 + 2, y + 2, x0 + bw - 2, y + 48), fill=(225, 248, 233))
            check(d, x0 + bw / 2 - F(BF, "s").getlength(label) / 2 - 30, y + 15, 20, LINE_GREEN, 3)
        d.text((x0 + bw / 2, y + 25), label, font=F(BF, "s"), fill=LINE_GREEN if (pressed is None or on)
               else (160, 168, 178), anchor="mm")
    return img


@lru_cache(maxsize=None)
def typing_sprite(phase):
    img = Image.new("RGBA", (VW, 52), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((16, 0, 54, 38), fill=LINE_GREEN)
    d.text((35, 19), "夥", font=F(19, "s"), fill="white", anchor="mm")
    d.rounded_rectangle((66, 0, 146, 46), 18, fill="white")
    for k in range(3):
        c = (60, 66, 76) if k == phase else (176, 182, 192)
        d.ellipse((84 + k * 18, 18, 94 + k * 18, 28), fill=c)
    return img


def draw_chat(fr, t, mask):
    layer = Image.new("RGBA", (VW, VH), CHAT_BG + (255,))
    items = []
    for m in MSGS:
        typing = m["kind"] == "bot"
        start = m["t0"] - (1.1 if typing else 0)
        if t < start:
            continue
        if typing and t < m["t0"]:
            spr = typing_sprite(int(t * 4) % 3)
            p = ease(t, start, .3)
            items.append((spr, spr.height * p, p))
            continue
        pressed = None
        if m["press"] and t >= m["press"][1]:
            pressed = m["press"][0]
        spr = msg_sprite(m["id"], pressed)
        p = ease(t, m["t0"], .35)
        th = 52 if typing else 0
        items.append((spr, th + (spr.height - th) * p if typing else spr.height * p, p if not typing else 1))
    y, pos = 16, []
    for spr, h, a in items:
        pos.append((spr, y, a))
        y += h + 14
    off = max(0, y - VH + 10)
    for spr, yy, a in pos:
        yy -= off
        if yy + spr.height < 0 or yy > VH:
            continue
        blit(layer, spr, 0, yy + (1 - a) * 18, a)
    fr.paste(layer, (SCREEN[0], CHAT_TOP + 1), mask)


# ---------------------------------------------------------------- 右側：流程列
NODES = [("車機數據", "Telematics"), ("規則監測", "Rules"), ("探索者", "Explorer"),
         ("調查員", "Investigator"), ("反駁者", "Skeptic"), ("LINE 推播", "Notify")]


@lru_cache(maxsize=None)
def pipeline_sprite(active):
    w, h = CARD_W, 82
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    nw, gap = 158, (w - 158 * 6) / 5
    for k, (zh, en) in enumerate(NODES):
        x = k * (nw + gap)
        on = k == active
        d.rounded_rectangle((x, 6, x + nw, h - 6), 14, fill=(37, 99, 235) if on else CARD,
                            outline=BLUE if on else BORDER, width=2)
        d.text((x + nw / 2, 30), zh, font=F(21, "s"), fill="white" if on else TEXT, anchor="mm")
        d.text((x + nw / 2, 56), en, font=F(15), fill="white" if on else MUTED, anchor="mm")
        if k < 5:
            ax = x + nw + 6
            d.line((ax, h / 2, ax + gap - 12, h / 2), fill=DIM, width=3)
            d.polygon([(ax + gap - 12, h / 2 - 6), (ax + gap - 4, h / 2), (ax + gap - 12, h / 2 + 6)], fill=DIM)
    return img


def node_schedule():
    return [(at("nightly"), 0), (at("rule"), 1), (at("invest"), 3), (at("push1", .4), 5),
             (at("constraint"), 3), (at("experiment", .5), 5), (at("explore"), 2), (at("skeptic"), 4),
             (at("dpf"), 3), (at("dpf", .6), 4), (at("service"), 5), (SC["settle"][0], -1)]


def active_node(t):
    a = -1
    for tt, k in node_schedule():
        if t >= tt:
            a = k
    return a


# ---------------------------------------------------------------- 右側：各場景卡片
# 王老闆車隊：每月油錢＝日均用油 × 出勤比例 × 365 ÷ 12 × 29.9 元/L（數據推算_每車.csv）
FLEET = [("ABC-6776", "700 系列", "冷凍車", 94517), ("ABC-5310", "500 系列", "", 29398),
         ("ABC-7610", "700 系列", "", 25465), ("ABC-7756", "500 系列", "", 23067),
         ("ABC-1999", "300 系列", "", 15236)]


@lru_cache(maxsize=64)
def c_fleet(p):
    img, d = card(340, "王老闆的車隊", "Mr. Wang's fleet: 5 HINO trucks", BLUE)
    hdr = F(17)
    d.text((30, 70), "車牌", font=hdr, fill=MUTED)
    d.text((230, 70), "車型", font=hdr, fill=MUTED)
    d.text((620, 70), "每月油錢（元）", font=hdr, fill=MUTED, anchor="ra")
    for k, (plate, typ, note, nt) in enumerate(FLEET):
        a = clamp(p * 6 - k)
        if a <= 0:
            continue
        y = 100 + k * 38
        col = tuple(int(c * a + CARD[i] * (1 - a)) for i, c in enumerate(TEXT))
        d.text((30, y), plate, font=F(22, "m"), fill=col)
        d.text((230, y), typ + ("  " + note if note else ""), font=F(20), fill=col)
        d.text((620, y), f"{nt:,.0f}", font=F(22, "m"), fill=col, anchor="ra")
        d.rounded_rectangle((650, y + 6, 650 + 420 * nt / 94517 * a, y + 26), 5, fill=BLUE)
    if p >= 0.95:
        d.line((30, 292, 1090, 292), fill=BORDER, width=1)
        d.text((30, 298), "合計", font=F(20, "s"), fill=TEXT)
        d.text((620, 298), "≈ 187,700 / 月", font=F(22, "s"), fill=AMBER, anchor="ra")
    return img


@lru_cache(maxsize=8)
def c_nightly(p):
    img, d = card(176, "每晚自動檢查", "Nightly automatic check", TEAL)
    d.text((30, 72), "每台車每天約 1,400 筆車機數據：位置、車速、累積油耗、轉速與負載、故障碼。", font=F(20), fill=TEXT)
    d.text((30, 108), "固定規則先跑，有異常才交給 LLM 調查：AI 與雲端成本約 NT$970／台／年。", font=F(20), fill=TEXT)
    d.text((30, 140), "Rules run first; the LLM is called only when something looks off.", font=F(17), fill=MUTED)
    return img


@lru_cache(maxsize=32)
def c_rule(p):
    img, d = card(180, "規則監測", "Rule-based monitor · no LLM", RED, ("同型車的 2.6 倍", RED))
    d.text((30, 64), "每日怠速時數 Idle hours per day", font=F(17), fill=MUTED)
    rows = [("ABC-6776", 6.1, RED), ("同型車匿名基準", 2.3, DIM)]
    for k, (lab, v, col) in enumerate(rows):
        y = 94 + k * 40
        d.text((30, y), lab, font=F(20, "m"), fill=TEXT)
        L = 760 * v / 6.5 * p
        d.rounded_rectangle((230, y + 2, 230 + max(L, 4), y + 28), 6, fill=col)
        d.text((240 + L, y), f"{v * p:.1f} 小時", font=F(20, "s"), fill=TEXT)
    return img


# 官方原始數據：ABC-6776 在 25.02°N, 121.50°E 附近的長時間怠速
STOPS = [("3/11", 10.27, 16.44, 6.2), ("7/1", 8.20, 16.44, 8.2), ("7/31", 8.70, 16.04, 7.3),
         ("11/1", 9.58, 16.55, 7.0)]


@lru_cache(maxsize=32)
def c_invest(p):
    img, d = card(268, "調查員 Agent", "Investigator: query idle stops", AMBER)
    chip(d, 30, 62, "同一地點：萬華一帶（25.02°N, 121.50°E）", AMBER, 18)
    x0, x1 = 160, 860
    for hh in (6, 9, 12, 15, 18):
        x = x0 + (hh - 6) / 12 * (x1 - x0)
        d.text((x, 108), f"{hh:02d}:00", font=F(15), fill=MUTED, anchor="ma")
        d.line((x, 130, x, 254), fill=(40, 54, 78), width=1)
    for k, (day, a, b, hrs) in enumerate(STOPS):
        al = clamp(p * 5 - k)
        if al <= 0:
            continue
        y = 136 + k * 30
        d.text((30, y), day, font=F(19, "m"), fill=TEXT)
        xa = x0 + (a - 6) / 12 * (x1 - x0)
        xb = x0 + (b - 6) / 12 * (x1 - x0)
        d.rounded_rectangle((xa, y + 4, xa + (xb - xa) * al, y + 22), 5, fill=AMBER)
        d.text((890, y), f"怠速 {hrs} 小時", font=F(19, "m"), fill=TEXT)
    return img


@lru_cache(maxsize=8)
def c_impact(p):
    img, d = card(150, "影響估算", "Impact estimate", GREEN)
    d.text((30, 66), "怠速油耗 1.9 公升/小時（實測）", font=F(20), fill=TEXT)
    d.text((30, 100), "這個停靠點每月約 54 小時 → 約 100 公升", font=F(20), fill=TEXT)
    chip(d, 520, 100, "可信度：高・直接量測", GREEN, 16)
    d.text((1090, 50), "NT$3,000", font=F(54, "s"), fill=AMBER, anchor="ra")
    d.text((1090, 116), "每月多花 / month", font=F(18), fill=MUTED, anchor="ra")
    return img


@lru_cache(maxsize=8)
def c_constraint(p):
    img, d = card(212, "重新建模", "Re-model with the owner's constraint", TEAL)
    check(d, 30, 66, 22, GREEN)
    d.text((64, 62), "限制條件：冷凍機必須運轉（老闆提供）", font=F(21, "m"), fill=GREEN)
    d.text((30, 102), "每次停靠 5～8 小時，其中：", font=F(18), fill=MUTED)
    x0, x1, y = 30, 1090, 132
    xm = x0 + (x1 - x0) * .55
    d.rounded_rectangle((x0, y, xm - 3, y + 40), 8, fill=(20, 120, 110))
    d.rounded_rectangle((xm + 3, y, x1, y + 40), 8, fill=(150, 110, 20))
    d.text(((x0 + xm) / 2, y + 20), "冷藏必要  Must keep cold  ？", font=F(19, "m"), fill="white", anchor="mm")
    d.text(((xm + x1) / 2, y + 20), "排隊等卸貨  Waiting to unload  ？", font=F(19, "m"), fill="white",
           anchor="mm")
    d.text((30, 180), "數據分不出兩者 → 需要實驗  Data can't separate them → run an experiment", font=F(17),
           fill=AMBER)
    return img


@lru_cache(maxsize=8)
def c_experiment(p):
    img, d = card(196, "消融實驗", "Two-week ablation experiment", BLUE)
    for k, (wk, rng, how, col) in enumerate([("第 1 週", "11/17–11/23", "照舊到場", DIM),
                                             ("第 2 週", "11/24–11/30", "晚 1 小時到場", BLUE)]):
        x = 30 + k * 540
        d.rounded_rectangle((x, 64, x + 520, 132), 12, outline=col, width=3)
        d.text((x + 20, 72), f"{wk}  {rng}", font=F(18), fill=MUTED)
        d.text((x + 20, 98), how, font=F(23, "s"), fill=TEXT)
    d.text((30, 146), "比較：每次停靠的怠速時數。其他 4 台車同期當對照，扣掉天氣與淡旺季。", font=F(19), fill=TEXT)
    return img


@lru_cache(maxsize=8)
def c_options(p):
    img, d = card(178, "方案比較", "Options for the owner", GREEN)
    cols = (30, 330, 640)
    for x, h in zip(cols, ("方案", "需要什麼", "每月可省")):
        d.text((x, 62), h, font=F(17), fill=MUTED)
    rows = [("錯開到場時間", "不用花錢", "實驗後確定"),
            ("冷凍機插電待機", "收貨方提供插座", "最多約 100 L ≈ NT$3,000")]
    for k, r in enumerate(rows):
        y = 92 + k * 38
        for x, s in zip(cols, r):
            d.text((x, y), s, font=F(21, "m" if x == 30 else "r"), fill=TEXT)
    return img


SIGNALS = ["怠速", "油耗", "故障碼", "油量", "車速", "引擎負載", "水溫", "轉速", "停靠點"]


@lru_cache(maxsize=16)
def c_explore(k):
    img, d = card(108, "每週探索", "Open exploration · no preset rules", BLUE)
    x = 30
    for j, s in enumerate(SIGNALS):
        on = j == k
        x += chip(d, x, 62, s, BLUE if on else MUTED, 17, fill=(30, 64, 120) if on else None) + 12
    return img


# 官方原始數據：ABC-5310 於 2025/5/1 15:07–15:44 每分鐘最後一筆油量讀數（%）
FUEL_LVL = [50, 50, 55, 55, 25, 25, 60, 25, 40, 45, 50, 55, 40, 40, 35, 40, 50, 40, 40, 50, 50, 35, 40, 40, 50,
            45, 50, 45, 50, 50, 50, 50, 50, 50, 50, 50, 50, 45]


@lru_cache(maxsize=32)
def c_theft(p, verdict):
    img, d = card(236, "線索：ABC-5310 被偷油？", "Lead: fuel theft?", RED if not verdict else DIM)
    x0, x1, y0, y1 = 70, 600, 70, 210
    for v in (25, 50, 75):
        y = y1 - (y1 - y0) * v / 100
        d.line((x0, y, x1, y), fill=(40, 54, 78), width=1)
        d.text((x0 - 10, y), f"{v}%", font=F(14), fill=MUTED, anchor="rm")
    n = len(FUEL_LVL)
    pts = [(x0 + (x1 - x0) * i / (n - 1), y1 - (y1 - y0) * v / 100) for i, v in enumerate(FUEL_LVL)]
    m = max(2, int(n * p))
    d.line(pts[:m], fill=BLUE, width=3)
    d.text((x0, y1 + 6), "15:07", font=F(14), fill=MUTED)
    d.text((x1, y1 + 6), "15:44", font=F(14), fill=MUTED, anchor="ra")
    d.text((640, 66), "油量讀數 1 分鐘內 55% → 25%", font=F(21, "m"), fill=TEXT)
    d.text((640, 98), "同期累積油耗只增加 2 公升", font=F(21, "m"), fill=TEXT)
    if verdict:
        d.rounded_rectangle((640, 136, 1090, 214), 12, outline=RED, width=3)
        cross(d, 660, 152, 20, RED)
        d.text((694, 144), "反駁者：讀數馬上跳回 → 感測器跳動", font=F(19, "m"), fill=TEXT)
        d.text((694, 178), "線索排除・不通知老闆", font=F(19, "s"), fill=RED)
    return img


# 官方原始數據：ABC-7610 的 DPF 故障碼（抽樣日）
DTC = [("3/31", "P2002", 1), ("7/31", "P2002", 1), ("11/1", "P244B", 2), ("11/21", "P244B", 9), ("11/30", "P244B", 4)]


@lru_cache(maxsize=32)
def c_dpf(p):
    img, d = card(270, "發現：ABC-7610 濾煙器（DPF）", "Finding: DPF clogging trend", RED, ("可信度：中高", AMBER))
    x0, x1, yb = 90, 600, 222
    d.line((x0 - 50, yb, x1 + 50, yb), fill=BORDER, width=2)
    for k, (day, code, n) in enumerate(DTC):
        x = x0 + (x1 - x0) * k / (len(DTC) - 1)
        d.text((x, yb + 8), day, font=F(15), fill=MUTED, anchor="ma")
        a = clamp(p * 6 - k)
        if a <= 0:
            continue
        hgt = 13 * n * a
        col = AMBER if code == "P2002" else RED
        d.rounded_rectangle((x - 18, yb - hgt - 2, x + 18, yb), 4, fill=col)
        d.text((x, yb - hgt - 26), f"{n} 次", font=F(16, "s"), fill=TEXT, anchor="ma")
    d.rounded_rectangle((40, 66, 52, 78), 2, fill=AMBER)
    d.text((60, 62), "P2002 效率低", font=F(15), fill=MUTED)
    d.rounded_rectangle((180, 66, 192, 78), 2, fill=RED)
    d.text((200, 62), "P244B 壓差", font=F(15), fill=MUTED)
    tx = 690
    d.text((tx, 70), "11 月出現 15 次，而且越來越頻繁", font=F(20, "m"), fill=TEXT)
    d.text((tx, 104), "怠速用油佔 24%，車隊最高", font=F(20, "m"), fill=TEXT)
    d.text((tx, 138), "長時間低負載運轉，濾煙器較難再生", font=F(18), fill=MUTED)
    d.text((tx, 176), "反駁者：兩種碼偵測不同項目，", font=F(18), fill=MUTED)
    d.text((tx, 202), "先後出現 → 結論維持，建議進廠確認", font=F(18), fill=MUTED)
    return img


@lru_cache(maxsize=32)
def c_result(p):
    img, d = card(186, "實驗結果", "Experiment result", BLUE, ("示意 Illustrative", AMBER))
    rows = [("第 1 週：照舊", 6.4, DIM), ("第 2 週：晚 1 小時到", 4.6, BLUE)]
    for k, (lab, v, col) in enumerate(rows):
        y = 70 + k * 44
        d.text((30, y), lab, font=F(20, "m"), fill=TEXT)
        L = 600 * v / 7 * p
        d.rounded_rectangle((280, y + 2, 280 + max(L, 4), y + 30), 6, fill=col)
        d.text((290 + L, y), f"{v * p:.1f} 小時/次", font=F(20, "s"), fill=TEXT)
    if p > .9:
        d.text((1090, 150), "每次少怠速 1.8 小時", font=F(22, "s"), fill=GREEN, anchor="ra")
    return img


@lru_cache(maxsize=32)
def c_billing(p):
    img, d = card(262, "結算方式", "How savings are billed", GREEN)
    rules = ["只算老闆按過「採納」的建議", "以匿名同型車為對照，扣掉季節、油價、路況", "按信賴區間下限計費，不確定的部分算老闆的"]
    for k, s in enumerate(rules):
        if p * 4 - k <= 0:
            continue
        y = 64 + k * 34
        check(d, 30, y + 4, 18, GREEN, 3)
        d.text((60, y), s, font=F(20), fill=TEXT)
    if p > .75:
        x0, x1, y = 60, 1060, 212
        sx = lambda v: x0 + (x1 - x0) * v / 350
        d.line((x0, y, x1, y), fill=BORDER, width=2)
        for v in (0, 100, 200, 300):
            d.text((sx(v), y + 8), f"{v} L", font=F(13), fill=MUTED, anchor="ma")
        d.rounded_rectangle((sx(190), y - 12, sx(320), y + 12), 8, fill=(30, 90, 60))
        d.ellipse((sx(256) - 8, y - 8, sx(256) + 8, y + 8), fill=GREEN)
        d.text((sx(256), y - 40), "估計 256 L", font=F(16, "m"), fill=TEXT, anchor="ma")
        d.line((sx(190), y - 22, sx(190), y + 16), fill=AMBER, width=3)
        d.text((sx(190) - 10, y - 40), "按 190 L 計費", font=F(16, "s"), fill=AMBER, anchor="ra")
        d.text((sx(320) + 12, y - 10), "80% 信賴區間", font=F(14), fill=MUTED)
    return img


@lru_cache(maxsize=8)
def c_net(p):
    img, d = card(150, "王老闆本月淨省", "Mr. Wang's net saving this month", GREEN, ("示意", AMBER))
    d.text((30, 62), "NT$6,518", font=F(58, "s"), fill=GREEN)
    d.text((1090, 70), "和泰分成 NT$1,136（190 L × 29.9 元 × 20%）", font=F(19), fill=TEXT, anchor="ra")
    d.text((1090, 104), "依試算表基準情境：可省 426 L/月 × 採納率 60%", font=F(17), fill=MUTED, anchor="ra")
    return img


def panel_cards(key, t):
    """回傳 [(sprite, 出現時間)]，由上往下排。"""
    if key == "fleet":
        return [(c_fleet(q(ease(t, at("fleet", .1), 2.0))), at("fleet", .1)),
                (c_nightly(1), at("nightly", .2))]
    if key == "detect":
        return [(c_rule(q(ease(t, at("rule", .15), 1.2))), at("rule", .1)),
                (c_invest(q(ease(t, at("invest", .1), 2.2))), at("invest", .05)),
                (c_impact(1), at("push1"))]
    if key == "push":
        return [(c_constraint(1), at("constraint")),
                (c_experiment(1), at("experiment", .3)),
                (c_options(1), at("options"))]
    if key == "explore":
        k = int((t - at("explore")) * 3) % len(SIGNALS) if t < at("theft") else 3
        return [(c_explore(k), at("explore")),
                (c_theft(q(ease(t, at("theft", .05), 1.8)), t >= at("skeptic", .3)), at("theft")),
                (c_dpf(q(ease(t, at("dpf", .1), 2.0))), at("dpf"))]
    if key == "settle":
        return [(c_result(q(ease(t, at("result", .2), 1.2))), at("result", .1)),
                (c_billing(q(ease(t, at("billing"), 3.0))), at("billing")),
                (c_net(1), at("net", .1))]
    return []


FOOT = {
    "fleet": "數據：和泰 AI 黑客松 HINO 官方車聯網數據（20 台車、2025 年抽樣日）。每月金額依抽樣日推估，柴油 29.9 元/L。",
    "detect": "同型車匿名基準：Demo 以比賽數據中其他 700 系列車輛模擬。停靠紀錄為官方數據原始值。",
    "push": "實驗與方案為 AI 依數據提出的建議；插電待機需冷凍機支援電源待機。",
    "explore": "油量讀數與故障碼皆為官方數據原始值；故障碼只涵蓋抽樣日，實際次數可能更多。",
    "settle": "標示「示意」者為情境模擬；金額依試算表（HINO省油分成試算.xlsx）基準情境。",
}


# ---------------------------------------------------------------- 畫面組合
def scene_at(t):
    for key, *_ in SCENES:
        if SC[key][0] <= t < SC[key][1]:
            return key
    return SCENES[-1][0]


def draw_header(fr, key, t):
    d = ImageDraw.Draw(fr)
    meta = {k: (zh, en) for k, zh, en, *_ in SCENES}
    zh, en = meta[key]
    a = ease(t, SC[key][0], .5)
    col = tuple(int(c * a + BG[i] * (1 - a)) for i, c in enumerate(TEXT))
    d.text((100, 26), zh, font=F(34, "s"), fill=col)
    d.text((100 + F(34, "s").getlength(zh) + 18, 40), en, font=F(20), fill=MUTED)
    d.text((1840, 30), "頭家夥伴", font=F(26, "s"), fill=TEXT, anchor="ra")
    d.text((1840, 64), "Boss Buddy · HINO x GenAI", font=F(15), fill=MUTED, anchor="ra")


def draw_panel(fr, key, t):
    d = ImageDraw.Draw(fr)
    d.text((PANEL[0] + 22, PANEL[1] + 14), "AI 背後在做什麼", font=F(20, "s"), fill=MUTED)
    d.text((PANEL[0] + 22 + F(20, "s").getlength("AI 背後在做什麼") + 12, PANEL[1] + 19),
           "What the AI is doing behind the scenes", font=F(15), fill=DIM)
    if t >= at("nightly"):
        blit(fr, pipeline_sprite(active_node(t)), CARD_X, PANEL[1] + 46 + 0, ease(t, at("nightly"), .6))
    y = 236
    for spr, t0 in panel_cards(key, t):
        if t >= t0:
            a = ease(t, t0, .5)
            blit(fr, spr, CARD_X, y + (1 - a) * 24, a)
        y += spr.height + CARD_GAP
    if key in FOOT:
        a = ease(t, SC[key][0] + .5, .8)
        if a > 0:
            col = tuple(int(c * a + PANEL_BG[i] * (1 - a)) for i, c in enumerate(DIM))
            d.text((CARD_X, PANEL[3] - 30), FOOT[key], font=F(15), fill=col)


def draw_title(fr, t):
    d = ImageDraw.Draw(fr)
    a = ease(t, 0.2, 1.0)
    col = lambda c: tuple(int(v * a + BG[i] * (1 - a)) for i, v in enumerate(c))
    d.text((W / 2, 330), "頭家夥伴", font=F(120, "s"), fill=col(TEXT), anchor="mm")
    d.text((W / 2, 440), "Boss Buddy", font=F(46, "m"), fill=col(BLUE), anchor="mm")
    d.text((W / 2, 530), "HINO 商用車頭家的 AI 省油顧問", font=F(34), fill=col(TEXT), anchor="mm")
    d.text((W / 2, 580), "An AI fuel-saving advisor for HINO truck owners", font=F(24), fill=col(MUTED), anchor="mm")
    b = ease(t, 1.4, .8)
    x = W / 2 - 520
    for k, (zh, en) in enumerate([("自己看數據", "Reads the data"), ("自己找問題", "Finds problems"),
                                  ("主動告訴你", "Tells you first")]):
        bx = x + k * 360
        c2 = tuple(int(v * b + BG[i] * (1 - b)) for i, v in enumerate(BORDER))
        d.rounded_rectangle((bx, 660, bx + 320, 760), 18, outline=c2, width=2)
        d.text((bx + 160, 692), zh, font=F(30, "s"), fill=tuple(int(v * b + BG[i] * (1 - b))
                                                                  for i, v in enumerate(TEXT)), anchor="mm")
        d.text((bx + 160, 734), en, font=F(18), fill=tuple(int(v * b + BG[i] * (1 - b))
                                                            for i, v in enumerate(MUTED)), anchor="mm")
    d.text((W / 2, 860), "和泰 AI 黑客松 2026 · HINO x GenAI", font=F(20), fill=col(DIM), anchor="mm")


def draw_close(fr, t):
    d = ImageDraw.Draw(fr)
    t0 = SC["close"][0]
    d.text((W / 2, 170), "頭家夥伴", font=F(80, "s"), fill=TEXT, anchor="mm")
    d.text((W / 2, 250), "Boss Buddy", font=F(34, "m"), fill=BLUE, anchor="mm")
    items = [("自己找問題", "Finds problems on its own", "規則監測 ＋ 調查員 ＋ 每週探索", GREEN),
             ("聽得懂限制，也會推翻自己", "Respects constraints, drops weak leads", "反駁者把關，結論標示可信度", AMBER),
             ("省下來才收費", "Paid only from savings", "按信賴區間下限分成 20%", BLUE)]
    for k, (zh, en, sub, col) in enumerate(items):
        a = ease(t, t0 + .4 + k * .5, .6)
        if a <= 0:
            continue
        bx = 160 + k * 540
        spr = Image.new("RGBA", (500, 250), (0, 0, 0, 0))
        sd = ImageDraw.Draw(spr)
        sd.rounded_rectangle((0, 0, 499, 249), 22, fill=CARD, outline=col, width=3)
        sd.text((250, 70), zh, font=F(32, "s"), fill=TEXT, anchor="mm")
        sd.text((250, 122), en, font=F(20), fill=MUTED, anchor="mm")
        sd.text((250, 186), sub, font=F(20, "m"), fill=col, anchor="mm")
        blit(fr, spr, bx, 360 + (1 - a) * 30, a)
    if t >= at("next"):
        a = ease(t, at("next"), .6)
        spr = Image.new("RGBA", (1100, 90), (0, 0, 0, 0))
        sd = ImageDraw.Draw(spr)
        sd.rounded_rectangle((0, 0, 1099, 89), 20, fill=(30, 50, 90))
        sd.text((550, 30), "下一階段：行車安全（分心、疲勞、車道偏移事件）", font=F(28, "s"), fill=TEXT, anchor="mm")
        sd.text((550, 66), "Next: driving safety with the same engine", font=F(18), fill=MUTED, anchor="mm")
        blit(fr, spr, (W - 1100) / 2, 690, a)


def draw_subs(fr, t):
    for lid, (s, e) in T.items():
        if s - 0.1 <= t < e + GAP * .8:
            zh, en, _ = LINE[lid]
            fz, fe = F(31, "m"), F(23)
            zl, el = wrap(zh, fz, 1680), wrap(en, fe, 1680)
            hgt = len(zl) * 44 + len(el) * 32 + 22
            y0 = H - 10 - hgt
            ov = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            od = ImageDraw.Draw(ov)
            od.rounded_rectangle((100, y0, W - 100, H - 10), 14, fill=(0, 0, 0, 190))
            y = y0 + 10
            for s_ in zl:
                od.text((W / 2, y), s_, font=fz, fill="white", anchor="ma")
                y += 44
            for s_ in el:
                od.text((W / 2, y), s_, font=fe, fill=(210, 220, 232), anchor="ma")
                y += 32
            fr.alpha_composite(ov)
            return


def fade(fr, a):
    if a > 0.01:
        fr.alpha_composite(Image.new("RGBA", (W, H), BG + (int(255 * clamp(a)),)))


def render(t, static, mask):
    key = scene_at(t)
    if key == "title":
        fr = Image.new("RGBA", (W, H), BG + (255,))
        draw_title(fr, t)
        fade(fr, (t - (SC["title"][1] - .5)) / .5)
    elif key == "close":
        fr = Image.new("RGBA", (W, H), BG + (255,))
        draw_close(fr, t)
        fade(fr, 1 - (t - SC["close"][0]) / .6)
        fade(fr, (t - (TOTAL - .8)) / .8)
    else:
        fr = static.copy()
        draw_header(fr, key, t)
        draw_chat(fr, t, mask)
        draw_panel(fr, key, t)
        fade(fr, 1 - (t - SC["fleet"][0]) / .6)
        fade(fr, (t - (SC["settle"][1] - .5)) / .5)
    draw_subs(fr, t)
    return fr


# ---------------------------------------------------------------- 音效時間點
def reveal(t0, d, mul, k):
    """卡片用 q(ease(t, t0, d)) * mul - k > 0 逐項出現；回傳第 k 項出現的時間。"""
    for i in range(int(d * FPS) + 2):
        t = t0 + i / FPS
        if q(ease(t, t0, d)) * mul - k > 0:
            return t
    return t0 + d


def sound_events():
    """[(秒數, 音效, 參數)]：手機在左（pan<0）、右側面板在右（pan>0）。"""
    ev = [(0.25, "hit", {}), (SC["close"][0] + 0.3, "hit", {})]
    ev += [(SC[k][0] - 0.35, "whoosh", {}) for k, *_ in SCENES if k != "title"]
    for m in MSGS:
        if m["kind"] == "bot":
            ev.append((m["t0"], "pop", {"pan": -.45}))
        elif m["kind"] == "boss":
            ev.append((m["t0"], "sent", {"pan": -.3}))
        if m["press"]:
            ev.append((m["press"][1], "tap", {"pan": -.4}))
    special = {at("rule", .1): "alert", at("dpf"): "found"}
    for key, *_ in SCENES:
        if key in ("title", "close"):
            continue
        for _spr, t0 in panel_cards(key, SC[key][0] + 0.01):
            ev.append((t0, special.get(t0, "blip"), {"pan": .45}))
    ev += [(reveal(at("fleet", .1), 2.0, 6, k), "tick", {"pan": .4, "pitch": 1 + k * .06}) for k in range(5)]
    ev += [(reveal(at("dpf", .1), 2.0, 6, k), "tick", {"pan": .4, "pitch": 1 + k * .06}) for k in range(5)]
    ev += [(reveal(at("billing"), 3.0, 4, k), "tick", {"pan": .4, "pitch": 1.2}) for k in range(3)]
    ev += [(at("skeptic", .3), "reject", {"pan": .35}), (at("net", .1), "coin", {"pan": .3})]
    ev += [(tt, "blip", {"pan": .2, "gain": .7}) for tt, k in node_schedule() if k >= 0]
    return ev


def write_srt():
    def ts(x):
        ms = int(round(x * 1000))
        return f"{ms // 3600000:02d}:{ms // 60000 % 60:02d}:{ms // 1000 % 60:02d},{ms % 1000:03d}"
    rows = []
    for k, (lid, (s, e)) in enumerate(sorted(T.items(), key=lambda kv: kv[1][0]), 1):
        zh, en, _ = LINE[lid]
        rows.append(f"{k}\n{ts(s)} --> {ts(e + GAP * .8)}\n{zh}\n{en}\n")
    SRT.write_text("\n".join(rows), encoding="utf-8")


def main():
    global T, SC, TOTAL
    ap = argparse.ArgumentParser()
    ap.add_argument("--stills", help="以逗號分隔的秒數，只輸出這些時間點的畫面")
    ap.add_argument("--audio", action="store_true", help="只輸出混音 _build/mix.wav")
    args = ap.parse_args()
    T, SC, TOTAL, narr = build_timeline()
    build_msgs()
    static, mask = make_static(), chat_mask()
    print(f"總長 {TOTAL:.1f} 秒；場景：" + "、".join(f"{k} {a:.1f}–{b:.1f}" for k, (a, b) in SC.items()))
    if args.stills:
        for s in args.stills.split(","):
            render(float(s), static, mask).convert("RGB").save(BUILD / f"still_{float(s):06.1f}.png")
        return
    audio = soundtrack.mix(narr, sound_events(), TOTAL, {"drums": SC["fleet"][0], "end": SC["close"][0]})
    sf.write(str(BUILD / "mix.wav"), audio, SR, subtype="PCM_16")
    if args.audio:
        return
    write_srt()
    n = int(TOTAL * FPS)
    ff = subprocess.Popen(
        ["ffmpeg", "-y", "-loglevel", "error", "-f", "rawvideo", "-pix_fmt", "rgb24", "-s", f"{W}x{H}",
         "-r", str(FPS), "-i", "-", "-i", str(BUILD / "mix.wav"), "-c:v", "libx264", "-preset", "medium",
         "-crf", "18", "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "192k", "-shortest",
         "-movflags", "+faststart", str(OUT)], stdin=subprocess.PIPE)
    for i in range(n):
        ff.stdin.write(render(i / FPS, static, mask).convert("RGB").tobytes())
        if i % (FPS * 10) == 0:
            print(f"  {i / FPS:5.0f}s / {TOTAL:.0f}s", flush=True)
    ff.stdin.close()
    ff.wait()
    print("完成：", OUT)


if __name__ == "__main__":
    main()
