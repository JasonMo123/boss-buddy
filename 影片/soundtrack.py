"""Demo 影片的背景音樂與音效：全部用 numpy 即時合成，不使用外部音檔（沒有授權問題）。

make_demo.py 呼叫 mix(narration, events, total, marks)：
    narration  單聲道 float32 旁白（SR 取樣率）
    events     [(秒數, 音效名稱, 參數 dict)]，音效名稱見 SFX
    marks      {"drums": 鼓組開始秒數, "end": 收尾和弦開始秒數}
回傳立體聲 float32 陣列 (n, 2)，峰值 ≤ -1 dBFS。
"""

from functools import lru_cache

import numpy as np

SR = 48000
RNG = np.random.default_rng(7)

BPM = 92
BEAT = 60 / BPM
BAR = 4 * BEAT
PROG = [[53, 57, 60, 64], [55, 59, 62, 67], [52, 55, 59, 64], [57, 60, 64, 69]]  # F G Em Am
ROOT = [41, 43, 40, 45]
ENDING = [([53, 57, 60, 64], 41), ([55, 59, 62, 65], 43), ([48, 55, 60, 64, 67], 36)]  # F G7 C


def hz(n):
    return 440.0 * 2 ** ((n - 69) / 12)


def tt(dur):
    return np.arange(int(dur * SR)) / SR


def place(bus, t0, sig, pan=0.0, gain=1.0):
    """把單聲道 sig 放到立體聲 bus 的 t0 秒，pan -1（左）～ 1（右），等功率聲像。"""
    i = int(t0 * SR)
    if i >= len(bus):
        return
    sig = sig[: len(bus) - i] * gain
    a = (pan + 1) * np.pi / 4
    bus[i:i + len(sig), 0] += sig * np.cos(a)
    bus[i:i + len(sig), 1] += sig * np.sin(a)


def onepole_lp(x, cutoff):
    """一階低通；cutoff 可為常數或逐樣本陣列（Hz）。"""
    c = np.broadcast_to(np.asarray(cutoff, float), x.shape)
    k = 1 - np.exp(-2 * np.pi * c / SR)
    y, s = np.empty_like(x), 0.0
    for i in range(len(x)):
        s += k[i] * (x[i] - s)
        y[i] = s
    return y


def fft_convolve(x, ir):
    n = len(x) + len(ir) - 1
    m = 1 << (n - 1).bit_length()
    return np.fft.irfft(np.fft.rfft(x, m) * np.fft.rfft(ir, m), m)[: len(x)]


def reverb(bus, sec=2.2, wet=0.28):
    t = tt(sec)
    out = bus.copy()
    for ch in range(2):
        ir = RNG.standard_normal(len(t)) * np.exp(-t * 6.9 / sec)
        ir = np.convolve(ir, np.ones(24) / 24, "same")  # 去掉刺耳高頻
        ir /= np.sqrt(np.sum(ir ** 2))
        out[:, ch] += wet * fft_convolve(bus[:, ch], ir)
    return out


# ---------------------------------------------------------------- 樂器
def pad_note(f, dur, rel=1.2):
    t = tt(dur + rel)
    sig = np.zeros_like(t)
    for det, ph in ((0.997, 0.0), (1.0, 1.1), (1.003, 2.3)):
        for h in range(1, 7):
            sig += np.sin(2 * np.pi * f * det * h * t + ph * h) / h ** 1.8
    env = np.minimum(t / 0.7, 1.0) * np.where(t > dur, np.exp(-(t - dur) / (rel / 3)), 1.0)
    env *= 1 + 0.08 * np.sin(2 * np.pi * 0.23 * t)
    return sig * env / 3


def pluck(f, decay=0.32):
    t = tt(decay * 5)
    sig = np.sin(2 * np.pi * f * t) + 0.35 * np.sin(4 * np.pi * f * t) * np.exp(-t / 0.08) \
        + 0.12 * np.sin(6 * np.pi * f * t) * np.exp(-t / 0.04)
    return sig * np.minimum(t / 0.004, 1) * np.exp(-t / decay)


def bass_note(f, dur):
    t = tt(dur + 0.3)
    sig = np.sin(2 * np.pi * f * t) + 0.22 * np.sin(4 * np.pi * f * t)
    return sig * np.minimum(t / 0.012, 1) * np.exp(-t / 0.55) * np.where(t > dur, np.exp(-(t - dur) / .06), 1)


@lru_cache(maxsize=None)
def kick(var=0):
    t = tt(0.45)
    f = 45 + 75 * np.exp(-t / 0.035)
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.16)


@lru_cache(maxsize=None)
def hat(var=0):
    t = tt(0.08)
    n = RNG.standard_normal(len(t))
    n = n - onepole_lp(n, 6000)
    return n * np.exp(-t / 0.018)


@lru_cache(maxsize=None)
def clap(var=0):
    t = tt(0.25)
    n = RNG.standard_normal(len(t))
    n = onepole_lp(n - onepole_lp(n, 900), 5000)
    return n * np.exp(-t / 0.06)


def bell(f, decay=0.9):
    t = tt(decay * 4)
    sig = np.zeros_like(t)
    for ratio, amp, dk in ((1, 1, 1), (2.76, .45, .55), (5.4, .2, .3), (8.93, .08, .2)):
        sig += amp * np.sin(2 * np.pi * f * ratio * t) * np.exp(-t / (decay * dk))
    return sig * np.minimum(t / 0.002, 1)


# ---------------------------------------------------------------- 背景音樂
def music(total, drums_from, end_from):
    n = int((total + 1) * SR)
    pads, keys, low, drums = (np.zeros((n, 2)) for _ in range(4))

    bars = []
    t, k = 0.0, 0
    end_start = max(BAR, (int((end_from - 2 * BAR) / BAR) + 1) * BAR)
    while t < end_start - 1e-6:
        bars.append((t, PROG[k % 4], ROOT[k % 4], BAR, t >= drums_from - BAR / 2))
        t += BAR
        k += 1
    for j, (chord, root) in enumerate(ENDING):
        dur = BAR if j < 2 else total - t
        bars.append((t, chord, root, dur, j < 2))
        t += BAR

    for b, (t0, chord, root, dur, dr) in enumerate(bars):
        last = dur > BAR + 1e-6
        for m, note in enumerate(chord):
            place(pads, t0, pad_note(hz(note), dur), pan=(-.5, .5)[m % 2] * .6)
        # 琶音：八分音符，最後的和弦逐漸放慢
        pattern = [0, 1, 2, 3, 2, 1, 2, 3]
        steps = int(dur / (BEAT / 2)) if not last else int(dur / BEAT)
        for s in range(steps):
            st = s * BEAT / 2 if not last else s * BEAT
            if last and st > dur - 2.5:
                break
            note = chord[pattern[s % 8] % len(chord)] + 12
            vel = (1.0, .6)[s % 2] * (1 - .5 * st / dur if last else 1)
            place(keys, t0 + st, pluck(hz(note)), pan=(-.35, .35)[s % 2], gain=vel)
        place(low, t0, bass_note(hz(root), min(dur, BEAT * 2.4) if last else BEAT * 1.4))
        if not last:
            place(low, t0 + BEAT * 2.5, bass_note(hz(root), BEAT * 1.2), gain=.8)
        if dr and t0 >= drums_from - BAR / 2:
            for bt in range(4):
                if bt in (0, 2):
                    place(drums, t0 + bt * BEAT, kick(), gain=.9)
                if bt in (1, 3):
                    place(drums, t0 + bt * BEAT, clap(b % 2), gain=.35, pan=.1)
                place(drums, t0 + (bt + .5) * BEAT, hat(bt % 3), gain=.22 * (1.15, .85)[bt % 2], pan=.4)

    # 琶音的乒乓延遲
    d = int(BEAT * .75 * SR)
    echo = np.zeros_like(keys)
    echo[d:, 1] += keys[:-d, 0] * .35
    echo[d:, 0] += keys[:-d, 1] * .35
    echo[2 * d:] += keys[:-2 * d] * .12
    keys += echo

    bus = pads * .28 + keys * .16 + low * .32 + drums * .30
    bus = reverb(bus)
    t = np.arange(n) / SR
    bus *= np.clip(t / 2.0, 0, 1)[:, None]                         # 開頭淡入
    bus *= np.clip((total - t) / 3.5, 0, 1)[:, None]               # 結尾淡出
    return bus


def duck_gain(narr, depth_db=-9.0):
    """旁白出現時把音樂壓低 depth_db；每 10 ms 一格，起音 60 ms、釋放 450 ms。"""
    hop = SR // 100
    m = len(narr) // hop
    rms = np.sqrt(np.mean(narr[: m * hop].reshape(m, hop) ** 2, axis=1))
    on = (rms > 0.01).astype(float)
    on = np.convolve(on, np.ones(25), "full")[:m] > 0                # 句尾多撐 250 ms
    g_lo = 10 ** (depth_db / 20)
    g, cur = np.empty(m), 1.0
    for i in range(m):
        target = g_lo if on[i] else 1.0
        k = 1 - np.exp(-1 / (6 if target < cur else 45))
        cur += k * (target - cur)
        g[i] = cur
    return np.interp(np.arange(len(narr)) / hop, np.arange(m), g)


# ---------------------------------------------------------------- 音效
def sfx_pop(pitch=1.0):
    t = tt(0.16)
    f = pitch * (650 + 500 * (1 - np.exp(-t / 0.02)))
    sig = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.045)
    return sig * np.minimum(t / 0.003, 1)


def sfx_sent():
    t = tt(0.22)
    n = RNG.standard_normal(len(t))
    sw = onepole_lp(n, 800 + 5000 * t / 0.22) * np.sin(np.pi * np.clip(t / 0.14, 0, 1)) * .5
    pop = np.zeros_like(t)
    p = sfx_pop(.78)
    pop[int(.08 * SR):int(.08 * SR) + len(p)] = p[: len(t) - int(.08 * SR)]
    return sw + pop


def sfx_tap():
    t = tt(0.05)
    n = RNG.standard_normal(len(t))
    return (n - onepole_lp(n, 2500)) * np.exp(-t / 0.004) * .6 + np.sin(2 * np.pi * 1900 * t) * np.exp(-t / 0.012)


def sfx_whoosh(dur=0.8):
    t = tt(dur)
    x = t / dur
    n = RNG.standard_normal(len(t))
    cut = 250 + 3800 * np.sin(np.pi * x) ** 2
    sig = onepole_lp(onepole_lp(n, cut), cut)
    return sig * np.sin(np.pi * x) ** 1.5 * 2.2


def sfx_chime(notes, gap=0.09, decay=0.6):
    sig = np.zeros(int((gap * len(notes) + decay * 4) * SR))
    for k, nn in enumerate(notes):
        b = bell(hz(nn), decay)
        i = int(k * gap * SR)
        sig[i:i + len(b)] += b[: len(sig) - i]
    return sig / len(notes) ** .5


def sfx_thud():
    t = tt(0.5)
    f = 55 + 90 * np.exp(-t / 0.05)
    body = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.12)
    n = RNG.standard_normal(len(t))
    return body + onepole_lp(n, 1200) * np.exp(-t / 0.02) * .8


def sfx_tick(pitch=1.0):
    t = tt(0.06)
    return (np.sin(2 * np.pi * 1300 * pitch * t) + .4 * np.sin(2 * np.pi * 2900 * pitch * t)) * np.exp(-t / 0.012)


def sfx_hit():
    """片頭／片尾：低頻 boom＋鐘聲和弦。"""
    t = tt(2.5)
    boom = np.sin(2 * np.pi * np.cumsum(38 + 40 * np.exp(-t / .08)) / SR) * np.exp(-t / .7)
    return boom * .9 + sfx_chime([72, 79, 84, 88], gap=.02, decay=.9)[: len(t)] * .7


def sfx_coin():
    return sfx_chime([88, 93, 96, 100], gap=.05, decay=.35)


SFX = {  # 第二個數字是音量：提示音約比旁白低 6～11 dB，按鈕與泡泡音更輕
    "pop": (sfx_pop, .165), "sent": (sfx_sent, .16), "tap": (sfx_tap, .17), "whoosh": (sfx_whoosh, .125),
    "alert": (lambda: sfx_chime([76, 81], gap=.12, decay=.4), .09),
    "found": (lambda: sfx_chime([72, 76, 79, 84], gap=.06, decay=.5), .09),
    "reject": (lambda: sfx_thud() * .7 + sfx_chime([67, 60], gap=.13, decay=.3)[: int(.5 * SR)] * .6, .14),
    "tick": (sfx_tick, .12), "blip": (lambda: sfx_pop(1.6) * .6, .12),
    "hit": (sfx_hit, .15), "coin": (sfx_coin, .115),
}


def sfx_bus(events, n):
    bus = np.zeros((n, 2))
    for t0, name, kw in events:
        fn, gain = SFX[name]
        kw = dict(kw)
        pan = kw.pop("pan", 0.0)
        g = kw.pop("gain", 1.0)
        place(bus, t0, fn(**kw), pan=pan, gain=gain * g)
    return reverb(bus, sec=1.2, wet=.15)


# ---------------------------------------------------------------- 混音
def mix(narration, events, total, marks):
    n = int((total + 1) * SR)
    narr = np.zeros(n)
    narr[: min(n, len(narration))] = narration[:n]
    bgm = music(total, marks["drums"], marks["end"])[:n]
    bgm *= 0.05 / np.sqrt(np.mean(bgm ** 2) + 1e-12)                # 音樂 RMS 約 -26 dBFS
    bgm *= duck_gain(narr)[:, None]
    out = (bgm + sfx_bus(events, n) + narr[:, None]) * 10 ** (3 / 20)  # 整體約 -16 LUFS
    return limit(out).astype(np.float32)


def limit(x, ceiling=0.89, block=96, release=0.12):
    """前瞻限幅：每 2 ms 一格取需要的增益，往前看一格，慢慢放開，峰值 ≤ ceiling（-1 dBFS）。"""
    m = -(-len(x) // block)
    pk = np.zeros(m * block)
    pk[: len(x)] = np.max(np.abs(x), axis=1)
    need = np.minimum(1.0, ceiling / np.maximum(pk.reshape(m, block).max(1), 1e-9))
    need = np.minimum(need, np.append(need[1:], 1.0))
    g, cur = np.empty(m), 1.0
    k = 1 - np.exp(-block / (release * SR))
    for i in range(m):
        cur = need[i] if need[i] < cur else cur + k * (need[i] - cur)
        g[i] = cur
    g = np.minimum(g, np.append(g[1:], 1.0))
    gs = np.interp(np.arange(len(x)), np.arange(m) * block + block / 2, g)
    gs = np.minimum(gs, np.repeat(need, block)[: len(x)])
    return x * gs[:, None]
