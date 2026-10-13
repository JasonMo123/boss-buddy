"""從官方競賽數據推算每台車的油耗、怠速、營運天數與故障碼指標，供試算表「數據推算」頁使用。

用法：
    python3 -I 分析/derive_metrics.py

輸入：原始官方資料/解壓原檔/.../output data_Hotai_20260511.xlsx（output 工作表）
輸出：
    分析/數據推算_每車.csv   每台車一列，試算表直接引用
    分析/數據推算_每日.csv   每台車每天一列，供稽核
快取：分析/_cache/output.parquet（第一次讀 Excel 約需數分鐘，之後讀快取）
"""

from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SRC = (ROOT / "原始官方資料" / "解壓原檔" / "HINO 競賽數據"
       / "HINO x GenAI：運用商用車車聯網大數據，發展商用車頭家的智慧夥伴"
       / "output data_Hotai_20260511.xlsx")
OUT = Path(__file__).resolve().parent
CACHE = OUT / "_cache" / "output.parquet"

MAX_GAP_S = 120        # 相鄰兩筆超過 120 秒視為斷線，不計入時長
MAX_STEP = 5.0         # 相鄰兩筆里程或油耗跳動超過 5（km／L）視為異常（例：總里程歸零）
MIN_DAY_KM = 30        # 一天行駛未滿 30 km 不列入有效日
CORE_DATE_MIN_CARS = 5 # 當天至少 5 台車有資料，才算主要抽樣日
DPF_CODES = {"P2002", "P244B"}


def load() -> pd.DataFrame:
    if CACHE.exists():
        return pd.read_parquet(CACHE)
    df = pd.read_excel(SRC, sheet_name="output", dtype=str, engine="openpyxl")
    CACHE.parent.mkdir(exist_ok=True)
    df.to_parquet(CACHE)
    return df


def main() -> None:
    raw = load().replace("", np.nan)
    df = raw[["Type", "enabledCode", "carNum", "time", "carStatus",
              "can.totalMileage", "can.engine.totalFuelUsed"]].copy()
    df["t"] = pd.to_datetime(df["time"])
    df["d"] = df["t"].dt.date
    df["odo"] = pd.to_numeric(df["can.totalMileage"])
    df["fuel"] = pd.to_numeric(df["can.engine.totalFuelUsed"])
    df = df.sort_values(["enabledCode", "t"])

    # 每筆資料到下一筆之間的時長、里程、用油，歸屬於這一筆的車輛狀態
    g = df.groupby("enabledCode")
    df["dt"] = -g["t"].diff(-1).dt.total_seconds()
    df["dkm"] = -g["odo"].diff(-1)
    df["dL"] = -g["fuel"].diff(-1)
    df.loc[(df.dt <= 0) | (df.dt > MAX_GAP_S), "dt"] = np.nan
    df.loc[(df.dkm < 0) | (df.dkm > MAX_STEP), "dkm"] = np.nan
    df.loc[(df.dL < 0) | (df.dL > MAX_STEP), "dL"] = np.nan
    idle = df.carStatus == "2"
    df["idle_L"] = np.where(idle, df.dL, 0.0)
    df["idle_s"] = np.where(idle, df.dt, 0.0)

    day = df.groupby(["enabledCode", "d"]).agg(
        Type=("Type", "first"), km=("dkm", "sum"), L=("dL", "sum"),
        idle_L=("idle_L", "sum"), idle_h=("idle_s", lambda s: s.sum() / 3600),
    ).reset_index()
    day["出勤日"] = day.L > 0
    day["有效日"] = (day.km >= MIN_DAY_KM) & (day.L > 0)
    day.to_csv(OUT / "數據推算_每日.csv", index=False, encoding="utf-8-sig")

    # 日均量用所有出勤日（與出勤率口徑一致）；油耗比率只用有效日，避免短程日雜訊
    v = day[day["出勤日"]].groupby("enabledCode").agg(
        Type=("Type", "first"), 出勤日數=("d", "size"), 日均里程_km=("km", "mean"),
        日均用油_L=("L", "mean"), 日均怠速用油_L=("idle_L", "mean"),
        日均怠速時數_h=("idle_h", "mean"),
    )
    valid = day[day["有效日"]].groupby("enabledCode").agg(
        有效日數=("d", "size"), 總里程_km=("km", "sum"), 總用油_L=("L", "sum"),
        總怠速用油_L=("idle_L", "sum"), 總怠速時數_h=("idle_h", "sum"),
    )
    v = v.join(valid, how="inner")
    v["日均行駛用油_L"] = v.日均用油_L - v.日均怠速用油_L
    v["行駛油耗_L每百公里"] = (v.總用油_L - v.總怠速用油_L) / v.總里程_km * 100
    v["怠速每小時用油_L"] = v.總怠速用油_L / v.總怠速時數_h.replace(0, np.nan)
    v["怠速用油佔比"] = v.總怠速用油_L / v.總用油_L

    # 抽樣日出現比例：主要抽樣日中，這台車有資料的比例，近似出勤率
    cars_per_day = df.groupby("d").enabledCode.nunique()
    core = cars_per_day[cars_per_day >= CORE_DATE_MIN_CARS].index
    present = df[df.d.isin(core)].groupby("enabledCode").d.nunique()
    v["主要抽樣日數"] = len(core)
    v["抽樣日出現比例"] = present / len(core)

    # 行駛油耗與同車型前 25% 的差距：超出部分佔行駛用油的比例
    p25 = v.groupby("Type")["行駛油耗_L每百公里"].transform(lambda s: s.quantile(0.25))
    v["同車型前25%行駛油耗"] = p25
    v["超出前25%比例"] = (1 - p25 / v["行駛油耗_L每百公里"]).clip(lower=0)
    v["日均超出前25%用油_L"] = v.日均行駛用油_L * v["超出前25%比例"]

    # 故障碼：U 開頭多為通訊類，只把 P 開頭視為機械保修警示
    codes = pd.concat([
        raw.loc[raw[f"event[{i}].type"] == "9", ["enabledCode", f"event[{i}].info.dtcCodes[0]"]]
        .rename(columns={f"event[{i}].info.dtcCodes[0]": "code"}) for i in range(3)
    ]).dropna()
    by_car = codes.groupby("enabledCode").code
    v["故障碼次數"] = by_car.size().reindex(v.index).fillna(0).astype(int)
    v["故障碼種類"] = by_car.agg(lambda s: "、".join(sorted(s.unique()))).reindex(v.index).fillna("")
    v["有P碼"] = by_car.agg(lambda s: int(s.str.startswith("P").any())).reindex(v.index).fillna(0).astype(int)
    v["有DPF碼"] = by_car.agg(lambda s: int(s.isin(DPF_CODES).any())).reindex(v.index).fillna(0).astype(int)
    v["車牌"] = raw.groupby("enabledCode").carNum.agg(lambda s: s.dropna().iloc[0])

    v = v.reset_index().rename(columns={"enabledCode": "納管碼", "Type": "車型"})
    v["車型"] = v["車型"].astype(int)
    cols = ["納管碼", "車牌", "車型", "出勤日數", "有效日數", "主要抽樣日數", "抽樣日出現比例",
            "日均里程_km", "日均用油_L", "日均怠速用油_L", "日均行駛用油_L", "日均怠速時數_h",
            "怠速每小時用油_L", "怠速用油佔比", "行駛油耗_L每百公里", "同車型前25%行駛油耗",
            "超出前25%比例", "日均超出前25%用油_L", "故障碼次數", "故障碼種類", "有P碼", "有DPF碼"]
    v = v[cols].sort_values(["車型", "納管碼"])
    v.to_csv(OUT / "數據推算_每車.csv", index=False, encoding="utf-8-sig")
    print(v.round(3).to_string(index=False))


if __name__ == "__main__":
    main()
