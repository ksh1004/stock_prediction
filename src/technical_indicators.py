"""
technical_indicators.py
────────────────────────
단일 종목 DataFrame을 받아서 기술적 지표를 계산해 반환하는 모듈.

지표 목록:
    MA(이동평균)   : 20일, 60일, 120일
    RSI            : 14일 상대강도지수 (0~100, 70↑ 과매수 / 30↓ 과매도)
    MACD           : 12일 EMA - 26일 EMA / 시그널(9일 EMA)
    Bollinger Band : 20일 MA ± 2σ
    ATR            : 14일 평균 실제 범위 (변동성 지표)
    Volume MA      : 20일 거래량 이동평균
"""

import pandas as pd
import numpy as np


def add_moving_averages(df, windows=(20, 60, 120)):
    """이동평균선 추가"""
    for w in windows:
        df[f"MA_{w}"] = df["Close"].rolling(window=w).mean()
    return df


def add_rsi(df, period=14):
    """
    RSI (상대강도지수) 계산.
    70 이상 → 과매수(매도 신호), 30 이하 → 과매도(매수 신호)
    """
    delta = df["Close"].diff()                    # 전일 대비 가격 변화
    gain = delta.clip(lower=0)                    # 상승분만
    loss = -delta.clip(upper=0)                   # 하락분만

    # 지수이동평균으로 평균 상승/하락 계산
    avg_gain = gain.ewm(com=period - 1, min_periods=period).mean()
    avg_loss = loss.ewm(com=period - 1, min_periods=period).mean()

    rs = avg_gain / avg_loss.replace(0, np.nan)   # 상대강도 (0으로 나누기 방지)
    df["RSI"] = 100 - (100 / (1 + rs))
    return df


def add_macd(df, fast=12, slow=26, signal=9):
    """
    MACD (이동평균 수렴·확산 지표) 계산.
    MACD선이 시그널선을 상향 돌파 → 매수 신호
    MACD선이 시그널선을 하향 돌파 → 매도 신호
    """
    ema_fast = df["Close"].ewm(span=fast, adjust=False).mean()   # 단기 EMA
    ema_slow = df["Close"].ewm(span=slow, adjust=False).mean()   # 장기 EMA
    df["MACD"] = ema_fast - ema_slow                             # MACD선
    df["MACD_Signal"] = df["MACD"].ewm(span=signal, adjust=False).mean()  # 시그널선
    df["MACD_Hist"] = df["MACD"] - df["MACD_Signal"]            # 히스토그램
    return df


def add_bollinger_bands(df, window=20, num_std=2):
    """
    볼린저 밴드 계산.
    상단 밴드(저항선), 중간선(이동평균), 하단 밴드(지지선)
    가격이 상단 돌파 → 과매수, 하단 이탈 → 과매도
    """
    ma = df["Close"].rolling(window=window).mean()
    std = df["Close"].rolling(window=window).std()
    df["BB_Upper"] = ma + num_std * std   # 상단 밴드
    df["BB_Middle"] = ma                  # 중간선
    df["BB_Lower"] = ma - num_std * std   # 하단 밴드
    df["BB_Width"] = (df["BB_Upper"] - df["BB_Lower"]) / df["BB_Middle"]  # 밴드폭
    df["BB_Pct"] = (df["Close"] - df["BB_Lower"]) / (df["BB_Upper"] - df["BB_Lower"])  # 밴드 내 위치(0~1)
    return df


def add_atr(df, period=14):
    """
    ATR (평균 실제 범위): 변동성 크기 측정.
    값이 클수록 변동성 높음
    """
    high_low = df["High"] - df["Low"]
    high_close = (df["High"] - df["Close"].shift()).abs()
    low_close = (df["Low"] - df["Close"].shift()).abs()
    true_range = pd.concat([high_low, high_close, low_close], axis=1).max(axis=1)
    df["ATR"] = true_range.rolling(period).mean()
    return df


def add_volume_ma(df, window=20):
    """거래량 이동평균 (거래량 급증 여부 판단용)"""
    df["Volume_MA"] = df["Volume"].rolling(window=window).mean()
    df["Volume_Ratio"] = df["Volume"] / df["Volume_MA"]  # 평균 대비 거래량 배율
    return df


def add_all_indicators(df):
    """모든 기술적 지표를 한 번에 추가"""
    df = df.copy()
    df = add_moving_averages(df)
    df = add_rsi(df)
    df = add_macd(df)
    df = add_bollinger_bands(df)
    df = add_atr(df)
    df = add_volume_ma(df)
    return df


def generate_signals(df):
    """
    기술적 지표를 기반으로 매수/매도 신호 생성.

    신호 값:
        1  = 매수
       -1  = 매도
        0  = 중립
    """
    signals = pd.DataFrame(index=df.index)

    # ── RSI 신호 ──
    # 30 이하에서 반등 → 매수 / 70 이상에서 하락 → 매도
    signals["RSI_Signal"] = 0
    signals.loc[df["RSI"] < 30, "RSI_Signal"] = 1
    signals.loc[df["RSI"] > 70, "RSI_Signal"] = -1

    # ── MACD 신호 ──
    # MACD선이 시그널선 상향 돌파 → 매수 / 하향 돌파 → 매도
    signals["MACD_Signal"] = 0
    macd_cross_up = (df["MACD"] > df["MACD_Signal"]) & (df["MACD"].shift(1) <= df["MACD_Signal"].shift(1))
    macd_cross_dn = (df["MACD"] < df["MACD_Signal"]) & (df["MACD"].shift(1) >= df["MACD_Signal"].shift(1))
    signals.loc[macd_cross_up, "MACD_Signal"] = 1
    signals.loc[macd_cross_dn, "MACD_Signal"] = -1

    # ── 볼린저 밴드 신호 ──
    signals["BB_Signal"] = 0
    signals.loc[df["BB_Pct"] < 0.05, "BB_Signal"] = 1   # 하단 밴드 근처 → 매수
    signals.loc[df["BB_Pct"] > 0.95, "BB_Signal"] = -1  # 상단 밴드 근처 → 매도

    # ── 골든크로스 / 데드크로스 ──
    signals["MA_Signal"] = 0
    golden = (df["MA_20"] > df["MA_60"]) & (df["MA_20"].shift(1) <= df["MA_60"].shift(1))
    dead = (df["MA_20"] < df["MA_60"]) & (df["MA_20"].shift(1) >= df["MA_60"].shift(1))
    signals.loc[golden, "MA_Signal"] = 1
    signals.loc[dead, "MA_Signal"] = -1

    # ── 종합 신호 (다수결) ──
    signals["Total_Score"] = signals.sum(axis=1)
    signals["Final_Signal"] = 0
    signals.loc[signals["Total_Score"] >= 2, "Final_Signal"] = 1   # 2개 이상 매수 → 매수
    signals.loc[signals["Total_Score"] <= -2, "Final_Signal"] = -1  # 2개 이상 매도 → 매도

    return signals
