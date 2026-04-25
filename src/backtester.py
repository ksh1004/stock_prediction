"""
backtester.py
─────────────
기술적 신호 또는 모델 예측값을 기반으로 투자 전략을 백테스팅하는 모듈.

주요 클래스:
    Backtester  : 단일 종목 전략 검증
    PortfolioBacktester : 다중 종목 포트폴리오 전략 검증

사용 예시:
    bt = Backtester(prices, signals)
    result = bt.run()
    bt.plot()
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams["font.family"] = "Malgun Gothic"
matplotlib.rcParams["axes.unicode_minus"] = False


class Backtester:
    """
    단일 종목 백테스터.

    Parameters
    ----------
    prices   : Series  (날짜 인덱스, 종가)
    signals  : Series  (날짜 인덱스, 1=매수 보유 / -1=공매도 / 0=현금)
    init_cash : float  초기 투자금 (원)
    """

    def __init__(self, prices: pd.Series, signals: pd.Series, init_cash=10_000_000):
        self.prices = prices.copy()
        self.signals = signals.copy()
        self.init_cash = init_cash
        self.result = None

    def run(self):
        """백테스팅 실행. 포지션 진입/청산 내역과 포트폴리오 가치를 계산."""
        prices = self.prices.reindex(self.signals.index).ffill()

        # 일별 수익률
        daily_return = prices.pct_change().fillna(0)

        # 전략 수익률: 신호가 1이면 다음 날 수익률을 얻음 (신호 발생 다음 날 거래 가정)
        strategy_return = self.signals.shift(1).fillna(0) * daily_return

        # 누적 수익률
        cumulative_market = (1 + daily_return).cumprod()          # 단순 보유 전략
        cumulative_strategy = (1 + strategy_return).cumprod()     # 신호 기반 전략

        # 포트폴리오 가치
        portfolio_value = self.init_cash * cumulative_strategy

        self.result = pd.DataFrame({
            "Price": prices,
            "Signal": self.signals,
            "Daily_Return": daily_return,
            "Strategy_Return": strategy_return,
            "Buy_Hold": self.init_cash * cumulative_market,
            "Portfolio": portfolio_value,
        })
        return self

    def summary(self):
        """성과 지표 계산 및 출력."""
        if self.result is None:
            self.run()

        r = self.result
        strat = r["Strategy_Return"]
        market = r["Daily_Return"]

        # 총 수익률
        total_return_strat = (r["Portfolio"].iloc[-1] / self.init_cash - 1) * 100
        total_return_bh = (r["Buy_Hold"].iloc[-1] / self.init_cash - 1) * 100

        # 연간 수익률 (CAGR)
        n_years = len(r) / 252
        cagr_strat = ((r["Portfolio"].iloc[-1] / self.init_cash) ** (1 / n_years) - 1) * 100
        cagr_bh = ((r["Buy_Hold"].iloc[-1] / self.init_cash) ** (1 / n_years) - 1) * 100

        # 샤프 비율: 수익률 / 변동성 (높을수록 리스크 대비 수익 좋음)
        sharpe = strat.mean() / strat.std() * np.sqrt(252) if strat.std() > 0 else 0

        # 최대 낙폭 (MDD): 고점 대비 최대 하락률
        rolling_max = r["Portfolio"].cummax()
        drawdown = (r["Portfolio"] - rolling_max) / rolling_max
        mdd = drawdown.min() * 100

        # 승률
        trades = strat[self.result["Signal"].shift(1) != 0]
        win_rate = (trades > 0).mean() * 100 if len(trades) > 0 else 0

        metrics = {
            "전략 총수익률": f"{total_return_strat:.1f}%",
            "단순보유 총수익률": f"{total_return_bh:.1f}%",
            "전략 연간수익률(CAGR)": f"{cagr_strat:.1f}%",
            "단순보유 CAGR": f"{cagr_bh:.1f}%",
            "샤프 비율": f"{sharpe:.2f}",
            "최대 낙폭(MDD)": f"{mdd:.1f}%",
            "거래 승률": f"{win_rate:.1f}%",
        }

        print("=" * 40)
        print("        백테스팅 성과 요약")
        print("=" * 40)
        for k, v in metrics.items():
            print(f"  {k:<22}: {v}")
        print("=" * 40)
        return metrics

    def plot(self, title="백테스팅 결과"):
        """포트폴리오 가치 및 매수/매도 신호 시각화."""
        if self.result is None:
            self.run()
        r = self.result

        fig, axes = plt.subplots(3, 1, figsize=(14, 10), sharex=True)

        # ── 상단: 가격 + 신호 ──
        axes[0].plot(r.index, r["Price"], color="black", linewidth=1, label="주가")
        buy_pts = r[r["Signal"] == 1]
        sell_pts = r[r["Signal"] == -1]
        axes[0].scatter(buy_pts.index, buy_pts["Price"], marker="^", color="red",
                        s=60, zorder=5, label="매수 신호")
        axes[0].scatter(sell_pts.index, sell_pts["Price"], marker="v", color="blue",
                        s=60, zorder=5, label="매도 신호")
        axes[0].set_title(f"{title} — 주가 및 매매 신호")
        axes[0].legend(fontsize=9)
        axes[0].grid(True, alpha=0.3)

        # ── 중단: 포트폴리오 가치 비교 ──
        axes[1].plot(r.index, r["Portfolio"] / 1e6, color="tomato", linewidth=1.5, label="전략")
        axes[1].plot(r.index, r["Buy_Hold"] / 1e6, color="steelblue", linewidth=1.5,
                     linestyle="--", label="단순 보유")
        axes[1].set_title("포트폴리오 가치 비교 (백만원)")
        axes[1].legend()
        axes[1].grid(True, alpha=0.3)

        # ── 하단: 낙폭 ──
        rolling_max = r["Portfolio"].cummax()
        drawdown = (r["Portfolio"] - rolling_max) / rolling_max * 100
        axes[2].fill_between(r.index, drawdown, 0, color="salmon", alpha=0.6)
        axes[2].set_title("전략 낙폭 (%)")
        axes[2].grid(True, alpha=0.3)

        plt.tight_layout()
        plt.show()


def compute_portfolio_metrics(returns: pd.DataFrame):
    """
    다중 종목 수익률 DataFrame을 받아 종목별 성과 지표를 계산.

    Parameters
    ----------
    returns : DataFrame  (날짜 인덱스, 종목코드 컬럼, 일별 수익률 값)

    Returns
    -------
    DataFrame  종목별 성과 지표
    """
    n_years = len(returns) / 252

    metrics = pd.DataFrame({
        "총수익률(%)": (((1 + returns).prod() - 1) * 100),
        "CAGR(%)": (((1 + returns).prod() ** (1 / n_years) - 1) * 100),
        "변동성(연환산%)": (returns.std() * np.sqrt(252) * 100),
        "샤프비율": (returns.mean() / returns.std() * np.sqrt(252)),
        "최대낙폭(%)": (
            returns.apply(lambda x: ((1 + x).cumprod() / (1 + x).cumprod().cummax() - 1).min() * 100)
        ),
    })
    return metrics.round(2)
