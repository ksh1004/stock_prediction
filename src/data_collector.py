"""
data_collector.py
─────────────────
KOSPI / KOSDAQ 전 종목 데이터를 수집하고 로컬에 저장하는 모듈.

주요 함수:
    get_stock_list()     → 전 종목 리스트 DataFrame 반환
    download_all()       → 전 종목 가격 데이터 다운로드 후 parquet 저장
    load_stock()         → 저장된 parquet에서 단일 종목 로드
    load_all_prices()    → 저장된 전 종목 종가 데이터를 하나의 DataFrame으로 합치기
"""

import os
import time
import warnings
import pandas as pd
import FinanceDataReader as fdr
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

warnings.filterwarnings("ignore")

# ── 경로 설정 ───────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRICE_DIR = os.path.join(BASE_DIR, "data", "prices")
LIST_DIR = os.path.join(BASE_DIR, "data", "stock_list")
os.makedirs(PRICE_DIR, exist_ok=True)
os.makedirs(LIST_DIR, exist_ok=True)


def get_stock_list(markets=("KOSPI", "KOSDAQ"), refresh=False):
    """
    KOSPI / KOSDAQ 전 종목 리스트를 가져온다.

    Parameters
    ----------
    markets  : 가져올 시장 목록
    refresh  : True면 캐시 무시하고 새로 다운로드

    Returns
    -------
    DataFrame  columns: Code, Name, Market, Sector, Industry, ListingDate
    """
    cache_path = os.path.join(LIST_DIR, "stock_list.parquet")

    # 캐시가 있으면 재사용
    if not refresh and os.path.exists(cache_path):
        df = pd.read_parquet(cache_path)
        print(f"캐시에서 로드: {len(df)}개 종목")
        return df

    frames = []
    for market in markets:
        try:
            df = fdr.StockListing(market)
            df["Market"] = market
            frames.append(df)
            print(f"  {market}: {len(df)}개 종목")
        except Exception as e:
            print(f"  {market} 실패: {e}")

    if not frames:
        raise RuntimeError("종목 리스트 수집 실패")

    result = pd.concat(frames, ignore_index=True)

    # 종목코드 6자리 맞추기
    if "Code" in result.columns:
        result["Code"] = result["Code"].astype(str).str.zfill(6)

    # Dept(기업규모: 대기업/중견기업 등) → 컬럼명을 직관적으로 변경
    if "Dept" in result.columns:
        result = result.rename(columns={"Dept": "CompanySize"})

    # 시가총액 컬럼 이름 통일
    if "Marcap" in result.columns:
        result = result.rename(columns={"Marcap": "MarketCap"})

    result.to_parquet(cache_path, index=False)
    print(f"\n총 {len(result)}개 종목 저장 완료")
    return result


def _download_one(code, name, start, end, retry=2):
    """
    단일 종목 데이터를 다운로드하고 parquet으로 저장.
    성공하면 True, 실패하면 False 반환.
    """
    save_path = os.path.join(PRICE_DIR, f"{code}.parquet")

    # 이미 있으면 스킵
    if os.path.exists(save_path):
        return True

    for attempt in range(retry + 1):
        try:
            df = fdr.DataReader(code, start, end)
            if df is None or len(df) < 20:  # 데이터가 너무 적으면 스킵
                return False
            df.index.name = "Date"
            df.to_parquet(save_path)
            return True
        except Exception:
            if attempt < retry:
                time.sleep(0.5)
    return False


def download_all(start="2019-01-01", end="2024-12-31",
                 markets=("KOSPI", "KOSDAQ"), max_workers=8, refresh=False):
    """
    전 종목 가격 데이터를 병렬로 다운로드한다.

    Parameters
    ----------
    start       : 시작일 (YYYY-MM-DD)
    end         : 종료일 (YYYY-MM-DD)
    max_workers : 동시 다운로드 스레드 수 (너무 높으면 서버 차단 위험)
    refresh     : True면 기존 파일 무시하고 재다운로드

    Returns
    -------
    dict  { "success": int, "fail": int, "skipped": int }
    """
    stock_list = get_stock_list(markets=markets)

    if refresh:
        # 기존 파일 삭제
        for f in os.listdir(PRICE_DIR):
            os.remove(os.path.join(PRICE_DIR, f))

    codes = stock_list["Code"].tolist()
    names = stock_list["Name"].tolist() if "Name" in stock_list.columns else [""] * len(codes)

    success, fail = 0, 0
    already = sum(1 for c in codes if os.path.exists(os.path.join(PRICE_DIR, f"{c}.parquet")))

    print(f"전체: {len(codes)}개  이미 저장됨: {already}개  다운로드 필요: {len(codes)-already}개")

    tasks = [(c, n) for c, n in zip(codes, names)
             if not os.path.exists(os.path.join(PRICE_DIR, f"{c}.parquet"))]

    if not tasks:
        print("모든 종목이 이미 저장돼 있습니다.")
        return {"success": already, "fail": 0, "skipped": already}

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_download_one, c, n, start, end): (c, n)
                   for c, n in tasks}

        with tqdm(total=len(futures), desc="다운로드 중", unit="종목") as pbar:
            for future in as_completed(futures):
                ok = future.result()
                if ok:
                    success += 1
                else:
                    fail += 1
                pbar.update(1)
                pbar.set_postfix(성공=success, 실패=fail)

    print(f"\n완료  성공: {success + already}  실패: {fail}")
    return {"success": success + already, "fail": fail, "skipped": already}


def load_stock(code):
    """
    저장된 parquet에서 단일 종목 데이터를 로드한다.

    Returns
    -------
    DataFrame or None
    """
    code = str(code).zfill(6)
    path = os.path.join(PRICE_DIR, f"{code}.parquet")
    if not os.path.exists(path):
        return None
    return pd.read_parquet(path)


def load_all_close_prices(min_days=200):
    """
    저장된 전 종목 종가(Close)를 하나의 DataFrame으로 합친다.

    Parameters
    ----------
    min_days : 이 값보다 거래일 수가 적은 종목은 제외

    Returns
    -------
    DataFrame  index=Date, columns=종목코드
    """
    files = [f for f in os.listdir(PRICE_DIR) if f.endswith(".parquet")]
    print(f"{len(files)}개 파일 로드 중...")

    frames = {}
    for f in tqdm(files, desc="종가 합치는 중"):
        code = f.replace(".parquet", "")
        try:
            df = pd.read_parquet(os.path.join(PRICE_DIR, f))
            if "Close" in df.columns and len(df) >= min_days:
                frames[code] = df["Close"]
        except Exception:
            pass

    result = pd.DataFrame(frames)
    result.index = pd.to_datetime(result.index)
    result = result.sort_index()
    print(f"완료: {result.shape[1]}개 종목, {result.shape[0]}거래일")
    return result
