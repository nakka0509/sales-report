import pandas as pd
import requests
import io
import os
import re
import ssl
from datetime import datetime
from imap_tools import MailBox, AND

# .envファイルがあればローカル実行時に読み込む（python-dotenv不要版）
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    with open(_env_path, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# ==========================================
# 【設定エリア】
# ==========================================
FETCH_MODE = "LATEST_20_DAYS"  # 直近20日のみ

# GitHub Actions（Linux）でもWindows（ローカル）でも動くパス設定
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(_SCRIPT_DIR, "データベース")
EXCEL_FILE = os.path.join(OUTPUT_DIR, "売上管理表.xlsx")

# パスワードは環境変数から取得（未設定の場合はエラー）
def _get_env(key):
    val = os.environ.get(key)
    if not val:
        print(f"[エラー] 環境変数 {key} が設定されていません。")
        print("  ローカル実行: .envファイルを作成してください")
        print("  GitHub Actions: Secretsに登録してください")
        raise SystemExit(1)
    return val

ACCOUNTS = [
    {
        "name": "きむら", "type": "EXCEL",
        "server": "imap.softbank.jp",
        "user": _get_env("KIMURA_USER"),
        "password": _get_env("KIMURA_PASSWORD")
    },
    {
        "name": "JA", "type": "TEXT",
        "server": "imap.gmail.com",
        "user": _get_env("JA_USER"),
        "password": _get_env("JA_PASSWORD")
    }
]


LATITUDE = 34.34
LONGITUDE = 134.04
# ==========================================

def get_ssl_context():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try: ctx.set_ciphers('DEFAULT@SECLEVEL=1')
    except: pass
    return ctx

WEATHER_MAP = {0: "晴れ", 1: "晴れ", 2: "曇り", 3: "曇り",
               45: "霧", 48: "霧", 51: "小雨", 53: "小雨", 55: "小雨",
               61: "雨", 63: "雨", 65: "大雨", 71: "雪", 73: "雪", 75: "大雪",
               80: "雨", 81: "雨", 82: "大雨", 95: "雷雨"}

def get_weather_batch(dates):
    """日付リストを受け取り、1回のAPIコールで全日付の天気を取得する（高速化）"""
    if not dates:
        return {}
    sorted_dates = sorted(dates)
    start = sorted_dates[0].strftime("%Y-%m-%d")
    end = sorted_dates[-1].strftime("%Y-%m-%d")
    url = (f"https://archive-api.open-meteo.com/v1/archive"
           f"?latitude={LATITUDE}&longitude={LONGITUDE}"
           f"&start_date={start}&end_date={end}"
           f"&daily=weather_code,temperature_2m_max,temperature_2m_min"
           f"&timezone=Asia%2FTokyo")
    try:
        res = requests.get(url, timeout=15).json()
        days = res['daily']['time']
        codes = res['daily']['weather_code']
        maxts = res['daily']['temperature_2m_max']
        mints = res['daily']['temperature_2m_min']
        result = {}
        for i, day_str in enumerate(days):
            result[day_str] = (
                WEATHER_MAP.get(codes[i], "曇り"),
                maxts[i],
                mints[i]
            )
        return result
    except Exception as e:
        print(f"  天気API取得エラー: {e}")
        return {}

def parse_ja_text(body):
    data = []
    date_match = re.search(r'(\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}:\d{2})', body)
    if date_match:
        # 年はmain()側でメール受信日から補正される
        current_year = datetime.now().year
        date_str = f"{current_year}年{date_match.group(1)}"
        try: dt = datetime.strptime(date_str, "%Y年%m月%d日 %H:%M:%S")
        except: dt = datetime.now()
    else: return []

    pattern = re.compile(r'(\S+)\s+(\d+)円\s+(\d+)\s*￥\s*([\d,]+)')
    for match in pattern.finditer(body):
        data.append({
            "日付": dt,
            "店舗名": "JA産直空の街",
            "品名": match.group(1),
            "単価": int(match.group(2)),
            "数量": int(match.group(3)),
            "小計": int(match.group(4).replace(',', ''))
        })
    return data

def normalize_kimura_df(df, fallback_date=None):
    """きむらExcel解析：店舗名とコードを厳密に区別する"""
    cols_str = ' '.join([str(c) for c in df.columns])
    if not any(x in cols_str for x in ['日付', '品名', '金額']):
        # 列名にそれらしきものがなければ、最初の数行からヘッダを探す
        for i in range(min(5, len(df))):
            row_vals = [str(v) for v in df.iloc[i].values]
            if any('日付' in v or '品名' in v or '金額' in v for v in row_vals):
                df.columns = df.iloc[i]
                df = df.iloc[i+1:].reset_index(drop=True)
                break
                
    df.columns = [str(c).replace(' ', '').replace('　', '').replace('\n', '').strip() for c in df.columns]
    
    col_map = {}
    for col in df.columns:
        if col in ['店舗名', '支店名', '店名']:
            col_map[col] = '店舗名'
        elif '店' in col and not any(x in col for x in ['コード', 'CD', 'ID', '番号', 'NO']):
            if '店舗名' not in col_map.values():
                col_map[col] = '店舗名'
        elif any(x in col for x in ['商品名', '品名', '商品', '品目']): col_map[col] = '品名'
        elif any(x in col for x in ['金額', '小計', '売上金額']): col_map[col] = '小計'
        elif any(x in col for x in ['個数', '数量', '点数', '売上数量', '売上点数']): col_map[col] = '数量'
        elif '日付' in col: col_map[col] = '日付'
    
    df.rename(columns=col_map, inplace=True)
    
    if '小計' not in df.columns: return None
    
    if '品名' in df.columns:
        df.dropna(subset=['品名'], inplace=True)
        df = df[df['品名'].astype(str).str.strip() != '']
        df = df[~df['品名'].astype(str).str.contains('合計', na=False, regex=False)]
    else: return None

    df['数量'] = pd.to_numeric(df['数量'], errors='coerce').fillna(0)
    df['小計'] = pd.to_numeric(df['小計'], errors='coerce').fillna(0)
    df['単価'] = 0
    mask = df['数量'] > 0
    df.loc[mask, '単価'] = (df.loc[mask, '小計'] / df.loc[mask, '数量']).astype(int)

    if '日付' in df.columns:
        df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
        if fallback_date is not None:
            df['日付'] = df['日付'].fillna(pd.to_datetime(fallback_date).normalize().tz_localize(None))
    else:
        if fallback_date is not None:
            df['日付'] = pd.to_datetime(fallback_date).normalize().tz_localize(None)
    
    if '店舗名' not in df.columns or df['店舗名'].isna().all():
        if len(df.columns) >= 3:
             df['店舗名'] = df.iloc[:, 2]

    df['店舗名'] = df['店舗名'].fillna("きむら(店舗不明)")
    
    return df

def main():
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    print(f"[{FETCH_MODE}] データを収集中...")
    new_dfs = {"きむら": [], "JA": []}
    from datetime import timedelta
    fetch_start_date = datetime.now().date() - timedelta(days=20)
    limit = None  # 直近20日のみ処理
    ssl_ctx = get_ssl_context()

    for acc in ACCOUNTS:
        print(f"[{acc['name']}] ({acc['user']}) に接続中...")
        try:
            with MailBox(acc['server'], ssl_context=ssl_ctx).login(acc['user'], acc['password']) as mailbox:
                if acc['name'] == 'JA':
                    # ① JA直接メール（sanchoku@jakagawaken.jp から直接受信）
                    criteria_direct = AND(from_='sanchoku@jakagawaken.jp', date_gte=fetch_start_date)
                    # ② Outlook転送メール（sawaka878@outlook.jp 経由で "FW: JA" として届くもの）
                    criteria_forward = AND(from_='sawaka878@outlook.jp', date_gte=fetch_start_date)

                    ja_count = 0
                    for criteria in [criteria_direct, criteria_forward]:
                        for msg in mailbox.fetch(criteria, reverse=True):
                            # Outlook転送の場合は件名に "JA" or "売上" が含まれるものだけ処理
                            if criteria == criteria_forward:
                                subj = msg.subject or ""
                                if not any(kw in subj for kw in ["JA", "売上情報", "直売所"]):
                                    continue
                            content = msg.text or msg.html
                            if content:
                                rows = parse_ja_text(content)
                                if rows:
                                    if msg.date:
                                        mail_year = msg.date.year
                                        for row in rows:
                                            if hasattr(row['日付'], 'replace'):
                                                row['日付'] = row['日付'].replace(year=mail_year)
                                    df = pd.DataFrame(rows)
                                    df['取得元'] = "JA"
                                    new_dfs["JA"].append(df)
                                    ja_count += 1
                    print(f"  → JA: 直接+転送メール合計 {ja_count}通を処理")

                else:
                    # きむら（添付Excelを処理）
                    criteria = AND(date_gte=fetch_start_date)
                    for msg in mailbox.fetch(criteria, reverse=True):
                        if msg.attachments:
                            for att in msg.attachments:
                                if att.filename.endswith(('.xlsx', '.xls')):
                                    try:
                                        df = pd.read_excel(io.BytesIO(att.payload), header=None)
                                        filename = att.filename
                                        date_match = re.search(r'(\d{1,2})月(\d{1,2})日', filename)
                                        mail_date = msg.date
                                        if date_match:
                                            month = int(date_match.group(1))
                                            day = int(date_match.group(2))
                                            year = mail_date.year
                                            if month > mail_date.month + 6:
                                                year -= 1
                                            try:
                                                mail_date = datetime(year, month, day)
                                            except ValueError:
                                                pass
                                        df = normalize_kimura_df(df, fallback_date=mail_date)
                                        if df is not None:
                                            df['取得元'] = "きむら"
                                            new_dfs["きむら"].append(df)
                                    except Exception: pass
        except Exception as e: print(f"[エラー] ({acc['name']}): {e}")

    old_dfs = {"きむら": None, "JA": None}
    if os.path.exists(EXCEL_FILE):
        try:
            full_df = pd.read_excel(EXCEL_FILE, sheet_name="Data")
            old_dfs["きむら"] = full_df[full_df['取得元'] == 'きむら'].copy()
            old_dfs["JA"] = full_df[full_df['取得元'] == 'JA'].copy()
        except: pass

    # 取得件数ログ
    for name in ["きむら", "JA"]:
        cnt = sum(len(df) for df in new_dfs[name]) if new_dfs[name] else 0
        old_cnt = len(old_dfs[name]) if old_dfs[name] is not None else 0
        print(f"[{name}] 新規取得: {cnt}件 / DB既存: {old_cnt}件")

    final_dfs = []
    for name in ["きむら", "JA"]:
        current_new = pd.concat(new_dfs[name], ignore_index=True) if new_dfs[name] else pd.DataFrame()
        current_old = old_dfs[name] if old_dfs[name] is not None else pd.DataFrame()

        # 新しいデータがある場合は、そのデータの最小日付以降の古いデータを削除して完全上書きする
        if not current_new.empty:
            current_new['日付'] = pd.to_datetime(current_new['日付'])
            min_new_date = current_new['日付'].min()
            if not current_old.empty:
                current_old = current_old.copy()
                current_old['日付'] = pd.to_datetime(current_old['日付'])
                current_old = current_old[current_old['日付'] < min_new_date]
        
        if current_new.empty and current_old.empty: continue
        combined = pd.concat([current_old, current_new], ignore_index=True)
        combined['日付'] = pd.to_datetime(combined['日付'])
        if name == "JA":
            combined['TempDate'] = combined['日付'].dt.date
            max_dates = combined.groupby('TempDate')['日付'].max().reset_index()
            combined = pd.merge(combined, max_dates, on=['TempDate', '日付'], how='inner')
            del combined['TempDate']
            combined.drop_duplicates(subset=['日付', '店舗名', '品名', '小計'], keep='first', inplace=True)
        else:
            combined.drop_duplicates(subset=['日付', '店舗名', '品名', '小計'], keep='first', inplace=True)
        final_dfs.append(combined)

    if not final_dfs: return

    final_df = pd.concat(final_dfs, ignore_index=True)
    final_df.sort_values('日付', inplace=True)

    print("[天気] 天気を一括取得中...")
    for col in ['天気', '最高気温', '最低気温']:
        if col not in final_df.columns: final_df[col] = None
    missing_mask = (final_df['天気'].isna()) | (final_df['天気'] == "不明")
    if missing_mask.any():
        missing_dates = final_df.loc[missing_mask, '日付'].dt.date.unique()
        print(f"  {len(missing_dates)}日分を1回のAPIコールで取得中...")
        weather_dict = get_weather_batch(list(missing_dates))
        for d in missing_dates:
            day_str = d.strftime("%Y-%m-%d")
            w, t_max, t_min = weather_dict.get(day_str, ("不明", None, None))
            day_mask = final_df['日付'].dt.date == d
            final_df.loc[day_mask, '天気'] = w
            final_df.loc[day_mask, '最高気温'] = t_max
            final_df.loc[day_mask, '最低気温'] = t_min
        print(f"  天気取得完了（{len(weather_dict)}日分）")

    target_cols = ['日付', '取得元', '店舗名', '品名', '単価', '数量', '小計', '天気', '最高気温', '最低気温']
    output_cols = [c for c in target_cols if c in final_df.columns]
    final_df = final_df[output_cols]

    try:
        if os.path.exists(EXCEL_FILE):
            with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                final_df.to_excel(writer, sheet_name="Data", index=False)
        else:
            with pd.ExcelWriter(EXCEL_FILE, mode='w', engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name="Data", index=False)
        print(f"完了: {EXCEL_FILE}")
    except PermissionError: print("エラー: Excelを閉じてください！")

if __name__ == "__main__":
    main()