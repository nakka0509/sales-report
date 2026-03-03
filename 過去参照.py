import pandas as pd
import requests
import io
import os
import re
import ssl
import win32com.client
from datetime import datetime, timedelta
from imap_tools import MailBox, AND

# ==========================================
# 【設定エリア】
# ==========================================
FETCH_MODE = "2年分過去参照"
OUTPUT_DIR = os.path.join(r"C:\Users\sawak\OneDrive\デスクトップ\売上メール", "データベース")
EXCEL_FILE = os.path.join(OUTPUT_DIR, "売上管理表.xlsx")

KIMURA_ACCOUNT = {
    "name": "きむら",
    "server": "imap.softbank.jp",
    "user": "nakka878@i.softbank.jp",
    "password": "YpFxAeRL3H" 
}
JA_SENDER = "sanchoku@jakagawaken.jp"

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

def parse_ja_text(body, mail_year=None):
    data = []
    date_match = re.search(r'(\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}:\d{2})', body)
    if date_match:
        year = mail_year if mail_year else datetime.now().year
        date_str = f"{year}年{date_match.group(1)}"
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

def fetch_ja_outlook(since):
    print("OutlookからJAメールを検索中...")
    new_rows = []
    searched = 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"❌ Outlook起動エラー: {e}")
        return []

    def search_folder(folder, depth=0):
        nonlocal searched
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            for item in items:
                try:
                    if not hasattr(item, 'ReceivedTime'): continue
                    recv = item.ReceivedTime
                    if not hasattr(recv, 'year'): continue
                    recv_naive = datetime(recv.year, recv.month, recv.day)
                    if recv_naive < since: break
                    
                    sender = str(item.SenderEmailAddress).lower()
                    if JA_SENDER.lower() not in sender: continue
                    
                    searched += 1
                    body = item.Body or ""
                    rows = parse_ja_text(body, mail_year=recv.year)
                    if rows:
                        df = pd.DataFrame(rows)
                        df['取得元'] = 'JA'
                        new_rows.append(df)
                except Exception: continue
        except Exception: pass
        if depth == 0:
            try:
                for sub in folder.Folders:
                    search_folder(sub, depth=1)
            except Exception: pass

    for store in ns.Stores:
        try:
            inbox = store.GetDefaultFolder(6)
            search_folder(inbox)
        except Exception: continue

    print(f"  Outlook検索完了: {searched}件処理")
    return new_rows

def fetch_kimura_imap(since):
    print(f"i.softbank (IMAP) からきむらメールを検索中...")
    new_rows = []
    ssl_ctx = get_ssl_context()
    try:
        with MailBox(KIMURA_ACCOUNT['server'], ssl_context=ssl_ctx).login(KIMURA_ACCOUNT['user'], KIMURA_ACCOUNT['password']) as mailbox:
            criteria = AND(date_gte=since.date())
            for msg in mailbox.fetch(criteria, reverse=True):
                if msg.attachments:
                    for att in msg.attachments:
                        if att.filename.endswith(('.xlsx', '.xls')):
                            try:
                                df = pd.read_excel(io.BytesIO(att.payload), header=None)
                                
                                # ファイル名から日付らしき文字列の抽出を試みる（例: "10月2日.xlsx" -> "10月2日"）
                                filename = att.filename
                                date_match = re.search(r'(\d{1,2})月(\d{1,2})日', filename)
                                mail_date = msg.date
                                if date_match:
                                    # 年はメール受信年から推測（基本的には同じ年）
                                    month = int(date_match.group(1))
                                    day = int(date_match.group(2))
                                    # 年越し(例: 1月に12月のデータが届いた場合など)の考慮は一旦シンプルにメール受信年とする
                                    year = mail_date.year
                                    
                                    # メール受信月よりファイル名月が数ヶ月未来（例えば1月に12月のファイル等）の場合は前年とする
                                    if month > mail_date.month + 6:
                                        year -= 1
                                        
                                    try:
                                        mail_date = datetime(year, month, day)
                                    except ValueError:
                                        # 日付として不正(例: 2月30日など)な場合は元のメール受信日をフォールバックに使う
                                        pass
                                
                                df = normalize_kimura_df(df, fallback_date=mail_date)
                                if df is not None:
                                    df['取得元'] = "きむら"
                                    new_rows.append(df)
                            except Exception: pass
    except Exception as e:
        print(f"❌ i.softbankエラー: {e}")
    return new_rows

def main():
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    print(f"[{FETCH_MODE}] 過去2年分のデータを収集中...")
    
    two_years_ago = datetime.now() - timedelta(days=730)
    
    # JA (Outlook)
    ja_dfs = fetch_ja_outlook(two_years_ago)
    
    # きむら (IMAP)
    kimura_dfs = fetch_kimura_imap(two_years_ago)
    
    new_dfs = {
        "JA": ja_dfs,
        "きむら": kimura_dfs
    }

    old_dfs = {"きむら": None, "JA": None}
    if os.path.exists(EXCEL_FILE):
        try:
            full_df = pd.read_excel(EXCEL_FILE, sheet_name="Data")
            old_dfs["きむら"] = full_df[full_df['取得元'] == 'きむら']
            old_dfs["JA"] = full_df[full_df['取得元'] == 'JA']
        except: pass

    final_dfs = []
    for name in ["きむら", "JA"]:
        current_new = pd.concat(new_dfs[name], ignore_index=True) if new_dfs[name] else pd.DataFrame()
        current_old = old_dfs[name] if old_dfs[name] is not None else pd.DataFrame()
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

    if not final_dfs:
        print("データがありません。")
        return

    final_df = pd.concat(final_dfs, ignore_index=True)
    final_df.sort_values('日付', inplace=True)

    print("[天気] 天気を一括取得中...")
    for col in ['天気', '最高気温', '最低気温']:
        if col not in final_df.columns: final_df[col] = None
    missing_mask = (final_df['天気'].isna()) | (final_df['天気'] == "不明")
    if missing_mask.any():
        missing_dates = pd.Series(final_df.loc[missing_mask, '日付'].dt.date.unique()).dropna()
        if not missing_dates.empty:
            print(f"  {len(missing_dates)}日分の天気を取得中...")
            weather_dict = get_weather_batch(list(missing_dates))
            for d in missing_dates:
                day_str = d.strftime("%Y-%m-%d")
                w, t_max, t_min = weather_dict.get(day_str, ("不明", None, None))
                day_mask = final_df['日付'].dt.date == d
                final_df.loc[day_mask, '天気'] = w
                final_df.loc[day_mask, '最高気温'] = t_max
                final_df.loc[day_mask, '最低気温'] = t_min
            print(f"  天気取得完了")

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
        print(f"✅ 完了: {EXCEL_FILE} (2年分データ)")
    except PermissionError:
        print("❌ Excelファイルが開いています。閉じてから実行してください！")

if __name__ == "__main__":
    main()
