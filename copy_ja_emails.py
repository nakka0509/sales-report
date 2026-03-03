"""
copy_ja_emails.py
OutlookのJAメール（sanchoku@jakagawaken.jp）を読み込み、
売上管理表.xlsxに直接データを追加するスクリプト。

※Outlook Classicがインストール・設定済みであること
"""
import re
import os
import win32com.client
import pandas as pd
from datetime import datetime, timedelta

OUTPUT_DIR = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール\データベース"
EXCEL_FILE = os.path.join(OUTPUT_DIR, "売上管理表.xlsx")
JA_SENDER = "sanchoku@jakagawaken.jp"
LATITUDE = 34.34
LONGITUDE = 134.04

WEATHER_MAP = {0: "晴れ", 1: "晴れ", 2: "曇り", 3: "曇り",
               45: "霧", 48: "霧", 51: "小雨", 53: "小雨", 55: "小雨",
               61: "雨", 63: "雨", 65: "大雨", 71: "雪", 73: "雪", 75: "大雪",
               80: "雨", 81: "雨", 82: "大雨", 95: "雷雨"}

def parse_ja_text(body, mail_year=None):
    data = []
    date_match = re.search(r'(\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}:\d{2})', body)
    if date_match:
        year = mail_year if mail_year else datetime.now().year
        date_str = f"{year}年{date_match.group(1)}"
        try:
            dt = datetime.strptime(date_str, "%Y年%m月%d日 %H:%M:%S")
        except:
            dt = datetime.now()
    else:
        return []

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

def get_weather_batch(dates):
    import requests
    if not dates:
        return {}
    sorted_dates = sorted(dates)
    start = sorted_dates[0].strftime("%Y-%m-%d")
    end   = sorted_dates[-1].strftime("%Y-%m-%d")
    url = (f"https://archive-api.open-meteo.com/v1/archive"
           f"?latitude={LATITUDE}&longitude={LONGITUDE}"
           f"&start_date={start}&end_date={end}"
           f"&daily=weather_code,temperature_2m_max,temperature_2m_min"
           f"&timezone=Asia%2FTokyo")
    try:
        import requests as req
        res = req.get(url, timeout=15).json()
        result = {}
        for i, day_str in enumerate(res['daily']['time']):
            result[day_str] = (
                WEATHER_MAP.get(res['daily']['weather_code'][i], "曇り"),
                res['daily']['temperature_2m_max'][i],
                res['daily']['temperature_2m_min'][i]
            )
        return result
    except Exception as e:
        print(f"  天気API取得エラー: {e}")
        return {}

def main():
    print("Outlookに接続中...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"❌ Outlook起動エラー: {e}")
        print("Outlook Classicが起動済みか確認してください。")
        return

    # 2年分の期間
    since = datetime.now() - timedelta(days=730)

    # 全フォルダ（受信トレイ）を検索
    new_rows = []
    searched = 0

    def search_folder(folder, depth=0):
        nonlocal searched
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # 新しい順
            for item in items:
                try:
                    if not hasattr(item, 'ReceivedTime'):
                        continue
                    recv = item.ReceivedTime
                    # datetimeに変換
                    if hasattr(recv, 'year'):
                        recv_dt = recv
                    else:
                        continue
                    # 2年以上前はスキップ
                    recv_naive = datetime(recv_dt.year, recv_dt.month, recv_dt.day)
                    if recv_naive < since:
                        break  # Sort降順なので以降は古い
                    # 差出人チェック
                    sender = str(item.SenderEmailAddress).lower()
                    if JA_SENDER.lower() not in sender:
                        continue
                    searched += 1
                    body = item.Body or ""
                    rows = parse_ja_text(body, mail_year=recv_dt.year)
                    if rows:
                        df = pd.DataFrame(rows)
                        df['取得元'] = 'JA'
                        new_rows.append(df)
                        print(f"  ✅ {recv_naive.date()} | {len(rows)}件")
                except Exception:
                    continue
        except Exception:
            pass
        # サブフォルダも検索（1階層のみ）
        if depth == 0:
            try:
                for sub in folder.Folders:
                    search_folder(sub, depth=1)
            except Exception:
                pass

    # 全アカウントの受信トレイを検索
    for store in ns.Stores:
        try:
            inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox
            print(f"\n📂 {store.DisplayName} を検索中...")
            search_folder(inbox)
        except Exception:
            continue

    print(f"\n検索済みフォルダ: JAメール {searched}件見つかりました")

    if not new_rows:
        print("新しいJAデータは見つかりませんでした。")
        return

    new_df = pd.concat(new_rows, ignore_index=True)
    new_df['日付'] = pd.to_datetime(new_df['日付'])

    # 既存データを読み込み
    if os.path.exists(EXCEL_FILE):
        full_df = pd.read_excel(EXCEL_FILE, sheet_name="Data")
        full_df['日付'] = pd.to_datetime(full_df['日付'])
        ja_old = full_df[full_df['取得元'] == 'JA'].copy()
        other = full_df[full_df['取得元'] != 'JA'].copy()
    else:
        ja_old = pd.DataFrame()
        other = pd.DataFrame()

    # JA分をマージ（日付重複は最新タイムスタンプで上書き）
    combined = pd.concat([ja_old, new_df], ignore_index=True)
    combined['TempDate'] = combined['日付'].dt.date
    max_dates = combined.groupby('TempDate')['日付'].max().reset_index()
    combined = pd.merge(combined, max_dates, on=['TempDate', '日付'], how='inner')
    del combined['TempDate']

    # 天気を一括取得
    target_cols = ['日付', '取得元', '店舗名', '品名', '単価', '数量', '小計', '天気', '最高気温', '最低気温']
    for col in ['天気', '最高気温', '最低気温']:
        if col not in combined.columns:
            combined[col] = None

    missing = combined[(combined['天気'].isna()) | (combined['天気'] == '不明')]
    if not missing.empty:
        dates = missing['日付'].dt.date.unique().tolist()
        print(f"☁ 天気を一括取得中（{len(dates)}日分）...")
        wd = get_weather_batch(dates)
        for d in dates:
            day_str = d.strftime("%Y-%m-%d")
            w, tmax, tmin = wd.get(day_str, ('不明', None, None))
            mask = combined['日付'].dt.date == d
            combined.loc[mask, '天気'] = w
            combined.loc[mask, '最高気温'] = tmax
            combined.loc[mask, '最低気温'] = tmin

    final_df = pd.concat([other, combined], ignore_index=True)
    final_df.sort_values('日付', inplace=True)
    out_cols = [c for c in target_cols if c in final_df.columns]
    final_df = final_df[out_cols]

    # 保存
    try:
        if os.path.exists(EXCEL_FILE):
            with pd.ExcelWriter(EXCEL_FILE, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                final_df.to_excel(writer, sheet_name="Data", index=False)
        else:
            with pd.ExcelWriter(EXCEL_FILE, mode='w', engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name="Data", index=False)
        print(f"\n✅ 完了: {len(combined)}件のJAデータを保存しました → {EXCEL_FILE}")
    except PermissionError:
        print("❌ Excelファイルが開いています。閉じてから再実行してください。")

if __name__ == "__main__":
    main()
