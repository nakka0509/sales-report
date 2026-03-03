"""
debug_ja_imap.py
JAメールのIMAP接続状態を診断するスクリプト。
受信できている直近のメールと差出人を表示します。
"""
import sys
import ssl
from imap_tools import MailBox, AND
from datetime import datetime, timedelta

# stdout をUTF-8に強制
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# === 設定（Report.pyと同じ） ===
JA_SERVER   = "imap.gmail.com"
JA_USER     = "nakka110105@gmail.com"
JA_PASSWORD = "kxmtmuaedbzhmxri"
JA_SENDER   = "sanchoku@jakagawaken.jp"
# ================================

def get_ssl_context():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try: ctx.set_ciphers('DEFAULT@SECLEVEL=1')
    except: pass
    return ctx

print("=== JA IMAP 診断 ===")
print(f"接続先  : {JA_SERVER}")
print(f"ユーザー: {JA_USER}")
print()

try:
    ssl_ctx = get_ssl_context()
    with MailBox(JA_SERVER, ssl_context=ssl_ctx).login(JA_USER, JA_PASSWORD) as mailbox:
        print("[OK] ログイン成功！")

        # 直近30日のメールを全件確認
        since = (datetime.now() - timedelta(days=30)).date()
        all_msgs = list(mailbox.fetch(AND(date_gte=since), reverse=True, limit=20))
        print(f"\n[受信箱] 直近30日のメール（最大20件）: {len(all_msgs)}件")
        print("-" * 60)
        for msg in all_msgs:
            print(f"  日付: {msg.date}  差出人: {msg.from_}  件名: {msg.subject[:40]}")

        print()

        # JAメールを差出人でフィルタ
        print(f"[検索] 差出人 '{JA_SENDER}' からのメールを検索中...")
        ja_msgs = list(mailbox.fetch(AND(from_=JA_SENDER, date_gte=since), reverse=True))
        print(f"→ JAメールヒット数: {len(ja_msgs)}件")

        if len(ja_msgs) == 0:
            print()
            print("[NG] JAメールが見つかりません。考えられる原因:")
            print("   1. JAメール（sanchoku@jakagawaken.jp）が このGmailに届いていない")
            print("      → GmailでJAメールを受信できているか確認してください")
            print("   2. JAメールは別のメールアドレスに届いている")
            print("      → どのアドレスに届いているか確認し、Report.pyの設定を変更")
            print("   3. Gmailのフィルタ/スパムでブロックされている")
        else:
            print("[OK] JAメールを受信できています！")
            for msg in ja_msgs[:5]:
                print(f"  {msg.date}  {msg.subject[:50]}")

except Exception as e:
    print(f"[ERROR] 接続エラー: {e}")
    print()
    print("考えられる原因:")
    print("  1. Gmailのアプリパスワードが無効/期限切れ")
    print("  2. imap_toolsがインストールされていない → pip install imap-tools")
    print("  3. ネットワークの問題")
