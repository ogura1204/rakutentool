import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
from streamlit_authenticator import Hasher
import requests
import pandas as pd
from datetime import datetime
import time
from urllib.parse import urlparse
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ▼▼▼ 設定エリア ▼▼▼
DEFAULT_APP_ID = '1052224946268447244' 
REVIEW_RATE = 0.08  
PRICE_UPLIFT = 1.2  

# --- ページ設定 ---
st.set_page_config(page_title="楽天市場 競合分析ツール Pro", page_icon="📊", layout="wide")

# --- CSSスタイル ---
st.markdown("""
<style>
    .main { padding-top: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #BF0000; color: white; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #008000; color: white; }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------
# 1. 認証・ユーザー管理ロジック
# -------------------------------------------
def load_config():
    with open('config.yaml') as file:
        config = yaml.load(file, Loader=SafeLoader)
    return config

def save_config(config):
    with open('config.yaml', 'w') as file:
        yaml.dump(config, file, default_flow_style=False)

def show_login_page():
    config = load_config()
    
    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days'],
        preauthorized=config['preauthorized']
    )

    # タブでログインと新規登録を切り替え
    tab1, tab2 = st.tabs(["🔑 ログイン", "📝 新規アカウント作成"])

    # --- ログインタブ ---
    with tab1:
        st.subheader("ログイン")
        name, authentication_status, username = authenticator.login("Login", "main")

        if authentication_status:
            return True, name, username, authenticator
        elif authentication_status is False:
            st.error("ユーザー名またはパスワードが間違っています")
            return False, None, None, None
        elif authentication_status is None:
            st.warning("ユーザー名とパスワードを入力してください")
            return False, None, None, None

    # --- 新規登録タブ ---
    with tab2:
        st.subheader("アカウント作成")
        with st.form("register_form"):
            new_email = st.text_input("メールアドレス")
            new_user = st.text_input("ユーザーID (半角英数)", placeholder="例: yamada01")
            new_name = st.text_input("担当者名")
            new_pass = st.text_input("パスワード", type="password")
            new_pass2 = st.text_input("パスワード(確認)", type="password")
            
            # 追加項目
            new_company = st.text_input("会社名")
            new_tel = st.text_input("電話番号")
            
            submit = st.form_submit_button("登録する")

            if submit:
                if not (new_email and new_user and new_name and new_pass and new_company and new_tel):
                    st.warning("すべての項目を入力してください。")
                elif new_pass != new_pass2:
                    st.warning("パスワードが一致しません。")
                elif new_user in config['credentials']['usernames']:
                    st.error("そのユーザーIDは既に使用されています。")
                else:
                    # パスワードのハッシュ化
                    hashed_pass = Hasher([new_pass]).generate()[0]
                    
                    # データの保存
                    config['credentials']['usernames'][new_user] = {
                        'email': new_email,
                        'name': new_name,
                        'password': hashed_pass,
                        'company': new_company,
                        'tel': new_tel
                    }
                    save_config(config)
                    st.success("登録が完了しました！「ログイン」タブからログインしてください。")
    
    return False, None, None, None


# -------------------------------------------
# 2. 分析ロジック (既存コード)
# -------------------------------------------
def get_item_key_from_url(url):
    try:
        parsed = urlparse(url)
        path_parts = [p for p in parsed.path.split('/') if p]
        if len(path_parts) >= 2: return path_parts[-1]
        return url
    except: return url

def calculate_metrics(item, uplift, rate):
    price = item['itemPrice']
    review_count = item['reviewCount']
    item_name = item['itemName']
    catch_copy = item.get('catchcopy', '')
    
    adj_price = int(price * uplift)
    total_sales_vol = int(review_count / rate)
    total_sales_amt = total_sales_vol * adj_price
    
    full_text = (item_name + catch_copy).replace(" ", "")
    coupon_flg = "-"
    if any(x in full_text for x in ["クーポン", "OFF", "値引", "SALE"]):
        coupon_flg = "有"
    
    return {
        "商品名": item_name, "価格": price, "ポイント倍率": item['pointRate'],
        "クーポン有無": coupon_flg, "レビュー総数": review_count,
        "推定累積販売数": total_sales_vol, "推定累積売上": total_sales_amt,
        "ショップ名": item['shopName'], "ショップコード": item['shopCode'],
        "商品URL": item['itemUrl'], "ジャンルID": item['genreId']
    }

def search_items(query, limit=10):
    url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706"
    if "http" in query:
        keyword = get_item_key_from_url(query)
        search_type = "URL検索"
    elif query.isdigit() and len(query) > 7:
        keyword = query
        search_type = "JAN検索"
    else:
        keyword = query
        search_type = "ワード検索"

    params = {"applicationId": DEFAULT_APP_ID, "keyword": keyword, "hits": limit, "sort": "-reviewCount", "availability": 1}
    try:
        time.sleep(0.5)
        res = requests.get(url, params=params, timeout=10)
        data = res.json()
        results = []
        if 'Items' in data:
            for w in data['Items']:
                metrics = calculate_metrics(w['Item'], PRICE_UPLIFT, REVIEW_RATE)
                metrics['検索条件'] = query
                metrics['検索タイプ'] = search_type
                results.append(metrics)
        return results
    except: return []

def get_shop_top_items(shop_code, shop_name, limit=30):
    url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706"
    params = {"applicationId": DEFAULT_APP_ID, "shopCode": shop_code, "hits": limit, "sort": "-reviewCount", "availability": 1}
    try:
        time.sleep(0.5)
        res = requests.get(url, params=params, timeout=10)
        data = res.json()
        results = []
        if 'Items' in data:
            for w in data['Items']:
                metrics = calculate_metrics(w['Item'], PRICE_UPLIFT, REVIEW_RATE)
                metrics['対象店舗'] = shop_name
                results.append(metrics)
        return results
    except: return []

def format_worksheet(worksheet):
    left_align = Alignment(horizontal='left', vertical='center')
    fill_color = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    hyperlink_font = Font(color="0000FF", underline="single")
    num_cols = ["価格", "レビュー総数", "推定累積販売数", "推定累積売上"]
    
    for row in worksheet.iter_rows():
        worksheet.row_dimensions[row[0].row].height = 25
        for cell in row:
            cell.alignment = left_align
            if cell.row == 1:
                cell.fill = fill_color
                continue
            
            header_val = worksheet.cell(row=1, column=cell.column).value
            if header_val in num_cols:
                cell.number_format = '#,##0'
            if header_val == "商品URL" and cell.value:
                cell.hyperlink = cell.value
                cell.font = hyperlink_font

    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = worksheet.dimensions
    
    for col in worksheet.columns:
        column = get_column_letter(col[0].column)
        header_val = col[0].value
        if header_val in ["商品名", "商品URL"]: worksheet.column_dimensions[column].width = 50
        elif header_val in ["検索条件", "ショップ名", "対象店舗"]: worksheet.column_dimensions[column].width = 25
        else: worksheet.column_dimensions[column].width = 15

def create_excel_bytes(df1, df2):
    output = io.BytesIO()
    if not df1.empty: df1 = df1.sort_values(by='推定累積売上', ascending=False)
    if not df2.empty: df2 = df2.sort_values(by='推定累積売上', ascending=False)

    cols1 = ['検索タイプ', '検索条件', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', 'ショップ名', '商品URL']
    cols2 = ['対象店舗', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', '商品URL']
    
    df1 = df1.reindex(columns=cols1)
    df2 = df2.reindex(columns=cols2) if not df2.empty else pd.DataFrame()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df1.empty:
            df1.to_excel(writer, sheet_name='検索結果(売上順)', index=False)
            format_worksheet(writer.sheets['検索結果(売上順)'])
        if not df2.empty:
            df2.to_excel(writer, sheet_name='店舗別売れ筋(売上順)', index=False)
            format_worksheet(writer.sheets['店舗別売れ筋(売上順)'])
    return output.getvalue()


# -------------------------------------------
# 3. メインアプリケーション実行
# -------------------------------------------
def main():
    # 認証チェック
    is_logged_in, name, username, authenticator = show_login_page()

    if not is_logged_in:
        st.stop() # ログインしていない場合はここで処理を止める

    # ▼▼▼ 以下、ログイン成功後に表示される画面 ▼▼▼
    
    # サイドバーにユーザー情報とログアウトボタンを表示
    with st.sidebar:
        st.write(f"ようこそ、**{name}** 様")
        config = load_config()
        user_info = config['credentials']['usernames'][username]
        st.info(f"会社名: {user_info.get('company', '-')}\n\nTEL: {user_info.get('tel', '-')}")
        authenticator.logout("ログアウト", "sidebar")

    st.title("楽天市場 競合分析ツール Pro")
    st.markdown(f"調査したい **キーワード、JAN、URL** を入力してください。")

    input_text = st.text_area("検索リスト", height=150, placeholder="例:\n4968912801046\nダクトレールファン\nhttps://item.rakuten.co.jp/...")
    
    if st.button("分析を開始する"):
        if not input_text.strip():
            st.warning("キーワードが入力されていません。")
            return

        target_list = [{'query': line.strip()} for line in input_text.split('\n') if line.strip()]
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            sheet1_data = []
            sheet2_data = []
            analyzed_shops = set()
            
            # Phase 1
            total = len(target_list)
            for i, target in enumerate(target_list):
                q = target['query']
                status_text.text(f"検索中 ({i+1}/{total}): {q}")
                items = search_items(q, limit=10)
                sheet1_data.extend(items)
                
                for item in items:
                    if item['ショップコード'] not in analyzed_shops:
                        analyzed_shops.add(item['ショップコード'])
                
                progress_bar.progress(int((i+1) / total * 40))

            # Phase 2
            total_shops = len(analyzed_shops)
            status_text.text(f"店舗詳細分析中... (全{total_shops}店舗)")
            
            shop_map = {row['ショップコード']: row['ショップ名'] for row in sheet1_data}
            
            for i, s_code in enumerate(analyzed_shops):
                s_name = shop_map.get(s_code, "不明")
                shop_items = get_shop_top_items(s_code, s_name, limit=30)
                sheet2_data.extend(shop_items)
                
                current_progress = 40 + int((i+1) / max(1, total_shops) * 60)
                progress_bar.progress(min(100, current_progress))

            status_text.text("Excelファイル生成中...")
            
            if sheet1_data:
                df1 = pd.DataFrame(sheet1_data)
                df2 = pd.DataFrame(sheet2_data)
                excel_data = create_excel_bytes(df1, df2)
                
                progress_bar.progress(100)
                status_text.success("分析完了！")
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                st.download_button(
                    label="📊 Excelをダウンロード",
                    data=excel_data,
                    file_name=f"rakuten_analysis_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("データが見つかりませんでした。")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
