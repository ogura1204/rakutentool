import streamlit as st
import requests
import pandas as pd
from datetime import datetime
import time
from urllib.parse import urlparse
import io
import re
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# ▼▼▼ 設定エリア ▼▼▼
DEFAULT_APP_ID = '1052224946268447244' 
REVIEW_RATE = 0.08  
PRICE_UPLIFT = 1.2  

# --- ページ設定 ---
st.set_page_config(page_title="楽天市場 運営支援ツール Suite", page_icon="🛍️", layout="wide")

# --- CSSスタイル ---
st.markdown("""
<style>
    .main { padding-top: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #BF0000; color: white; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #008000; color: white; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 共通・ロジック関数群
# ==========================================

def get_item_key_from_url(url):
    try:
        parsed = urlparse(url)
        path_parts = [p for p in parsed.path.split('/') if p]
        if len(path_parts) >= 2: return path_parts[-1]
        return url
    except: return url

# --- 競合分析用ロジック ---
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

def search_items(query, app_id, limit=10):
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

    params = {"applicationId": app_id, "keyword": keyword, "hits": limit, "sort": "-reviewCount", "availability": 1}
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

def get_shop_top_items(shop_code, shop_name, app_id, limit=30):
    url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706"
    params = {"applicationId": app_id, "shopCode": shop_code, "hits": limit, "sort": "-reviewCount", "availability": 1}
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

# --- RPP改善用ロジック ---
def get_current_price_for_rpp(item_manage_number, shop_code, app_id):
    url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706"
    # itemCodeは通常 shop_code:item_manage_number の形式
    item_code_param = f"{shop_code}:{item_manage_number}"
    
    params = {
        "applicationId": app_id,
        "itemCode": item_code_param,
        "hits": 1
    }
    try:
        res = requests.get(url, params=params, timeout=5)
        
        if res.status_code != 200:
            return None, f"APIエラー({res.status_code})"
            
        data = res.json()
        if 'Items' in data and len(data['Items']) > 0:
            return data['Items'][0]['Item']['itemPrice'], "成功"
        else:
            return None, "該当商品なし"
    except Exception as e:
        return None, f"通信エラー: {str(e)}"

# --- ヘルパー関数: 列名のあいまい検索 ---
def find_col_value(row, candidates, default_val=0):
    """
    rowの中から、candidatesリストに含まれる列名を探して値を返す。
    数値への変換も試みる。
    """
    for col in candidates:
        if col in row.index:
            val = row[col]
            # 値のクリーニング (円, %, カンマを除去)
            try:
                if pd.isna(val): continue
                s_val = str(val).replace(',', '').replace('円', '').replace('%', '').strip()
                if s_val == '': continue
                return float(s_val)
            except:
                continue
    return default_val

def find_col_str(row, candidates, default_val=""):
    """文字列用"""
    for col in candidates:
        if col in row.index:
            val = row[col]
            if pd.isna(val): continue
            return str(val).strip()
    return default_val

# --- Excel生成 ---
def format_worksheet(worksheet):
    left_align = Alignment(horizontal='left', vertical='center')
    fill_color = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    hyperlink_font = Font(color="0000FF", underline="single")
    num_cols = ["価格", "レビュー総数", "推定累積販売数", "推定累積売上", "現在価格", "実績CPC", "推奨CPC", "ROAS", "クリック数"]
    
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
        worksheet.column_dimensions[column].width = 15

def create_excel_bytes(df1, df2):
    output = io.BytesIO()
    if not df1.empty: df1 = df1.sort_values(by='推定累積売上', ascending=False)
    # RPP結果用フォーマット
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df1.empty:
            df1.to_excel(writer, sheet_name='検索結果', index=False)
            format_worksheet(writer.sheets['検索結果'])
        if not df2.empty:
            df2.to_excel(writer, sheet_name='分析結果', index=False)
            format_worksheet(writer.sheets['分析結果'])
    return output.getvalue()

# ==========================================
# メインアプリケーション
# ==========================================
def main():
    st.title("楽天市場 運営支援ツール Suite v4")
    
    # サイドバー設定
    st.sidebar.header("⚙️ 共通設定")
    user_app_id = st.sidebar.text_input("楽天アプリID (任意)", value="", type="password", help="空欄の場合はデフォルトIDを使用しますが、大量検索時は独自のID推奨です。")
    APP_ID = user_app_id if user_app_id else DEFAULT_APP_ID

    # タブ設定
    tab1, tab2 = st.tabs(["📊 競合分析ツール", "💰 RPP広告改善ツール"])

    # -----------------------------------
    # Tab 1: 競合分析
    # -----------------------------------
    with tab1:
        st.subheader("競合・市場調査")
        st.markdown("調査したい **キーワード、JAN、URL** を入力してください。")
        input_text = st.text_area("検索リスト", height=150, placeholder="例:\n北欧 花瓶\n4968912801046\nhttps://item.rakuten.co.jp/...", key="comp_input")
        
        if st.button("分析を開始する", key="comp_btn"):
            if not input_text.strip():
                st.warning("キーワードが入力されていません。")
            else:
                target_list = [{'query': line.strip()} for line in input_text.split('\n') if line.strip()]
                status_text = st.empty()
                progress_bar = st.progress(0)
                
                try:
                    sheet1_data = []
                    analyzed_shops = set()
                    
                    # Phase 1: Search
                    total = len(target_list)
                    for i, target in enumerate(target_list):
                        q = target['query']
                        status_text.text(f"検索中 ({i+1}/{total}): {q}")
                        items = search_items(q, APP_ID, limit=10)
                        sheet1_data.extend(items)
                        for item in items:
                            if item['ショップコード'] not in analyzed_shops:
                                analyzed_shops.add(item['ショップコード'])
                        progress_bar.progress(int((i+1) / total * 40))

                    # Phase 2: Shop Analysis
                    sheet2_data = []
                    total_shops = len(analyzed_shops)
                    status_text.text(f"店舗詳細分析中... (全{total_shops}店舗)")
                    shop_map = {row['ショップコード']: row['ショップ名'] for row in sheet1_data}
                    
                    for i, s_code in enumerate(analyzed_shops):
                        s_name = shop_map.get(s_code, "不明")
                        shop_items = get_shop_top_items(s_code, s_name, APP_ID, limit=30)
                        sheet2_data.extend(shop_items)
                        current_progress = 40 + int((i+1) / max(1, total_shops) * 60)
                        progress_bar.progress(min(100, current_progress))

                    status_text.text("Excel生成中...")
                    df1 = pd.DataFrame(sheet1_data)
                    df2 = pd.DataFrame(sheet2_data)
                    
                    # 競合分析用のExcel出力
                    output = io.BytesIO()
                    if not df1.empty: df1 = df1.sort_values(by='推定累積売上', ascending=False)
                    cols1 = ['検索タイプ', '検索条件', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', 'ショップ名', '商品URL']
                    df1 = df1.reindex(columns=cols1) if not df1.empty else pd.DataFrame()
                    
                    with pd.ExcelWriter(output, engine='open
