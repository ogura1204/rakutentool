import streamlit as st
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
    """商品管理番号とショップIDから現在の価格を取得"""
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
        data = res.json()
        if 'Items' in data and len(data['Items']) > 0:
            return data['Items'][0]['Item']['itemPrice']
        return None
    except:
        return None

# --- Excel生成 ---
def format_worksheet(worksheet):
    left_align = Alignment(horizontal='left', vertical='center')
    fill_color = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    hyperlink_font = Font(color="0000FF", underline="single")
    num_cols = ["価格", "レビュー総数", "推定累積販売数", "推定累積売上", "現在価格", "実績CPC", "推奨CPC"]
    
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
    if not df2.empty: df2 = df2.sort_values(by='推定累積売上', ascending=False)

    cols1 = ['検索タイプ', '検索条件', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', 'ショップ名', '商品URL']
    cols2 = ['対象店舗', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', '商品URL']
    
    df1 = df1.reindex(columns=cols1) if not df1.empty else pd.DataFrame()
    df2 = df2.reindex(columns=cols2) if not df2.empty else pd.DataFrame()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not df1.empty:
            df1.to_excel(writer, sheet_name='検索結果', index=False)
            format_worksheet(writer.sheets['検索結果'])
        if not df2.empty:
            df2.to_excel(writer, sheet_name='店舗分析', index=False)
            format_worksheet(writer.sheets['店舗分析'])
    return output.getvalue()

# ==========================================
# メインアプリケーション
# ==========================================
def main():
    st.title("楽天市場 運営支援ツール Suite")
    
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
                    excel_data = create_excel_bytes(df1, df2)
                    
                    progress_bar.progress(100)
                    status_text.success("分析完了！")
                    
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                    st.download_button(
                        label="📊 分析結果Excelをダウンロード",
                        data=excel_data,
                        file_name=f"rakuten_analysis_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

    # -----------------------------------
    # Tab 2: RPP広告改善
    # -----------------------------------
    with tab2:
        st.subheader("RPP広告 CPC自動最適化")
        st.markdown("""
        **手順:**
        1. RMSからダウンロードした「パフォーマンスレポート(RPP)」のCSVをアップロード。
        2. 自店舗のショップID(URLの英数字)を入力。
        3. 目標ROASなどを設定して実行すると、現在の価格を取得して最適なCPCを提案します。
        """)

        col1, col2 = st.columns(2)
        with col1:
            my_shop_code = st.text_input("自店舗ID (URLの英数字)", value="lykke-hygge", help="価格取得のために必要です。例: lykke-hygge")
        with col2:
            uploaded_file = st.file_uploader("RPP実績CSVをアップロード", type=['csv'])

        # 設定エリア
        with st.expander("詳細設定", expanded=True):
            c1, c2, c3 = st.columns(3)
            target_roas = c1.number_input("目標ROAS (%)", min_value=100, value=400, step=50)
            min_cpc = c2.number_input("最低CPC (円)", min_value=10, value=25)
            max_cpc = c3.number_input("最高CPC (円)", min_value=10, value=100)

        if st.button("価格取得＆改善実行", key="rpp_btn"):
            if not uploaded_file or not my_shop_code:
                st.error("CSVファイルと自店舗IDは必須です。")
            else:
                try:
                    # --- 修正: 頑丈なCSV読み込みロジック ---
                    df_rpp = None
                    encodings = ['shift_jis', 'cp932', 'utf-8', 'utf-8-sig'] # 試行する文字コード順
                    
                    for enc in encodings:
                        try:
                            uploaded_file.seek(0) # 重要: 読み込み位置を必ず先頭にリセット
                            df_rpp = pd.read_csv(uploaded_file, encoding=enc)
                            break # 成功したらループを抜ける
                        except:
                            continue # 失敗したら次の文字コードへ
                    
                    if df_rpp is None:
                        st.error("CSVファイルの読み込みに失敗しました。ファイルが破損していないか、Excelで『CSV(UTF-8)』形式で保存し直して試してください。")
                        st.stop()
                    # -------------------------------------
                    
                    st.write(f"読み込みデータ: {len(df_rpp)}件")
                    
                    # (以下、以前のロジックと同じ...)
                    progress_rpp = st.progress(0)
                    status_rpp = st.empty()
                    results_rpp = []
                    
                    total_rows = len(df_rpp)
                    
                    for index, row in df_rpp.iterrows():
                        # ... (ここから下のループ処理は変更なし) ...
                        progress_rpp.progress((index + 1) / total_rows)
                        
                        # カラム名のゆらぎ対応
                        item_manage_number = row.get('商品管理番号', row.get('商品URL', '')).strip()
                        
                        # 価格取得
                        current_price = get_current_price_for_rpp(item_manage_number, my_shop_code, APP_ID)
                        time.sleep(0.3) # API負荷軽減
                        
                        # ロジック
                        current_cpc = row['実績CPC'] if '実績CPC' in row else 25
                        roas = row['ROAS'] if 'ROAS' in row else 0
                        clicks = row['クリック数'] if 'クリック数' in row else 0
                        
                        new_cpc = current_cpc
                        reason = "維持"
                        
                        if roas == 0 and clicks > 20:
                            new_cpc = max(min_cpc, current_cpc - 10)
                            reason = "クリック過多・売上なし"
                        elif 0 < roas < target_roas:
                            new_cpc = max(min_cpc, current_cpc - 5)
                            reason = "ROAS低・抑制"
                        elif roas > (target_roas + 200):
                            new_cpc = min(max_cpc, current_cpc + 10)
                            reason = "ROAS好調・強化"
                        
                        results_rpp.append({
                            "商品管理番号": item_manage_number,
                            "現在価格": current_price if current_price else "取得失敗",
                            "実績CPC": current_cpc,
                            "推奨CPC": new_cpc,
                            "変更理由": reason,
                            "ROAS": roas,
                            "クリック数": clicks
                        })
                    
                    df_res = pd.DataFrame(results_rpp)
                    st.success("計算完了！")
                    st.dataframe(df_res)
                    
                    # CSVダウンロード
                    csv_data = df_res.to_csv(index=False).encode('shift-jis')
                    st.download_button(
                        label="推奨CPCリストをダウンロード (CSV)",
                        data=csv_data,
                        file_name='rpp_optimized.csv',
                        mime='text/csv'
                    )

                except Exception as e:
                    st.error(f"エラー: {e}")

if __name__ == "__main__":
    main()
