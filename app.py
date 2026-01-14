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
# ユーザー様のアプリID
APP_ID = '1052224946268447244' 
REVIEW_RATE = 0.08  
PRICE_UPLIFT = 1.2  

# --- ページ設定 ---
st.set_page_config(page_title="楽天市場 運営支援ツール Suite v6", page_icon="🛍️", layout="wide")

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

    params = {"applicationId": APP_ID, "keyword": keyword, "hits": limit, "sort": "-reviewCount", "availability": 1}
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
    params = {"applicationId": APP_ID, "shopCode": shop_code, "hits": limit, "sort": "-reviewCount", "availability": 1}
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
def get_current_price_for_rpp(item_manage_number, shop_code):
    """
    修正版: itemCode指定ではなく、shopCode + keyword指定で検索する。
    これによりAPIエラー400を回避しやすくなる。
    """
    url = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20170706"
    
    # 商品管理番号から余計な文字列を除去（念のため）
    keyword = str(item_manage_number).strip()
    if ":" in keyword:
        # "lykke-hygge:abc-123" となっている場合、後半だけ取り出す
        keyword = keyword.split(":")[-1]

    params = {
        "applicationId": APP_ID,
        "shopCode": shop_code,
        "keyword": keyword, # 商品管理番号をキーワードとして検索
        "hits": 1
    }
    
    try:
        res = requests.get(url, params=params, timeout=5)
        data = res.json()
        
        if res.status_code == 200:
            if 'Items' in data and len(data['Items']) > 0:
                return data['Items'][0]['Item']['itemPrice'], "成功"
            else:
                return None, "該当なし"
        else:
            # エラー詳細を取得
            error_desc = data.get('error_description', '')
            error_msg = data.get('error', '')
            return None, f"APIエラー({res.status_code}): {error_desc} {error_msg}"
            
    except Exception as e:
        return None, f"通信エラー"

# --- ヘルパー関数: データクリーニング ---
def clean_number(val, default_val=0):
    """円、カンマ、%を取り除いて数値にする"""
    if pd.isna(val): return default_val
    s_val = str(val).replace(',', '').replace('円', '').replace('%', '').strip()
    if s_val == '' or s_val.lower() == 'nan': return default_val
    try:
        return float(s_val)
    except:
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

# ==========================================
# メインアプリケーション
# ==========================================
def main():
    st.title("楽天市場 運営支援ツール Suite v6")
    
    # 共通ID設定（内部）
    
    tab1, tab2 = st.tabs(["📊 競合分析ツール", "💰 RPP広告改善ツール"])

    # -----------------------------------
    # Tab 1: 競合分析
    # -----------------------------------
    with tab1:
        st.subheader("競合・市場調査")
        st.markdown("調査したい **キーワード、JAN、URL** を入力してください。")
        input_text = st.text_area("検索リスト", height=150, placeholder="例:\n北欧 花瓶\n4968912801046", key="comp_input")
        
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
                    
                    # Search
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

                    # Shop Analysis
                    sheet2_data = []
                    total_shops = len(analyzed_shops)
                    status_text.text(f"店舗詳細分析中... (全{total_shops}店舗)")
                    shop_map = {row['ショップコード']: row['ショップ名'] for row in sheet1_data}
                    
                    for i, s_code in enumerate(analyzed_shops):
                        s_name = shop_map.get(s_code, "不明")
                        shop_items = get_shop_top_items(s_code, s_name, limit=30)
                        sheet2_data.extend(shop_items)
                        current_progress = 40 + int((i+1) / max(1, total_shops) * 60)
                        progress_bar.progress(min(100, current_progress))

                    status_text.text("Excel生成中...")
                    df1 = pd.DataFrame(sheet1_data)
                    df2 = pd.DataFrame(sheet2_data)
                    
                    output = io.BytesIO()
                    if not df1.empty: df1 = df1.sort_values(by='推定累積売上', ascending=False)
                    cols1 = ['検索タイプ', '検索条件', '商品名', '価格', 'レビュー総数', '推定累積販売数', '推定累積売上', 'ポイント倍率', 'クーポン有無', 'ショップ名', '商品URL']
                    df1 = df1.reindex(columns=cols1) if not df1.empty else pd.DataFrame()
                    
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        if not df1.empty:
                            df1.to_excel(writer, sheet_name='検索結果', index=False)
                            format_worksheet(writer.sheets['検索結果'])
                        if not df2.empty:
                            df2.to_excel(writer, sheet_name='店舗分析', index=False)
                            format_worksheet(writer.sheets['店舗分析'])
                    
                    progress_bar.progress(100)
                    status_text.success("分析完了！")
                    
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
                    st.download_button(
                        label="📊 分析結果Excelをダウンロード",
                        data=output.getvalue(),
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
        st.markdown("RMSのパフォーマンスレポートをアップロードしてください。")

        col1, col2 = st.columns(2)
        with col1:
            my_shop_code = st.text_input("自店舗ID (URLの英数字)", value="lykke-hygge", help="例: lykke-hygge")
        with col2:
            uploaded_file = st.file_uploader("RPP実績ファイル (CSV/Excel)", type=['csv', 'xlsx', 'xls'])

        with st.expander("詳細設定", expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            target_roas = c1.number_input("目標ROAS (%)", min_value=100, value=400, step=50)
            min_cpc = c2.number_input("最低CPC (円)", min_value=10, value=25)
            max_cpc = c3.number_input("最高CPC (円)", min_value=10, value=100)
            # ★修正: デフォルトを7行目に設定
            skip_rows_num = c4.number_input("ヘッダー開始行", min_value=1, value=7, help="通常は7行目または8行目です。")

        if st.button("価格取得＆改善実行", key="rpp_btn"):
            if not uploaded_file or not my_shop_code:
                st.error("ファイルと自店舗IDは必須です。")
            else:
                try:
                    df_rpp = None
                    skip_rows_count = skip_rows_num - 1 

                    # --- 1. 読み込み ---
                    if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
                        uploaded_file.seek(0)
                        try:
                            df_rpp = pd.read_excel(uploaded_file, skiprows=skip_rows_count)
                        except: pass
                    else:
                        encodings = ['shift_jis', 'cp932', 'utf-8', 'utf-8-sig']
                        for enc in encodings:
                            try:
                                uploaded_file.seek(0)
                                df_rpp = pd.read_csv(uploaded_file, encoding=enc, skiprows=skip_rows_count)
                                if len(df_rpp.columns) > 1: break
                            except: continue
                    
                    if df_rpp is None:
                        st.error(f"読み込み失敗。ヘッダー開始行({skip_rows_num}行目)が正しいか確認してください。")
                        st.stop()
                    
                    # --- 2. 列名の正規化 ---
                    col_map = {}
                    
                    # 商品管理番号
                    for c in ['商品管理番号', '商品URL', 'item_code', 'management_no']:
                        if c in df_rpp.columns: col_map['item_code'] = c; break
                    
                    # 実績CPC
                    for c in ['実績CPC', 'クリック単価', 'CPC', '平均CPC', 'クリック単価(円)']:
                        if c in df_rpp.columns: col_map['cpc'] = c; break
                    
                    # ROAS
                    for c in ['ROAS', 'ROAS(%)', '売上対広告費比率']:
                        if c in df_rpp.columns: col_map['roas'] = c; break
                    
                    # クリック数
                    for c in ['クリック数', 'Clicks', 'クリック']:
                        if c in df_rpp.columns: col_map['clicks'] = c; break
                    
                    if 'item_code' not in col_map:
                        st.error(f"「商品管理番号」列が見つかりません。読み込んだ列名: {list(df_rpp.columns)}")
                        st.stop()

                    st.write(f"データ件数: {len(df_rpp)}件 / 読み取り列: {col_map}")
                    
                    progress_rpp = st.progress(0)
                    results_rpp = []
                    total_rows = len(df_rpp)
                    
                    for index, row in df_rpp.iterrows():
                        progress_rpp.progress((index + 1) / total_rows)
                        
                        # 商品コード
                        item_manage_number = str(row[col_map['item_code']]).strip()
                        if not item_manage_number or item_manage_number.lower() == 'nan': continue
                        
                        # 数値取得
                        current_cpc = clean_number(row.get(col_map.get('cpc')), default_val=25)
                        roas = clean_number(row.get(col_map.get('roas')), default_val=0)
                        clicks = int(clean_number(row.get(col_map.get('clicks')), default_val=0))
                        
                        # APIで価格取得
                        current_price, status_msg = get_current_price_for_rpp(item_manage_number, my_shop_code)
                        time.sleep(0.3)
                        
                        # ロジック
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
                            "APIステータス": status_msg,
                            "実績CPC": current_cpc,
                            "推奨CPC": int(new_cpc),
                            "変更理由": reason,
                            "ROAS": roas,
                            "クリック数": clicks
                        })
                    
                    if not results_rpp:
                        st.warning("処理できるデータがありませんでした。")
                    else:
                        df_res = pd.DataFrame(results_rpp)
                        st.success("完了しました！")
                        st.dataframe(df_res)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_res.to_excel(writer, sheet_name='RPP改善案', index=False)
                            format_worksheet(writer.sheets['RPP改善案'])
                        
                        st.download_button(
                            label="推奨CPCリストをダウンロード (Excel)",
                            data=output.getvalue(),
                            file_name='rpp_optimized_v6.xlsx',
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                except Exception as e:
                    st.error(f"予期せぬエラー: {e}")

if __name__ == "__main__":
    main()
