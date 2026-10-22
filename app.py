import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import datetime

st.set_page_config(page_title="在庫管理ポータル", layout="wide")

# ==========================================
# 簡易パスワード認証（セキュリティ保護）
# ==========================================
def check_password():
    def password_entered():
        if st.session_state["password"] == "oft2026":  # ← 社内用パスワードを設定
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.subheader("🔒 社内認証")
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.subheader("🔒 社内認証")
        st.text_input("パスワードを入力してください", type="password", on_change=password_entered, key="password")
        st.error("パスワードが正しくありません")
        return False
    return True

if not check_password():
    st.stop()

# ==========================================
# 本編アプリケーション
# ==========================================
st.title("📦 在庫管理ポータル (クラウド版)")

# 表示列の定義
def extract_target_columns(df):
    front_cols = [
        '商品コード', '商品名', '備考', '当月', '前月比', 'AVC',
        '総在庫数', '久代', '累計粗利', '評価単価', '粗利率',
        '月平均', '販売可能月', '次回入荷', '注文数量'
    ]
    month_cols = [c for c in df.columns if (isinstance(c, str) and len(c) == 7 and '.' in c and c[:4].isdigit())]
    back_cols = ['12ヶ月計', 'メーカー']
    
    ordered_cols = []
    for c in front_cols:
        if c in df.columns:
            ordered_cols.append(c)
    for c in month_cols:
        if c in df.columns:
            ordered_cols.append(c)
    for c in back_cols:
        if c in df.columns:
            ordered_cols.append(c)
            
    return df[ordered_cols].copy()

def load_and_clean_data(file_source, sheet_name):
    try:
        df = pd.read_excel(file_source, sheet_name=sheet_name, skiprows=1)
        drop_cols = [c for c in ['Unnamed: 0', '商品区分'] if c in df.columns]
        if drop_cols:
            df = df.drop(columns=drop_cols)
            
        if '商品コード' in df.columns:
            df = df.dropna(subset=['商品コード'])
            df = df[~df['商品コード'].astype(str).str.contains('<', na=False)]
            df['商品コード'] = df['商品コード'].astype(str).str.replace(r'\.0$', '', regex=True)
            
        new_cols = []
        for c in df.columns:
            if isinstance(c, (datetime.datetime, pd.Timestamp)):
                new_cols.append(c.strftime("%Y.%m"))
            else:
                new_cols.append(str(c).strip())
        df.columns = new_cols

        if '次回入荷' in df.columns:
            df['次回入荷'] = df['次回入荷'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
            df['次回入荷'] = df['次回入荷'].replace(['nan', 'None', 'NaT'], '')

        int_cols = ['総在庫数', '久代', '累計粗利', '評価単価', '当月', '前月比', 'AVC', '注文数量', '12ヶ月計']
        for col in int_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        for col in ['月平均', '販売可能月', '粗利率']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        return extract_target_columns(df), None
    except Exception as e:
        return None, str(e)

def safe_sort_dataframe(df, sort_col, is_asc):
    if sort_col not in df.columns:
        return df
    if pd.api.types.is_numeric_dtype(df[sort_col]):
        return df.sort_values(by=sort_col, ascending=is_asc, na_position='last')
    
    def sort_key(series):
        return series.apply(lambda x: (0, float(x)) if str(x).replace('.', '', 1).isdigit() else (1, str(x)))
    try:
        return df.sort_values(by=sort_col, ascending=is_asc, key=sort_key, na_position='last')
    except Exception:
        return df.sort_values(by=sort_col, ascending=is_asc, key=lambda s: s.astype(str), na_position='last')

# テーブル描画
def render_sortable_table_component(df, height=720):
    cols = df.columns.tolist()
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ margin: 0; padding: 0; background-color: #ffffff; font-family: "Meiryo", "Yu Gothic", sans-serif; }}
        .excel-table-container {{ width: 100%; height: {height - 20}px; overflow: auto; border: 1px solid #7f7f7f; box-sizing: border-box; }}
        table.excel-table {{ border-collapse: collapse; width: 100%; font-size: 12px; color: #000000; }}
        table.excel-table th {{ background-color: #d9d9d9; color: #000000; font-weight: bold; border: 1px solid #999999; padding: 6px 8px; position: sticky; top: 0; z-index: 2; text-align: center; white-space: nowrap; cursor: pointer; user-select: none; }}
        table.excel-table th:hover {{ background-color: #bfbfbf; }}
        table.excel-table th.th-stock {{ background-color: #C6EFCE !important; }}
        table.excel-table th::after {{ content: " ⇅"; font-size: 10px; color: #777777; }}
        table.excel-table td {{ border: 1px solid #d4d4d4; padding: 3px 6px; white-space: nowrap; }}
        table.excel-table tr.row-even td {{ background-color: #ffffff; }}
        table.excel-table tr.row-odd td {{ background-color: #d9e1f2; }}
        table.excel-table tr.row-even td.col-stock {{ background-color: #ffffff !important; border-left: 1px solid #70ad47 !important; border-right: 1px solid #70ad47 !important; font-weight: bold; }}
        table.excel-table tr.row-odd td.col-stock {{ background-color: #e2efda !important; border-left: 1px solid #70ad47 !important; border-right: 1px solid #70ad47 !important; font-weight: bold; }}
        table.excel-table tr:hover td {{ filter: brightness(0.92); }}
    </style>
    </head>
    <body>
    <div class="excel-table-container">
        <table id="target_table" class="excel-table">
            <thead><tr>
    """
    for c in cols:
        th_class = ' class="th-stock"' if c == '総在庫数' else ''
        html_content += f'<th{th_class}>{c}</th>'
    html_content += "</tr></thead><tbody>"

    for row_idx, (_, row) in enumerate(df.iterrows()):
        row_class = "row-even" if row_idx % 2 == 0 else "row-odd"
        html_content += f'<tr class="{row_class}">'
        for c in cols:
            val = row[c]
            td_classes = ['col-stock'] if c == '総在庫数' else []
            td_style = []
            text_align = "left"
            formatted_val = ""
            is_na = pd.isna(val) or val is None or val == '' or str(val).strip() in ['nan', 'None', 'NaT']

            if not is_na:
                if c == '販売可能月':
                    text_align = "right"
                    try:
                        num_val = float(val)
                        formatted_val = f"{num_val:.1f}"
                        if num_val < 3.0:
                            td_style.append('background-color: #FF6600 !important; color: #FFFFFF !important; font-weight: bold;')
                    except ValueError:
                        formatted_val = str(val)
                elif c == '前月比':
                    text_align = "right"
                    try:
                        num_val = float(val)
                        formatted_val = f"{num_val:,.0f}"
                        if num_val < 0:
                            td_style.append('color: #C00000 !important; font-weight: bold;')
                    except ValueError:
                        formatted_val = str(val)
                elif c in ['総在庫数', '久代', '当月', 'AVC', '注文数量', '12ヶ月計', '累計粗利', '評価単価'] or '.' in c:
                    text_align = "right"
                    try:
                        formatted_val = f"{float(val):,.0f}"
                    except ValueError:
                        formatted_val = str(val)
                elif c == '月平均':
                    text_align = "right"
                    try:
                        formatted_val = f"{float(val):.1f}"
                    except ValueError:
                        formatted_val = str(val)
                elif c == '粗利率':
                    text_align = "right"
                    try:
                        formatted_val = f"{float(val):.1%}"
                    except ValueError:
                        formatted_val = str(val)
                elif c == '商品コード':
                    text_align = "center"
                    formatted_val = str(val)
                else:
                    formatted_val = str(val)

            td_style.append(f'text-align: {text_align};')
            class_attr = f' class="{" ".join(td_classes)}"' if td_classes else ''
            style_attr = f' style="{" ".join(td_style)}"' if td_style else ''
            html_content += f'<td{class_attr}{style_attr}>{formatted_val}</td>'
        html_content += '</tr>'

    html_content += """
            </tbody>
        </table>
    </div>
    <script>
    document.addEventListener("DOMContentLoaded", function() {
        const table = document.getElementById("target_table");
        if (!table) return;
        const headers = table.querySelectorAll("th");
        headers.forEach((th, index) => {
            th.addEventListener("click", () => {
                const tbody = table.querySelector("tbody");
                const rows = Array.from(tbody.querySelectorAll("tr"));
                const isAsc = th.getAttribute("data-order") === "asc";
                headers.forEach(h => { h.removeAttribute("data-order"); h.style.color = "#000000"; });
                rows.sort((rowA, rowB) => {
                    const cellA = rowA.children[index].innerText.trim();
                    const cellB = rowB.children[index].innerText.trim();
                    if (cellA === "" && cellB !== "") return 1;
                    if (cellA !== "" && cellB === "") return -1;
                    if (cellA === "" && cellB === "") return 0;
                    const cleanA = cellA.replace(/,/g, '').replace(/%/g, '').replace(/¥/g, '');
                    const cleanB = cellB.replace(/,/g, '').replace(/%/g, '').replace(/¥/g, '');
                    const numA = parseFloat(cleanA);
                    const numB = parseFloat(cleanB);
                    if (!isNaN(numA) && !isNaN(numB)) return isAsc ? numB - numA : numA - numB;
                    return isAsc ? cellB.localeCompare(cellA, "ja") : cellA.localeCompare(cellB, "ja");
                });
                th.setAttribute("data-order", isAsc ? "desc" : "asc");
                th.style.color = "#0055ff";
                rows.forEach((row, i) => {
                    row.className = (i % 2 === 0) ? "row-even" : "row-odd";
                    tbody.appendChild(row);
                });
            });
        });
    });
    </script>
    </body>
    </html>
    """
    components.html(html_content, height=height, scrolling=False)

# ==========================================
# 画面構成
# ==========================================
st.subheader("📑 在庫管理表ビューア")

# クラウド用ファイルアップローダー
uploaded_file = st.file_uploader("📂 最新の軽量版Excelファイルをここにドラッグ＆ドロップ", type=["xlsx"])

# リポジトリ内の初期サンプルファイル（ある場合）
target_source = uploaded_file if uploaded_file else "sample_data.xlsx"

sheet_choice = st.radio("表示シート", ["商品別販売実績表", "商品別販売実績表_2minutes"], horizontal=True)

if uploaded_file or os.path.exists("sample_data.xlsx"):
    df_data, err = load_and_clean_data(target_source, sheet_choice)
    if err:
        st.error(f"データの読み込みに失敗しました: {err}")
    elif df_data is not None:
        col_s, col_k, col_o, col_r = st.columns([2.5, 1.2, 1, 1])
        with col_s:
            s_query = st.text_input("🔍 絞り込み検索（商品名 / コード / 備考 / メーカー）")
        with col_k:
            sort_options = ["(元の並び順)"] + [c for c in df_data.columns if c in ['商品コード', '販売可能月', '総在庫数', '当月', '前月比', '月平均', '粗利率', '商品名']]
            chosen_sort = st.selectbox("並び替え項目", sort_options)
        with col_o:
            chosen_order = st.selectbox("順序", ["昇順 (小→大)", "降順 (大→小)"])
        with col_r:
            st.write("")
            st.write("")
            if st.button("🔄 順番リセット"):
                st.rerun()

        filtered_df = df_data.copy()
        if s_query:
            sq = s_query.strip().lower()
            mask = filtered_df.astype(str).apply(lambda row: row.str.lower().str.contains(sq).any(), axis=1)
            filtered_df = filtered_df[mask]

        if chosen_sort != "(元の並び順)":
            is_asc = (chosen_order == "昇順 (小→大)")
            filtered_df = safe_sort_dataframe(filtered_df, chosen_sort, is_asc)

        st.caption(f"表示件数: **{len(filtered_df):,}** 行 / 全 **{len(df_data):,}** 行  ※ヘッダー項目クリックでも並び替わります。")
        render_sortable_table_component(filtered_df, height=720)
else:
    st.info("👆 上の枠に最新の軽量版Excel（`在庫管理表（xxxx）_軽量版.xlsx`）をアップロードしてください。")