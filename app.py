import streamlit as st
import anthropic
import openpyxl
import re
import time
import io
from pathlib import Path

# ── 頁面設定 ──────────────────────────────────────────
st.set_page_config(
    page_title="中文 → 日文 Excel 翻譯器",
    page_icon="🗾",
    layout="centered"
)

# ── 樣式 ──────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400;600&family=Noto+Serif+TC:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Noto Serif TC', 'Noto Serif JP', serif;
}

.main { background-color: #f5f0e8; }

.title-block {
    text-align: center;
    padding: 2rem 0 1rem 0;
}

.title-jp {
    font-family: 'Noto Serif JP', serif;
    font-size: 13px;
    letter-spacing: 6px;
    color: #7a6f60;
    margin-bottom: 8px;
}

.title-main {
    font-size: 2.2rem;
    font-weight: 600;
    color: #1a1a2e;
    letter-spacing: 2px;
}

.title-main span { color: #c0392b; }

.subtitle {
    font-size: 14px;
    color: #7a6f60;
    letter-spacing: 1px;
    margin-top: 6px;
}

.stat-box {
    background: white;
    border: 1px solid #d4c9b0;
    border-radius: 4px;
    padding: 1rem;
    text-align: center;
}

.stat-num {
    font-size: 2rem;
    font-weight: 600;
    color: #1a1a2e;
}

.stat-lbl {
    font-size: 11px;
    letter-spacing: 2px;
    color: #7a6f60;
    text-transform: uppercase;
}

.stButton > button {
    background-color: #1a1a2e !important;
    color: #f5f0e8 !important;
    font-family: 'Noto Serif TC', serif !important;
    letter-spacing: 3px !important;
    border: none !important;
    padding: 0.6rem 2rem !important;
    width: 100% !important;
    font-size: 15px !important;
}

.stButton > button:hover {
    background-color: #c0392b !important;
}

.success-box {
    background: #f0fff4;
    border: 1px solid #68d391;
    border-radius: 4px;
    padding: 1rem;
    text-align: center;
    margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

# ── 標題 ──────────────────────────────────────────────
st.markdown("""
<div class="title-block">
    <div class="title-jp">中文 → 日本語 翻訳ツール</div>
    <div class="title-main">Excel <span>中→日</span> 翻譯器</div>
    <div class="subtitle">上傳 Excel，以日本會計實務用語自動翻譯成日文</div>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── 工具函數 ──────────────────────────────────────────
def has_chinese(text: str) -> bool:
    return bool(re.search(r'[\u4e00-\u9fff\u3400-\u4dbf]', str(text)))

def collect_cells(wb):
    cells = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and has_chinese(cell.value):
                    cells.append((sheet_name, cell.row, cell.column, cell.value))
    return cells

def translate_batch(texts: list, client) -> list:
    numbered = "\n".join(f"{i+1}. {t}" for i, t in enumerate(texts))
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": f"""あなたは日本の公認会計士（CPA）資格を持つ、財務・会計専門の翻訳者です。
以下の中国語テキストを、日本の会計実務で実際に使用される正式な用語・表現に翻訳してください。

【翻訳ルール】
- 会計・財務用語は日本の会計基準（日本GAAP）・税務実務で使われる正式用語を優先する
- 例：「應收帳款」→「売掛金」、「應付帳款」→「買掛金」、「存貨」→「棚卸資産」、「折舊」→「減価償却」、「淨利」→「純利益」、「資產負債表」→「貸借対照表」、「損益表」→「損益計算書」、「現金流量表」→「キャッシュ・フロー計算書」
- 数字・記号・英文はそのまま保持する
- 単純な項目名・ラベルは簡潔に訳す（余分な説明を加えない）
- 文脈から会計用語でないと判断できる場合は自然な日本語にする

【出力形式】
翻訳後の日本語テキストのみを出力する。番号・説明・原文は不要。入力と同じ {len(texts)} 行を出力する。

{numbered}"""
        }]
    )
    lines = [l.strip() for l in message.content[0].text.strip().split("\n") if l.strip()]
    while len(lines) < len(texts):
        lines.append(texts[len(lines)])
    return lines[:len(texts)]

# ── 主介面 ──────────────────────────────────────────
api_key = st.text_input(
    "🔑 Claude API Key",
    type="password",
    placeholder="sk-ant-api03-...",
    help="到 console.anthropic.com 取得 API Key"
)

uploaded_file = st.file_uploader(
    "📊 上傳 Excel 檔案",
    type=["xlsx", "xls"],
    help="支援 .xlsx / .xls，所有工作表都會翻譯"
)

translate_sheet_names = st.checkbox("同時翻譯工作表名稱", value=True)

st.divider()

if st.button("🚀 開始翻譯　翻訳開始"):
    if not api_key:
        st.error("請輸入 Claude API Key")
        st.stop()
    if not api_key.startswith("sk-ant"):
        st.error("API Key 格式不正確，請確認是否完整複製")
        st.stop()
    if not uploaded_file:
        st.error("請上傳 Excel 檔案")
        st.stop()

    try:
        client = anthropic.Anthropic(api_key=api_key)
        wb = openpyxl.load_workbook(io.BytesIO(uploaded_file.read()))
    except Exception as e:
        st.error(f"無法讀取檔案：{e}")
        st.stop()

    cells = collect_cells(wb)
    total = len(cells)

    if total == 0:
        st.warning("找不到任何含中文的儲存格！")
        st.stop()

    # 顯示統計
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f'<div class="stat-box"><div class="stat-num">{total}</div><div class="stat-lbl">含中文儲存格</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="stat-box"><div class="stat-num">{len(wb.sheetnames)}</div><div class="stat-lbl">工作表</div></div>', unsafe_allow_html=True)

    st.write("")
    progress_bar = st.progress(0)
    status_text = st.empty()
    log_area = st.empty()

    BATCH_SIZE = 20
    done = 0
    errors = 0
    log_lines = []

    for i in range(0, total, BATCH_SIZE):
        batch = cells[i:i + BATCH_SIZE]
        texts = [c[3] for c in batch]
        batch_num = i // BATCH_SIZE + 1
        total_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE

        status_text.text(f"翻譯中... 批次 {batch_num}/{total_batches}　({done}/{total} 格完成)")

        for attempt in range(2):
            try:
                translations = translate_batch(texts, client)
                for j, (sheet_name, row, col, _) in enumerate(batch):
                    wb[sheet_name].cell(row=row, column=col).value = translations[j]
                done += len(batch)
                log_lines.append(f"✅ 批次 {batch_num}/{total_batches} 完成（{len(batch)} 格）")
                break
            except Exception as e:
                if attempt == 0:
                    log_lines.append(f"⚠️ 批次 {batch_num} 錯誤，重試中... ({e})")
                    time.sleep(2)
                else:
                    log_lines.append(f"❌ 批次 {batch_num} 失敗，跳過")
                    errors += len(batch)
                    done += len(batch)

        progress_bar.progress(done / total)
        log_area.code("\n".join(log_lines[-8:]))

        if i + BATCH_SIZE < total:
            time.sleep(0.3)

    # 翻譯工作表名稱
    if translate_sheet_names:
        chinese_names = [n for n in wb.sheetnames if has_chinese(n)]
        if chinese_names:
            status_text.text("翻譯工作表名稱...")
            try:
                translated_names = translate_batch(chinese_names, client)
                for old, new in zip(chinese_names, translated_names):
                    wb[old].title = new
                    log_lines.append(f"📋 工作表：{old} → {new}")
                log_area.code("\n".join(log_lines[-8:]))
            except Exception as e:
                log_lines.append(f"⚠️ 工作表名稱翻譯失敗：{e}")

    # 輸出檔案
    status_text.text("✅ 翻譯完成！")
    progress_bar.progress(1.0)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    original_name = Path(uploaded_file.name).stem
    output_name = f"{original_name}_日本語.xlsx"

    st.markdown('<div class="success-box">🎉 翻譯完成！點下方按鈕下載</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="⬇️ 下載翻譯後的 Excel",
            data=output,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 最終統計
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("已翻譯", f"{done - errors} 格")
    c2.metric("跳過", f"{errors} 格")
    c3.metric("工作表", f"{len(wb.sheetnames)} 個")
