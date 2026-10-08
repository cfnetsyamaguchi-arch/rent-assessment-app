#!/usr/bin/env python3
"""
賃料査定書 自動入力アプリ (Streamlit版)

物件資料PDF（様式はバラバラでOK・画像PDFも可）を読み取り、
賃料査定書Excelの「競合案件」シートに追記する。

ローカル起動:
    pip install -r requirements.txt   （画像PDFには tesseract-ocr / tesseract-ocr-jpn も必要）
    streamlit run app.py
"""

import re
from datetime import date

import pandas as pd
import streamlit as st

from extractor import extract_text, ocr_available, parse_pdf
from excel_writer import append_to_template, build_new

# ── 表示・編集用の列 ────────────────────────────────────────────────────────
COLUMNS = [
    ('how', '読取'), ('warn', '要確認'),
    ('name', '物件名'), ('age_str', '築年数'), ('area', '㎡数'), ('layout', '間取'),
    ('address', '住所'), ('station', '駅'), ('walk', '徒歩'), ('structure', '構造'),
    ('rent', '家賃'), ('total_rent', '合計賃料'), ('deposit', '敷金'),
    ('key_money', '礼金'), ('ad', '広告料'), ('notes', '特筆事項'),
    ('sepa', 'セパ'), ('ev', 'ＥＶ'), ('al', 'ＡＬ'), ('basin', '洗面台'),
    ('laundry', '室内洗濯'), ('dressing', '脱衣所'), ('kitchen', 'キッチン'), ('storage', '収納'),
    ('shampoo', 'シャンドレ'), ('washlet', 'ウォシュレット'),
    ('bath_dry', '浴室乾燥'), ('reheat', '追炊き'), ('delivery', '宅配BOX'), ('ac', 'エアコン'),
    ('filename', 'ファイル'),
]
EQ_KEYS = ['sepa', 'ev', 'al', 'basin', 'laundry', 'dressing', 'shampoo', 'washlet',
           'bath_dry', 'reheat', 'delivery', 'ac']
EQ_OPTIONS = ['〇', '×', '未確認']
INT_KEYS = ['walk', 'rent', 'total_rent', 'deposit', 'key_money']
FLOAT_KEYS = ['area']
READONLY = ['how', 'warn', 'filename']


def _fmt(v):
    if v is None:
        return ''
    if isinstance(v, list):
        return '／'.join(v)
    if isinstance(v, float):
        if pd.isna(v):
            return ''
        return str(int(v)) if v.is_integer() else str(v)
    return str(v)


def to_dataframe(records):
    """編集表用。すべて文字列で持つ（数値型だと空欄が「None」表示になるため）。"""
    rows = []
    for r in records:
        row = {}
        for key, label in COLUMNS:
            v = r.get(key, '')
            if key in EQ_KEYS and v not in EQ_OPTIONS:
                v = '未確認'
            row[label] = _fmt(v)
        rows.append(row)
    df = pd.DataFrame(rows, columns=[label for _, label in COLUMNS])
    return df.astype('string').fillna('')


def from_dataframe(df):
    """編集表 → レコード。数値項目は数値に戻してExcelへ渡す。"""
    records = []
    for _, row in df.iterrows():
        rec = {}
        for key, label in COLUMNS:
            v = row.get(label, '')
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                rec[key] = ''
                continue
            s = str(v).strip()
            if not s:
                rec[key] = ''
            elif key in INT_KEYS or key in FLOAT_KEYS:
                n = pd.to_numeric(re.sub(r'[,，\s円]', '', s), errors='coerce')
                if pd.isna(n):
                    rec[key] = s  # 「1ヶ月」など数値にできない記載はそのまま
                elif key in INT_KEYS:
                    rec[key] = int(round(float(n)))
                else:
                    rec[key] = float(n)
            else:
                rec[key] = s
        if rec.get('name') or rec.get('total_rent'):
            records.append(rec)
    return records


# ── 画面 ────────────────────────────────────────────────────────────────────
st.set_page_config(page_title='賃料査定書 自動入力', page_icon='🏢', layout='wide')
st.markdown("""
<style>
  .block-container { padding-top: 2rem; max-width: 1600px; }
  h1 { font-size: 1.6rem !important; }
</style>
""", unsafe_allow_html=True)

st.title('🏢 賃料査定書 自動入力アプリ')
st.caption('物件資料PDF（様式はバラバラでOK・画像PDFも可）をアップロード → 内容を確認・修正 → 査定書Excelに追記してダウンロード')

for k, v in (('records', []), ('errors', [])):
    st.session_state.setdefault(k, v)

col1, col2 = st.columns([3, 2])
with col1:
    pdf_files = st.file_uploader('📄 物件資料PDF（複数選択可）', type=['pdf'], accept_multiple_files=True)
with col2:
    tmpl_file = st.file_uploader('📊 賃料査定書Excel（任意）', type=['xlsx'],
                                 help='渡すと「競合案件」シートの最後の行の次に、同じ書式で追記します。'
                                      '渡さない場合は新しい表を作ります。')

if not ocr_available():
    st.info('ℹ️ この環境では文字認識（OCR）が使えないため、画像PDFは読み取れません（文字が埋め込まれたPDFは読み取れます）。')

if st.button('📄 PDFから読み取り', type='primary', disabled=not pdf_files):
    records, errors = [], []
    prog = st.progress(0.0, text='読み取り中…')
    for i, f in enumerate(pdf_files):
        prog.progress(i / len(pdf_files), text=f'読み取り中… {f.name}（画像PDFは1枚30秒ほどかかります）')
        try:
            text, how = extract_text(f.getvalue())
            data = parse_pdf(text, ocr=(how == 'OCR'))
            data['how'] = how
            if how == 'OCR':
                data['warn'] = ['画像PDF（OCR）のため全項目を確認'] + data['warn']
            data['filename'] = f.name
            records.append(data)
        except Exception as e:
            errors.append(f'{f.name}: {e}')
    prog.empty()
    st.session_state.records = records
    st.session_state.errors = errors

for err in st.session_state.errors:
    st.warning('⚠️ ' + err)

if st.session_state.records:
    recs = st.session_state.records
    n_warn = sum(1 for r in recs if r.get('warn'))
    st.success(f'✅ {len(recs)} 件を読み取りました。' +
               (f'うち {n_warn} 件は「要確認」欄を見て、資料と照らして直してください。' if n_warn else ''))

    df = to_dataframe(recs)
    cfg = {}
    for key, label in COLUMNS:
        if key in EQ_KEYS:
            cfg[label] = st.column_config.SelectboxColumn(label, options=EQ_OPTIONS, width='small')
        elif key in READONLY:
            cfg[label] = st.column_config.TextColumn(label, disabled=True,
                                                     width='medium' if key == 'warn' else 'small')
        elif key in ('name', 'address', 'notes'):
            cfg[label] = st.column_config.TextColumn(label, width='medium')
        else:
            cfg[label] = st.column_config.TextColumn(label, width='small')

    edited = st.data_editor(df, column_config=cfg, num_rows='dynamic', width='stretch',
                            hide_index=True, key='editor')
    records = from_dataframe(edited)

    try:
        if tmpl_file:
            xlsx, sheet, (r1, r2) = append_to_template(tmpl_file.getvalue(), records)
            fname = re.sub(r'\.xlsx$', '', tmpl_file.name) + '_競合追加.xlsx'
            st.caption(f'📊 「{sheet}」シートの {r1}〜{r2} 行目に追記します。')
        else:
            xlsx, sheet, _ = build_new(records)
            fname = f'賃料査定書_競合案件_{date.today().isoformat()}.xlsx'
        st.download_button('⬇ Excelをダウンロード', data=xlsx, file_name=fname,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           type='primary')
    except Exception as e:
        st.error(f'Excelを作れませんでした：{e}')

    if st.button('🗑 クリア'):
        st.session_state.records = []
        st.session_state.errors = []
        st.rerun()
else:
    st.info('PDFをアップロードして「PDFから読み取り」を押してください。')

st.divider()
st.caption('※ 合計賃料＝家賃＋管理費・共益費（水道定額があれば加算）。「家賃」列は確認用で、Excelには出力しません。')
st.caption('※ 敷金・礼金は金額で出力します。「◯ヶ月」表記は **家賃（管理費等を含まない）×月数** で換算します。')
st.caption('※ 築年数は西暦の年月（例：平成3年3月 → 1991年3月）にそろえます。')
st.caption('※ 画像PDFは文字認識（OCR）で読み取ります。誤読がありうるため、設備は見つからなければ「未確認」とし、行に「要確認」を表示します。')
st.caption('※ 見積書・請求書など物件資料でないPDFは自動で除外します。')
