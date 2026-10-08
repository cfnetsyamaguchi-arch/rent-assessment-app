"""
物件資料PDF → 項目データ の読み取り部。

- 文字が埋め込まれたPDF … pdfplumber でテキスト抽出
- 画像（スキャン）PDF   … Tesseract（日本語）で OCR
どちらも同じ正規化をかけてから、様式に依存しにくいルールで項目を拾う。
OCR は誤読がありうるので、OCRで読んだ行や怪しい値には「要確認」を付ける。
"""

import io
import re
import unicodedata
from datetime import date

import pdfplumber

CURRENT_YEAR = date.today().year

# ── テキスト取得 ────────────────────────────────────────────────────────────
OCR_DPI = 300
OCR_PSMS = (3, 4, 6)


def ocr_available():
    try:
        import pytesseract
        return 'jpn' in pytesseract.get_languages(config='')
    except Exception:
        return False


def extract_text(file_bytes):
    """(テキスト, 読取方法) を返す。読取方法は 'テキスト' か 'OCR'。"""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        pages = [p.extract_text() or '' for p in pdf.pages]
        text = '\n'.join(pages)
        if len(re.sub(r'\s', '', text)) >= 30:
            return normalize(text), 'テキスト'

        if not ocr_available():
            raise ValueError('画像PDFですが、この環境では文字認識（OCR）が使えません。表に手入力してください。')
        import pytesseract
        ocr_pages = []
        for p in pdf.pages:
            img = p.to_image(resolution=OCR_DPI).original
            # 読み取りモードで拾える項目が違うため、3通りで読んでつなげる
            for psm in OCR_PSMS:
                ocr_pages.append(pytesseract.image_to_string(img, lang='jpn', config=f'--psm {psm}'))
    text = normalize('\n'.join(ocr_pages), ocr=True)
    if len(re.sub(r'\s', '', text)) < 30:
        raise ValueError('文字を読み取れませんでした。表に手入力してください。')
    return text, 'OCR'


RADICALS = str.maketrans({'⻄': '西', '⻑': '長', '⻘': '青', '⻝': '食', '⻤': '鬼', '⻖': '阝', '⻌': '辶'})
CJK = r'　-ヿ㐀-鿿＀-￯々〆〇ヶ'


def normalize(text, ocr=False):
    # 互換文字（⾯→面、⾻→骨、全角英数→半角など）をそろえる
    t = unicodedata.normalize('NFKC', text).translate(RADICALS)
    t = t.replace('m2', '㎡').replace('m²', '㎡')
    if ocr:
        # OCRは日本語の文字の間に空白を入れがち → 詰める
        t = re.sub(r'(?<=[%s])[ \t]+(?=[%s])' % (CJK, CJK), '', t)
        # 「10.000円」「122 ,.000」のような桁区切りの誤読をカンマに直す
        t = re.sub(r'(\d)\s*[,.]+\s*(\d{3})(?!\d)', r'\1,\2', t)
        # よくある誤読
        t = t.replace('ILDK', '1LDK').replace('IK ', '1K ').replace('ヵ月', 'ヶ月')
    return t


# ── 小道具 ──────────────────────────────────────────────────────────────────
WAREKI = {'明治': 1868, '大正': 1912, '昭和': 1926, '平成': 1989, '令和': 2019}


def wareki_to_year(s):
    for era, base in WAREKI.items():
        m = re.search(era + r'\s*(\d+|元)\s*年', s)
        if m:
            n = 1 if m.group(1) == '元' else int(m.group(1))
            return base + n - 1
    m = re.search(r'((?:19|20)\d{2})\s*[年/.]', s)
    return int(m.group(1)) if m else None


def extract_month(s):
    m = re.search(r'(?:年|/|\.)\s*(\d{1,2})\s*(?:月|/|\.|$)', s)
    if not m:
        return None
    month = int(m.group(1))
    return month if 1 <= month <= 12 else None


def parse_num(s):
    if s is None:
        return None
    s = re.sub(r'[,\s]', '', str(s))
    m = re.search(r'\d+', s)
    return int(m.group()) if m else None


def L(label):
    """ラベル文字の間に空白が入っていても一致させる（例: 賃 料 / 竣 工）。"""
    return r'\s*'.join(map(re.escape, label))


NEXT_LABEL = r'(礼金|敷金|保証金|敷引|償却|解約|管理費|共益費|水道|駐車|更新|仲介|A\s*D|AD|賃料)'


def grab_field(label_re, text, width=24):
    m = re.search(label_re + r'[：:\s]*([^\n]{0,%d})' % width, text)
    if not m:
        return None
    return re.split(NEXT_LABEL, m.group(1))[0].strip()


def to_yen(raw, rent):
    """敷金・礼金を円に正規化。「◯ヶ月」は家賃（管理費等を含まない）基準で換算。"""
    if raw is None:
        return ''
    s = re.sub(r'[,\s/]', '', str(raw))
    if not s:
        return ''
    if re.match(r'(なし|無|不要|ゼロ|[-ー―]+$)', s):
        return 0
    m = re.search(r'([\d.]+)[ヶヵケカか箇]?月', s)
    if m:
        months = float(m.group(1))
        if months == 0:
            return 0
        return int(round(rent * months)) if rent else str(raw).strip()
    m = re.search(r'([\d.]+)万', s)
    if m:
        return int(float(m.group(1)) * 10000)
    m = re.search(r'\d+(?:\.\d+)?', s)
    if m:
        v = float(m.group())
        if v == 0:
            return 0
        if v <= 3 and rent and (v * 2).is_integer():
            return int(round(rent * v))  # 「敷金 1」＝1ヶ月
        if v < 1000:
            return ''
        return int(v)
    return ''


def vote(values):
    """OCRは同じページを複数回読むので、いちばん多く出た値を採用する（同数なら先に出た方）。"""
    values = [v for v in values if v not in (None, '', 0)]
    if not values:
        return None
    return max(values, key=lambda v: (values.count(v), -values.index(v)))


def first(patterns, text, flags=0):
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m
    return None


# ── 項目の抽出 ──────────────────────────────────────────────────────────────
NOT_PROPERTY = r'(御見積|見積書|見積金額|請求書|御請求)'
PREF = r'(?:大阪府|大阪市|東京都|京都府|京都市|兵庫県|神戸市|奈良県|滋賀県|和歌山県|愛知県|名古屋市|福岡県|神奈川県|埼玉県|千葉県)'
SKIP_ADDR = r'(〒|知事|免許|TEL|Tel|tel|FAX|株式会社|ビル\d*階|本店|支店|営業)'


def parse_pdf(text, ocr=False):
    """ocr=True のときは、見つからなかった設備を × ではなく「未確認」にする（OCRは読み落としがあるため）。"""
    d = {'warn': []}
    W = d['warn']

    if re.search(NOT_PROPERTY, text) and not re.search(L('賃料'), text):
        raise ValueError('物件の募集資料ではないようです（見積書・請求書など）。対象から外しました。')

    # ── 物件名・号室
    name, room = '', ''
    m = first([L('物件名') + r'[：:\s]*([^\n]+)', r'建?\s*物\s*名' + r'[：:\s]*([^\n]+)'], text)
    if m:
        raw = m.group(1).strip()
        if not raw.startswith('※'):
            name = re.split(r'\s{2,}|\s(?=種|所在|構造|場所)', raw)[0].strip()
    # 見出し行「ザ・ウィーヴ南船場 1002」「AIA難波南 101 大阪メトロ…」型
    cands = [(a, b) for a, b in re.findall(
        r'(?:^|\s)([^\s\d|\[\]「」:：〒※()][^\s|]{1,25}?)\s+(\d{3,4})(?:号室?)?(?=\s|$)', text, re.M)
        if len(re.findall(r'[ァ-ヶ一-龥A-Z]', a)) >= 2]
    best = vote(cands)  # (名前, 号室) の組
    if not name and best:
        name, room = best
    if not room:
        rm = first([r'号室名?\s*(\d{2,4})', r'場所\s*(\d{3,4})', r'(\d{3,4})\s*号室'], text)
        if rm:
            room = rm.group(1)
        elif best and best[0] == name:
            room = best[1]
    rooms = set(re.findall(r'(\d{3,4})\s*号室', text))
    if len(rooms) > 1:
        W.append('資料に複数の号室（%s）あり' % '・'.join(sorted(rooms)))
    name = re.sub(r'\s+', ' ', name).strip(' ・|')
    if room and room not in name:
        name = f'{name} {room}号室'.strip()
    d['name'] = name

    # ── 築年月
    am = first([
        r'(?:築年月日|築年月|築年|竣\s*工|完\s*成|建築年月?|新築年月)[：:\s]*([^\n]{0,20})',
    ], text)
    year = month = None
    if am:
        year = wareki_to_year(am.group(1))
        month = extract_month(am.group(1))
    if not year:
        # ラベルを誤読した場合：入居日・公開日などを除いた「YYYY年M月」
        for line in text.splitlines():
            if re.search(r'(入居|公開|発行|作成|更新|モデル|オープン|開始|入金|日付|印刷)', line):
                continue
            ym = re.search(r'((?:19[5-9]|20[0-3])\d)\s*年\s*(\d{1,2})\s*月', line)
            if ym:
                year, month = int(ym.group(1)), int(ym.group(2))
                W.append('築年月はラベル不明のため推定')
                break
    if year:
        d['age_str'] = f'{year}年{month}月' if month else f'{year}年'

    # ── 面積
    m = first([
        r'(?:専有面積|使用部分|面積)[^\d\n]{0,6}(\d{2,3}\.\d{1,2})',
        r'(\d{2,3}\.\d{1,2})\s*(?:㎡|平米|m)',
    ], text)
    if m:
        d['area'] = float(m.group(1))

    # ── 間取
    m = first([r'間取(?:タイプ|内訳|り)?[：:\s]*([^\n]+)'], text)
    lay = None
    if m:
        lay = re.search(r'(ワンルーム|[1-9]\s*[SLDK]{1,4})', m.group(1))
    if not lay:
        lay = re.search(r'(ワンルーム|(?<![\dA-Za-z])[1-9]\s*S?L?D?K(?![A-Za-z]))', text)
    if lay:
        d['layout'] = re.sub(r'\s', '', lay.group(1))

    # ── 住所
    addr = ''
    m = re.search(L('所在地') + r'[\s:：|\[]*([^\n]+)', text)
    if m and re.search(PREF, m.group(1)) and not re.search(SKIP_ADDR, m.group(1)):
        addr = m.group(1)
    if not addr:
        for line in text.splitlines():
            if re.search(SKIP_ADDR, line):
                continue
            am2 = re.search(PREF + r'[^\s]*?\d[\d\-‐−ー丁目番地号の]*', line)
            if am2:
                addr = am2.group()
                break
    if addr:
        am3 = re.search(PREF + r'.*', addr)
        addr = (am3.group() if am3 else addr)
        addr = re.split(r'\s{2,}|[|)\]]', addr)[0].strip()
        d['address'] = addr
    if 'address' in d and re.search(r'[孤]', d['address']):
        d['address'] = d['address'].replace('大孤府', '大阪府')

    # ── 駅・徒歩（「交通」「最寄」欄を優先、なければ最短）
    pat = r'([^\s「」:：|\[\]()]{1,20}?)[」 \t]*駅[」 \t]*徒歩[ \t]*(\d{1,2})[ \t]*分'
    hit = None
    lm = re.search(r'交[ \t]*通[:： \t]*([^\n]+)', text)
    if lm:
        hm = re.search(pat, lm.group(1))
        if hm and re.split(r'[線:：」「]', hm.group(1))[-1].strip():
            hit = (hm.group(1), hm.group(2))
    if not hit:
        hits = re.findall(r'[「]([^「」\n]{1,15}?)[」]\s*駅?\s*徒歩\s*(\d{1,2})\s*分', text)
        hits += re.findall(r'([^\s「」:：|]{1,20}駅)[ \t]*徒歩[ \t]*(\d{1,2})[ \t]*分', text)
        hits = [h for h in hits if re.split(r'[線:：」「]', h[0])[-1].replace('駅', '').strip()]
        if hits:
            hit = min(hits, key=lambda x: int(x[1]))
    if hit:
        st = re.split(r'[線:：」「]', hit[0])[-1].strip()
        st = re.sub(r'駅$', '', st)
        d['station'] = st
        d['walk'] = int(hit[1])

    # ── 構造
    m = first([r'(SRC)\s*造', r'(RC)\s*造', r'(鉄骨鉄筋コンクリート)', r'(鉄筋コンクリート|コンクリート)',
               r'(軽量鉄骨)', r'(鉄骨)', r'(木)\s*造', r'(?<![A-Z])(S)\s*造'], text)
    if m:
        d['structure'] = {'鉄骨鉄筋コンクリート': 'SRC', '鉄筋コンクリート': 'RC', 'コンクリート': 'RC',
                          '軽量鉄骨': '軽量鉄骨', '鉄骨': '鉄骨', '木': '木造', 'S': '鉄骨'}.get(m.group(1), m.group(1))

    # ── 賃料
    rent = 0
    for p in [
        L('賃料') + r'[^\d\n]{0,12}([\d,]{4,9})\s*円',
        r'^([\d,]{5,9})\s+[^\n]*\n\s*' + L('賃料'),       # 「83,000 敷金…\n賃料」型
        r'([\d,]{5,9})\s*円\s*\n\s*' + L('賃料'),
        r'(?<![新険数用理告託証送料金])料[^\d\n]{0,12}([\d,]{5,9})\s*円',  # ラベルの誤読対策
    ]:
        rent = vote([parse_num(x) for x in re.findall(p, text, re.M)]) or 0
        if rent:
            break
    if rent and rent < 10000:
        W.append('賃料の値が小さすぎます')
    if re.search(r'円\s*[~〜]|万円\s*[~〜]', text):
        W.append('賃料・礼金が「〜」表記（最低額で入力）')
    d['rent'] = rent

    # ── 管理費・共益費・水道
    fees = 0
    for lab in ('管理費', '共益費'):
        fm = re.search(L(lab) + r'等?[：:\s]*([\d,]+)\s*円', text)
        if fm:
            fees += parse_num(fm.group(1)) or 0
    wm = re.search(r'水道料金?[：:\s]*([\d,]+)\s*円', text)
    water = parse_num(wm.group(1)) if wm else 0
    d['mgmt'] = fees
    d['total_rent'] = rent + fees + (water or 0) if rent else ''

    # ── 敷金・礼金
    def money(label_re):
        vals = []
        for m in re.finditer(label_re + r'[：:\s]*([^\n]{0,24})', text):
            raw = re.split(NEXT_LABEL, m.group(1))[0].strip()
            if raw.startswith('追加') or not raw:
                continue
            vals.append(to_yen(raw, rent))
        ok = [v for v in vals if v != '' and not (isinstance(v, int) and rent and v > rent * 6)]
        if ok:
            return 0 if all(v == 0 for v in ok) else (vote(ok) if vote(ok) is not None else 0)
        return vals[0] if vals else ''

    d['deposit'] = money(L('敷金') + r'(?:/?保証金)?')
    d['key_money'] = money(L('礼金'))
    for k, lab in (('deposit', '敷金'), ('key_money', '礼金')):
        v = d[k]
        if isinstance(v, int) and rent and v > rent * 6:
            W.append(f'{lab}が家賃の6ヶ月超（誤読の可能性）')
    dm = re.search(r'敷金\s*[:：]\s*([\d,]+)\s*円', text)
    if dm and isinstance(d['deposit'], int) and parse_num(dm.group(1)) != d['deposit']:
        W.append(f'敷金の記載が2通りあり（{dm.group(1)}円）')

    # ── 広告料
    m = re.search(r'(?:A\s*D|広告料|広告費|業務委託料|業務)[)）]?[^\d\n]{0,8}([\d.]+)\s*(%|[ヶヵケカか]?月)'
                  r'\s*(\([^)\n]{1,6}\))?\s*([-−ー―]\s*[\d,]+\s*円)?', text)
    if m:
        unit = m.group(2)
        unit = '%' if unit == '%' else 'ヶ月'
        ad = m.group(1) + unit
        if m.group(3):
            ad += m.group(3).replace('(', '（').replace(')', '）')
        if m.group(4):
            ad += '－' + re.sub(r'[-−ー―\s]', '', m.group(4))
        d['ad'] = ad

    # ── 設備
    flat = re.sub(r'\s', '', text)  # 改行で単語が分かれていても拾えるように

    def has(*pats):
        return any(re.search(p, text) or re.search(p, flat) for p in pats)

    eq = {
        'sepa': has(r'バス.?トイレ(別|セパ)', r'セパレート', r'BT別', r'バストイレ別'),
        'ev': has(r'エレベーター', r'(?<![A-Z])EV(?![A-Z])'),
        'al': has(r'オートロック', r'(?<![A-Z])AL(?![A-Z])'),
        'basin': has(r'洗面化粧台', r'独立洗面', r'洗面台', r'シャンプードレッサー', r'シャンドレ'),
        'laundry': has(r'室内洗濯', r'洗濯機置場', r'洗濯パン', r'洗濯乾燥機'),
        'dressing': has(r'脱衣所', r'脱衣室'),
        'bath_dry': has(r'浴室(暖房)?乾燥'),
        'reheat': has(r'追[い]?[焚炊]き?', r'追焚'),
        'shampoo': has(r'シャンプードレッサー', r'シャンドレ', r'洗髪洗面'),
        'shoe_box': has(r'シューズ\s*(ボックス|BOX)', r'下駄箱', r'玄関収納'),
        'delivery': has(r'宅配\s*(ボックス|BOX)', r'宅配(?!便)'),
        'ac': has(r'エアコン', r'冷暖房'),
        'washlet': has(r'温水洗浄便座', r'ウォシュレット'),
    }
    for k, v in eq.items():
        d[k] = '〇' if v else ('未確認' if ocr else '×')

    # キッチン（文字で表現）
    kit = []
    if has(r'システムキッチン'):
        kit.append('システム')
    km = re.search(r'(\d)\s*口', text)
    kind = 'IH' if has(r'IH') else ('ガス' if has(r'ガスコンロ', r'都市ガス.*コンロ', r'ガス\s*\d\s*口') else '')
    if km or kind:
        kit.append(f'{km.group(1) + "口" if km else ""}{kind}')
    d['kitchen'] = '\n'.join(kit)

    # 収納（文字で表現）
    sto = []
    if has(r'ウォークイン', r'WIC', r'W\.?I\.?C'):
        sto.append('WIC')
    if has(r'クローゼット', r'(?<![A-Z])CL(?![A-Z])'):
        sto.append('CL')
    if eq['shoe_box']:
        sto.append('SB')
    if has(r'床下収納'):
        sto.append('床下収納')
    d['storage'] = '・'.join(sto)

    # ── 特筆事項（判断に効く条件を短く）
    notes = []
    if fees:
        notes.append(f'賃料{rent:,}円＋共益費等{fees:,}円')
    if has(r'インターネット(無料|1G)', r'ネット無料', r'Wi-?Fi無料', r'インターネット.{0,10}無料'):
        notes.append('ネット無料')
    if has(r'家具.{0,3}家電付'):
        notes.append('家具家電付')
    pm = re.search(r'ペット\s*[:：]?\s*(不可|可|相談)', text)
    if pm:
        notes.append('ペット' + pm.group(1))
    elif has(r'小型犬\s*1\s*[匹頭]'):
        notes.append('小型犬1匹可')
    if has(r'定期借家', r'定借'):
        notes.append('定借')
    cm = re.search(r'(?:ハウス)?クリーニング(?:代|費用)?[^\d\n]{0,10}([\d,]{5,7})\s*円', text)
    if cm:
        notes.append(f'クリーニング{cm.group(1)}円')
    d['notes'] = '／'.join(notes)

    # ── 取れなかった主要項目
    for k, lab in (('name', '物件名'), ('area', '㎡数'), ('rent', '賃料'), ('address', '住所'),
                   ('station', '駅'), ('age_str', '築年'), ('deposit', '敷金'), ('key_money', '礼金')):
        if d.get(k) in (None, ''):
            W.append(f'{lab}が読み取れません')
    return d
