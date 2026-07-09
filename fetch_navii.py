#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_navii.py — 医療情報ネット「ナビイ」(厚労省 医療機能情報提供制度) の
オープンデータ（診療所）から、泌尿器科クリニックの下り(かかりつけ・逆紹介先)候補を
抽出し、各院HPと突合して【医師確認表】を出力する。

★このスクリプトは facilities.json も fetch_kosei.py の DOWN_FACILITIES も変更しない。
  臨床判断に関わるデータは、この確認表を医師がレビュー・確定してから DOWN_FACILITIES に
  反映する（半自動化の「半」＝人手の監修を必ず挟む）というのが本プロジェクトの鉄則。

使い方:
  pip install requests            # openpyxl は fetch_kosei 流用時のみ（無くても動く）
  python3 fetch_navii.py --dump-columns          # ★最初に必ず：実データの列名・スキーマ確認
  python3 fetch_navii.py                          # 抽出 → navii_candidates.csv / .html
  python3 fetch_navii.py --hp-check               # ＋各院HPを突合してキーワード根拠を付記
  python3 fetch_navii.py --local 施設票.csv 診療科票.csv   # 手動DLしたCSVを使う
  python3 fetch_navii.py --to-down-stub 確認済み.csv       # 医師採用行 → DOWN_FACILITIES用JSON

出力（確度の高い順にソート）:
  navii_candidates.csv   … 医師確認表（UTF-8-BOM, Excelでそのまま開ける）
  navii_candidates.html  … 同・見やすいHTML版（確度で色分け）

確度スコア（下り紹介先としての「泌尿器の本気度」）の階層:
  院長=泌尿器科医（名称に泌尿器/腎泌尿・HP医師紹介の泌尿器科専門医記載）＞ 非常勤泌尿器
  ＞ HPの泌尿器濃度 ＞ 診療科の絞り込み（単科ほど加点／多科は減点）。
  重みは既知の紹介先（ばんどう/とねり/井口/金町腎）が上位に来るよう較正。--hp-check で満点運用。

注意:
- ナビイのオープンデータは「公表データの一部」。膀胱鏡/LH-RH/BCG等の細かい対応列が
  そもそも収録されているかは実データ依存 → まず --dump-columns で確認し CAP_KEYWORDS を調整。
- 収録が無い項目は HP突合(--hp-check) で補い、最終的に医師が確認表で判断する。
"""
import argparse, csv, io, json, re, sys, time, zipfile
from pathlib import Path

# requests はダウンロード/ジオコーディング/HP突合でのみ使用。
# --local + --no-geocode や --to-down-stub はネット不要なので遅延import（未導入でも動く）。
try:
    import requests
except ImportError:
    requests = None

def _need_requests():
    if requests is None:
        sys.exit("この処理には requests が必要です → pip install requests")

# ============================== 設定 ==============================

# ナビイ オープンデータ（半年ごと更新。日付を上げれば最新に追従）。
# 診療所は「施設票」＋「診療科・診療時間票」の2票組。
# ベース: https://www.mhlw.go.jp/content/11121000/<ファイル名>
NAVII_DATE = "20251201"                       # 最新版の基準日（YYYYMMDD）
NAVII_BASE = "https://www.mhlw.go.jp/content/11121000/"
FACILITY_ZIP  = f"02-1_clinic_facility_info_{NAVII_DATE}.zip"    # 施設票（名称/住所/電話/URL/対応内容）
SPECIALTY_ZIP = f"02-2_clinic_speciality_hours_{NAVII_DATE}.zip" # 診療科・診療時間票（診療科目）
# 参考: 項目定義書 https://www.mhlw.go.jp/content/11121000/001306376.xlsx

# 下り(かかりつけ)候補の対象エリア。かかりつけは近さ重視 → 葛飾＋隣接(足立/江戸川)。
# 空リストなら全域。上りの AREA_FILTER(5区) とは別に、より狭く持つ。
AREA_FILTER_DOWN = ["葛飾区", "足立区", "江戸川区"]

TARGET_SPECIALTY = "泌尿器科"   # 診療科票でこの科を標榜する診療所に絞る

# 対応内容 → アプリの down cap への変換キーワード（正規化・大文字化して部分一致）。
# 施設票の対応列(あれば)やHP文言に対して照合する。実列の有無は --dump-columns で確認のこと。
# アプリHTML側は keizoku/lhrh/bcg/bscope/female/botox を凡例・プリセット実装済み。
CAP_KEYWORDS = {
    "bscope": ["膀胱鏡", "膀胱ファイバー", "膀胱内視鏡", "ぼうこう鏡"],
    "lhrh":   ["LH-RH", "LHRH", "リュープロレリン", "リュープリン", "ゴセレリン",
               "ゾラデックス", "デガレリクス", "ゴナックス", "性腺刺激ホルモン放出ホルモン"],
    "bcg":    ["BCG膀胱内注入", "膀胱内BCG", "BCG注入", "膀胱内注入", "BCG膀胱"],
    "botox":  ["ボツリヌス", "ボトックス"],
    "female": ["女性泌尿器", "女性泌尿", "女性排尿", "女性外来", "女性のための泌尿器"],
}
# 泌尿器科を標榜していれば、少なくとも良性疾患の継続フォロー(keizoku)は可能とみなす保守的既定値。
DEFAULT_CAPS = ["keizoku"]

# フラグ型の対応列（ヘッダ=「膀胱鏡」／値=「有」）を拾うための肯定値。
AFFIRM = {"有", "可", "可能", "対応", "あり", "実施", "○", "◯", "はい", "1"}

# ---- 確度スコア（下り紹介先としての「泌尿器の本気度」）用の設定 ----
# 泌尿器の診療濃度をHPから測る語彙（種類数をカウント＝ページ長に頑健）。
URO_TERMS = ["前立腺", "膀胱", "排尿", "尿路結石", "尿管結石", "PSA", "過活動膀胱",
             "尿失禁", "血尿", "泌尿器", "腎盂", "前立腺肥大", "包茎", "尿漏れ", "頻尿", "ED"]
# HP巡回でサブページ（院長・医師紹介・診療案内等）を辿るためのヒント（href/リンク文字列）。
LINK_HINTS = ["院長", "挨拶", "医師", "ドクター", "スタッフ", "診療", "案内", "泌尿", "女性",
              "資格", "経歴", "担当", "doctor", "staff", "greeting", "about", "profile",
              "message", "urology", "hinyo", "gairai", "director"]
# JSナビでリンクが辿れないサイト向け：よくあるサブページURLを推測して叩く。
GUESS_PATHS = ["doctor", "doctors", "staff", "greeting", "urology", "hinyokika", "about", "profile"]

# 否定文脈の検出（HP文言の偽陽性対策）。
# 例:「泌尿器科専門医ではありません」「膀胱鏡は対応していない」等を除外するため。
NEG_AFTER = re.compile(r"(ではありません|ではない|ではなく|行っておりません|行っていません|"
                       r"対応しておりません|対応していません|対応していない|実施しておりません|"
                       r"できません|おりません|ありません|不可|非対応|受け付けており)")
NEG_BEFORE = re.compile(r"(❌|✕|×|対応していない|対応不可|行っておりません|非対応|できない)")

# 出力
OUT_CSV  = Path("navii_candidates.csv")
OUT_HTML = Path("navii_candidates.html")

UA = {"User-Agent": "referral-finder-navii/1.0 (clinical tool; contact site owner)"}

# --- 共通ヘルパは fetch_kosei から流用（無ければローカル定義でフォールバック） -----------
try:
    # fetch_kosei は openpyxl 未導入時に sys.exit(SystemExit) するため両方を捕捉。
    from fetch_kosei import norm, geocode, WARD_CENTERS, GEOCODE_CACHE  # DRY
except (Exception, SystemExit):
    import unicodedata
    def norm(s):
        return unicodedata.normalize("NFKC", str(s or "")).replace(" ", "").replace("　", "")
    WARD_CENTERS = {
        "葛飾区": (35.7434, 139.8472), "足立区": (35.7750, 139.8046),
        "江戸川区": (35.7068, 139.8683), "墨田区": (35.7107, 139.8016),
        "荒川区": (35.7362, 139.7830), "東京都": (35.6895, 139.6917),
    }
    GEOCODE_CACHE = Path("geocode_cache.json")
    _GSI = "https://msearch.gsi.go.jp/address-search/AddressSearch?q="
    def geocode(addr, cache, enabled=True):
        if addr in cache:
            return cache[addr]
        latlng = None
        q = addr if addr.startswith("東京都") else "東京都" + addr
        if enabled and requests is not None:
            try:
                r = requests.get(_GSI + requests.utils.quote(q), headers=UA, timeout=15)
                j = r.json()
                if j:
                    lng, lat = j[0]["geometry"]["coordinates"]
                    latlng = [lat, lng]
            except Exception:
                latlng = None
            time.sleep(0.25)
        if latlng is None:
            w = next((w for w in WARD_CENTERS if w in addr), "東京都")
            lat, lng = WARD_CENTERS[w]
            latlng = [lat, lng, "approx"]
        cache[addr] = latlng
        return latlng

def nkey(s):
    """照合用キー：正規化＋ラテン大文字化（LH-RH等の大小・全半角ゆらぎ吸収）"""
    return norm(s).upper()

# ============================== 取得 ==============================

def _read_csv_members(zbytes):
    """ZIPバイト列から中の全CSVを (メンバ名, 行イテレータ) で返す。UTF-8-BOM対応。"""
    zf = zipfile.ZipFile(io.BytesIO(zbytes))
    for name in zf.namelist():
        if not name.lower().endswith(".csv"):
            continue
        raw = zf.read(name)
        text = io.TextIOWrapper(io.BytesIO(raw), encoding="utf-8-sig", newline="")
        yield name, csv.reader(text)

def _load_source(zip_name, local_path=None):
    """ローカルCSV or ナビイZIP から (メンバ名, csv.reader) を列挙。"""
    if local_path:
        p = Path(local_path)
        text = io.TextIOWrapper(io.BytesIO(p.read_bytes()), encoding="utf-8-sig", newline="")
        yield p.name, csv.reader(text)
        return
    _need_requests()
    url = NAVII_BASE + zip_name
    print(f"  DL: {url}")
    r = requests.get(url, headers=UA, timeout=600)
    r.raise_for_status()
    yield from _read_csv_members(r.content)

# ============================== 列検出 ==============================

# ヘッダは版で揺れるため部分一致ヒント（先頭ほど優先）。
HINTS = {
    "id":   ["医療機関ID", "医療機関番号", "機関コード", "通し番号", "ID", "番号"],
    "name": ["医療機関名称", "医療機関名", "名称"],
    "addr": ["所在地", "住所"],
    "tel":  ["電話番号", "電話", "連絡先"],
    "url":  ["ホームページ", "ＵＲＬ", "URL", "ウェブサイト"],
    # ★診療科目名(テキスト)を最優先。コード列「診療科目コード」を先に拾わないよう名称を先頭に。
    "spec": ["診療科目名", "診療科名", "標榜科", "診療科目", "診療科"],
}

def detect_cols(header):
    """ヘッダ行(リスト)から各キーの列indexを推定（ヒント優先順）。"""
    row = [norm(h) for h in header]
    cols = {}
    for key, hints in HINTS.items():
        for h in hints:
            hit = next((i for i, v in enumerate(row) if v and norm(h) in v), None)
            if hit is not None:
                cols[key] = hit
                break
    return cols

# ============================== 抽出本体 ==============================

def collect_specialties(local=None):
    """診療科票から施設ごとの診療科目リスト（順序保持）を集め、
    泌尿器科を標榜するIDの集合と併せて返す。
    診療科リストは確度スコアの『科の絞り込み度』判定に使う。"""
    specs = {}   # id -> [診療科名, ...]（出現順、重複除去）
    for member, reader in _load_source(SPECIALTY_ZIP, local):
        header = next(reader, None)
        if not header:
            continue
        cols = detect_cols(header)
        if "id" not in cols or "spec" not in cols:
            print(f"  [warn] 診療科票 {member}: id/spec 列が検出できず（cols={cols}）スキップ")
            continue
        i_id, i_sp = cols["id"], cols["spec"]
        for row in reader:
            if i_id >= len(row) or i_sp >= len(row):
                continue
            fid = norm(row[i_id]); nm = (row[i_sp] or "").strip()
            if not fid or not nm:
                continue
            lst = specs.setdefault(fid, [])
            if nm not in lst:
                lst.append(nm)
    ids = {fid for fid, lst in specs.items() if any(TARGET_SPECIALTY in norm(x) for x in lst)}
    print(f"  泌尿器科を標榜する診療所ID: {len(ids)} 件（全診療所 {len(specs)} 件を走査）")
    return ids, specs

def row_blob(hdr_norm, row):
    """照合用テキストを作る。値はそのまま／値が肯定的な列はヘッダ語も加える
    （フラグ型: ヘッダ「膀胱鏡」＋値「有」→ 膀胱鏡 を拾えるように）。"""
    parts = []
    for h, v in zip(hdr_norm, row):
        vv = norm(v)
        if not vv:
            continue
        parts.append(vv)
        if vv in AFFIRM:
            parts.append(h)
    return "／".join(parts)

def match_caps(blob):
    """対応内容テキスト(blob)から down cap を抽出。ヒットしたcap→根拠語 の辞書も返す。"""
    key = nkey(blob)
    caps, evidence = set(), {}
    for cap, words in CAP_KEYWORDS.items():
        hit = [w for w in words if nkey(w) in key]
        if hit:
            caps.add(cap)
            evidence[cap] = hit[0]
    return caps, evidence

def extract_candidates(urology_ids, specs=None, local=None):
    """施設票から (泌尿器科ID ∧ 対象エリア) の診療所を抽出。
    各院に診療科リスト・名称シグナル（院長=泌尿器科医のプロキシ）を付与する。"""
    specs = specs or {}
    cands = []
    for member, reader in _load_source(FACILITY_ZIP, local):
        header = next(reader, None)
        if not header:
            continue
        cols = detect_cols(header)
        if "id" not in cols or "name" not in cols or "addr" not in cols:
            print(f"  [warn] 施設票 {member}: id/name/addr 列が検出できず（cols={cols}）スキップ")
            continue
        i = cols
        hdr_norm = [norm(h) for h in header]
        for row in reader:
            def cell(k):
                j = i.get(k)
                return norm(row[j]) if (j is not None and j < len(row)) else ""
            fid, name, addr = cell("id"), cell("name"), cell("addr")
            if not fid or fid not in urology_ids:
                continue
            if AREA_FILTER_DOWN and not any(a in addr for a in AREA_FILTER_DOWN):
                continue
            blob = row_blob(hdr_norm, row)
            caps, evidence = match_caps(blob)
            sp = specs.get(fid, [])
            # 名称シグナル：①「泌尿器/腎泌尿」＝院長が泌尿器科医の強いプロキシ
            #             ②「腎」名＋(透析 or 泌尿器)標榜＝腎・透析・泌尿器クリニック
            name1 = ("腎泌尿" in name) or ("泌尿器" in name)
            name2 = ("腎" in name) and (not name1) and any(("透析" in x or "泌尿器" in x) for x in sp)
            cands.append({
                "id": fid,
                "name": name,
                "addr": addr,
                "tel": cell("tel"),
                "url": cell("url"),
                "spec": TARGET_SPECIALTY,
                "specs": sp,                 # 全診療科（絞り込み度の判定用）
                "name1": name1, "name2": name2,
                "senmon": 0, "uro": 0, "hijoukin": 0,   # HP巡回で埋める
                "caps": caps,               # ナビイ由来（keizokuは後で必ず付与）
                "evidence": evidence,       # cap→根拠語
            })
    print(f"  対象エリアの泌尿器科診療所: {len(cands)} 件")
    return cands

# ============================== HP突合＋確度シグナル採取（半自動・任意） ==============================

def _strip_tags(h):
    return re.sub(r"<[^>]+>", " ", h)

# クリニックサイトはブラウザUA以外を弾く/古いTLSのことがある。HP巡回はブラウザUAで。
HP_UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) referral-finder/1.0"}

def _url_variants(url):
    """取得失敗に備えたURL候補（https化・www有無）。順に試す。"""
    if not url.startswith("http"):
        url = "http://" + url
    out = [url]
    m = re.match(r"(https?)://(www\.)?(.+)", url)
    if m:
        _, www, rest = m.groups()
        out += [f"https://{rest}", f"https://www.{rest}"] if not www else [f"https://{rest}"]
    seen = set()
    return [u for u in out if not (u in seen or seen.add(u))]

def _fetch(url, try_variants=False):
    """1URL取得。(本文, 最終URL)。try_variants時は失敗ならhttps化・www有無も試す
    （トップページ用。サブページ/推測URLは無駄打ちを避け単発）。"""
    urls = _url_variants(url) if try_variants else [url if url.startswith("http") else "http://" + url]
    for u in urls:
        try:
            r = requests.get(u, headers=HP_UA, timeout=12, allow_redirects=True)
            if r.status_code != 200:
                continue
            return r.content.decode(r.apparent_encoding or r.encoding or "utf-8", "ignore"), r.url
        except Exception:
            continue
    return "", url

def crawl_site(url):
    """トップ＋院長/医師/診療などのサブページを巡回して本文を連結して返す。
    リンクはhref・アンカーテキスト両方で判定。JSナビでリンクが辿れない場合は
    よくあるサブページURL(doctor/greeting/urology…)を推測して補う。"""
    top, final = _fetch(url, try_variants=True)
    if not top:
        return "", True
    m = re.match(r"https?://[^/]+", final)
    host = m.group(0) if m else ""
    pages = [top]
    picked = []
    for hm in re.finditer(r'<a\b[^>]*href\s*=\s*["\']?([^"\'>\s]+)[^>]*>(.*?)</a>',
                          top, flags=re.S | re.I):
        href, txt = hm.group(1), norm(_strip_tags(hm.group(2)))
        if href.startswith(("tel:", "mailto:", "#", "javascript")):
            continue
        full = (href if href.startswith("http")
                else host + ("" if href.startswith("/") else "/") + href).split("#")[0]
        if not full.startswith(host) or full == final:
            continue
        last = full.split("/")[-1]
        if "." in last and not re.search(r'\.(html?|php|aspx?)$', last):
            continue  # 画像/PDF等は除外
        if any(k.lower() in href.lower() or k in txt for k in LINK_HINTS):
            if full not in picked:
                picked.append(full)
        if len(picked) >= 4:
            break
    # リンクで拾えない（JSナビ）ならURL推測で補完
    if len(picked) < 2 and host:
        for g in GUESS_PATHS:
            for e in ("", ".html"):
                gu = f"{host}/{g}{e}"
                if gu in picked:
                    continue
                t, _ = _fetch(gu)
                if t:
                    picked.append(gu)
                    pages.append(t)
                if len(picked) >= 4:
                    break
            if len(picked) >= 4:
                break
    for p in picked:
        if len(pages) >= 6:
            break
        t, _ = _fetch(p)
        if t and t not in pages:
            pages.append(t)
        time.sleep(0.3)
    return "\n".join(pages), False

def count_senmon(flat):
    """泌尿器科専門医の記載数。『専門医療機関』の部分一致と否定文脈を除外し、
    学会の専門医/指導医の形も本物の証拠として数える（ガイドライン引用の学会名単独は拾わない）。
    flat: タグ除去・空白除去済みの本文。"""
    n = 0
    for m in re.finditer(r"泌尿器科専門医(?!療)", flat):     # 「専門医療(機関)」を除外
        if NEG_AFTER.search(flat[m.end():m.end() + 12]):    # 「…ではありません」等
            continue
        n += 1
    n += len(re.findall(r"泌尿器科学会(?:認定)?(?:専門医|指導医)", flat))
    return n

def hp_caps(text):
    """HP本文からdown capを抽出。否定文脈（❌/対応していない/行っておりません等）は除外。"""
    flat = nkey(_strip_tags(text))
    caps, evidence = set(), {}
    for cap, words in CAP_KEYWORDS.items():
        for w in words:
            hit = False
            for m in re.finditer(re.escape(nkey(w)), flat):
                pre = flat[max(0, m.start() - 30):m.start()]
                post = flat[m.end():m.end() + 14]
                if NEG_BEFORE.search(pre) or NEG_AFTER.search(post):
                    continue
                hit = True
                break
            if hit:
                caps.add(cap)
                evidence[cap] = w
                break
    return caps, evidence

def hp_check(cands):
    """各院HPを巡回し、(1)CAP_KEYWORDSヒット (2)確度シグナル
    （泌尿器科専門医の記載数・泌尿器語の種類数・非常勤泌尿器の共起）を採取。best-effort。
    HP文言は否定文脈・部分一致による偽陽性を除外して判定する。"""
    _need_requests()
    print("[HP突合] 各院サイトを巡回（院長/医師/診療ページも／失敗は無視）")
    for idx, c in enumerate(cands):
        if not c["url"]:
            continue
        text, failed = crawl_site(c["url"])
        if failed:
            c.setdefault("hp_note", "HP取得失敗")
            continue
        flat = norm(_strip_tags(text))     # 空白除去（漢字は大小無関係）
        key = flat.upper()                 # ラテン語(ED/PSA/LH-RH)照合用
        # T1③ 泌尿器科専門医の記載（否定・部分一致を除外＝院長/常勤の泌尿器科医の客観証拠）
        c["senmon"] = count_senmon(flat)
        # T3 泌尿器語の種類数（診療濃度）
        c["uro"] = sum(1 for t in set(nkey(x) for x in URO_TERMS) if t in key)
        # T2(弱) 泌尿器×非常勤の近接共起
        c["hijoukin"] = 1 if re.search(r"泌尿器.{0,15}非常勤|非常勤.{0,15}泌尿器", flat) else 0
        caps, evidence = hp_caps(text)     # 否定文脈を除外したcap抽出
        for cap in caps:
            c["caps"].add(cap)
            c["evidence"].setdefault(cap, evidence[cap] + "(HP)")
        if (idx + 1) % 10 == 0:
            print(f"  …{idx + 1}/{len(cands)} 件巡回")
    return cands

# ============================== 確度スコア ==============================

def score_candidate(c):
    """下り紹介先としての『泌尿器の本気度』を合成スコア化し、(score, 内訳文字列) を返す。
    階層: 院長=泌尿器科医(名称/専門医) > 非常勤 > HP濃度 > 科の絞り込み。
    重みは既知の紹介先(ばんどう/とねり/井口/金町腎)が上位に来るよう較正済み。"""
    sp = c.get("specs", [])
    n = len(sp)
    pos = next((k + 1 for k, s in enumerate(sp) if "泌尿器" in s), 0)  # 泌尿器の初出位置(1始まり)
    senmon, uro, hijo = c.get("senmon", 0), c.get("uro", 0), c.get("hijoukin", 0)
    caps_extra = [x for x in c["caps"] if x != "keizoku"]
    score, br = 0, []
    if c.get("name1"):
        score += 35; br.append("名称①泌尿器+35")
    if c.get("name2"):
        score += 25; br.append("腎+透析②+25")
    v = min(senmon, 3) * 10
    if v:
        score += v; br.append(f"専門医記載{senmon}→+{v}")
    v = min(uro, 8) * 3
    if v:
        score += v; br.append(f"HP泌尿器語{uro}→+{v}")
    if hijo:
        score += 10; br.append("非常勤泌尿+10")
    if pos == 1:
        score += 10; br.append("泌尿器が主科+10")
    if n and n <= 2:
        score += 8; br.append("単科/2科+8")
    elif 2 < n <= 4:
        score += 4; br.append("〜4科+4")
    elif n >= 8:
        score -= 10; br.append("8科以上-10")
    v = min(len(caps_extra), 2) * 6
    if v:
        score += v; br.append(f"cap{sorted(caps_extra)}→+{v}")
    if c.get("hp_note"):
        br.append("(HP取得失敗)")
    return score, "／".join(br)

# ============================== 出力（医師確認表） ==============================

CSV_HEADER = ["確度", "確度内訳", "施設名", "住所", "電話", "診療科", "ナビイ抽出cap",
              "HP URL", "HP突合ヒット", "医師判定", "備考", "lat", "lng", "医療機関ID"]

def finalize(cands, do_geocode=True):
    """keizoku付与・確度スコア算出・ジオコーディングし、確度の高い順に並べて出力用の行に整える。"""
    cache = json.loads(GEOCODE_CACHE.read_text(encoding="utf-8")) if GEOCODE_CACHE.exists() else {}
    # 確度スコアを付与して降順ソート（医師が上から見ていける）
    for c in cands:
        c["_score"], c["_break"] = score_candidate(c)
    rows = []
    for c in sorted(cands, key=lambda x: (-x["_score"], x["addr"], x["name"])):
        caps = sorted(set(c["caps"]) | set(DEFAULT_CAPS))
        ll = geocode(c["addr"], cache, do_geocode)
        approx = len(ll) == 3
        ev = "；".join(f"{k}:{v}" for k, v in sorted(c["evidence"].items()))
        note = []
        if approx: note.append("位置は概算")
        if c.get("hp_note"): note.append(c["hp_note"])
        rows.append({
            "確度": c["_score"], "確度内訳": c["_break"],
            "施設名": c["name"], "住所": c["addr"], "電話": c["tel"] or "要確認",
            "診療科": c["spec"], "ナビイ抽出cap": " ".join(caps), "HP URL": c["url"],
            "HP突合ヒット": ev, "医師判定": "", "備考": "／".join(note),
            "lat": ll[0], "lng": ll[1], "医療機関ID": c["id"],
        })
    GEOCODE_CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    return rows

def write_csv(rows):
    with OUT_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_HEADER)
        w.writeheader()
        w.writerows(rows)
    print(f"  → {OUT_CSV}（{len(rows)}件・Excelでそのまま開けます。医師判定列に採用/× を記入）")

def write_html(rows):
    def esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    def band(sc):
        return "hi" if sc >= 50 else ("mid" if sc >= 20 else "lo")
    trs = []
    for r in rows:
        link = f'<a href="{esc(r["HP URL"])}" target="_blank">HP</a>' if r["HP URL"] else ""
        sc = r["確度"]
        trs.append(
            f"<tr class='{band(sc)}'><td class='sc'>{sc}</td>"
            + "".join(f"<td>{esc(r[k])}</td>" for k in
                      ["施設名", "住所", "ナビイ抽出cap", "確度内訳", "HP突合ヒット"])
            + f"<td>{link}</td><td class='judge'></td></tr>")
    html = f"""<!doctype html><html lang="ja"><meta charset="utf-8">
<title>ナビイ候補 医師確認表（葛飾区＋隣接・泌尿器科）</title>
<style>
 body{{font-family:system-ui,'Hiragino Kaku Gothic ProN',sans-serif;margin:20px;color:#222}}
 h1{{font-size:18px}} p.note{{color:#666;font-size:13px}}
 table{{border-collapse:collapse;width:100%;font-size:13px}}
 th,td{{border:1px solid #ccc;padding:5px 7px;text-align:left;vertical-align:top}}
 th{{background:#f2efe9;position:sticky;top:0}}
 td.sc{{font-weight:700;text-align:right;white-space:nowrap}}
 td.judge{{background:#fffbe6;min-width:90px}}
 tr.hi td.sc{{color:#1a7f37}} tr.mid td.sc{{color:#9a6700}} tr.lo td.sc{{color:#999}}
 tr.hi{{background:#eaf6ec}} tr.mid{{background:#fdf7e3}}
</style>
<h1>ナビイ候補 医師確認表（{TARGET_SPECIALTY}・{'・'.join(AREA_FILTER_DOWN)}）</h1>
<p class="note">{len(rows)}件・<b>確度の高い順</b>。確度＝下り紹介先としての泌尿器の本気度
（名称・泌尿器科専門医記載・HP泌尿器濃度・科の絞り込み等の合成／緑=高・黄=中）。
最右列に採用/×・修正capを記入 → 採用行を DOWN_FACILITIES へ反映してください。</p>
<table><thead><tr>
 <th>確度</th><th>施設名</th><th>住所</th><th>ナビイ抽出cap</th><th>確度内訳</th><th>HP突合ヒット</th><th>HP</th><th>医師判定</th>
</tr></thead><tbody>
{''.join(trs)}
</tbody></table></html>"""
    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"  → {OUT_HTML}")

# ============================== --dump-columns ==============================

def dump_columns(local=None):
    """実データのスキーマ確認：両票のヘッダ、泌尿器科サンプル、cap関連列の当たりを出力。"""
    hint_words = ["膀胱", "尿", "前立腺", "LH-RH", "LHRH", "BCG", "ボツリヌス",
                  "ボトックス", "女性", "泌尿", "鏡", "注入", "生検"]
    for label, zip_name in [("診療科票", SPECIALTY_ZIP), ("施設票", FACILITY_ZIP)]:
        loc = None
        if local:
            loc = local[1] if (label == "診療科票" and len(local) > 1) else local[0]
        print(f"\n========== {label}（{loc or zip_name}） ==========")
        for member, reader in _load_source(zip_name, loc):
            header = next(reader, None)
            if not header:
                continue
            cols = detect_cols(header)
            print(f"[{member}] 列数={len(header)}  検出={cols}")
            for i, h in enumerate(header):
                print(f"   {i:3d}: {h}")
            # 列名に cap 関連語を含む列を当たりとして表示
            hit_cols = [(i, h) for i, h in enumerate(header)
                        if any(nkey(w) in nkey(h) for w in hint_words)]
            if hit_cols:
                print("  --- cap関連の可能性がある列 ---")
                for i, h in hit_cols:
                    print(f"   ★{i:3d}: {h}")
            # 泌尿器科を含むサンプル行を最大3件
            i_sp = cols.get("spec")
            shown = 0
            for row in reader:
                blob = "／".join(str(c) for c in row)
                is_uro = (i_sp is not None and i_sp < len(row) and TARGET_SPECIALTY in norm(row[i_sp])) \
                         or (i_sp is None and TARGET_SPECIALTY in norm(blob))
                if is_uro:
                    print(f"  [sample] {row}")
                    shown += 1
                    if shown >= 3:
                        break
            break  # 各票は先頭CSVメンバだけ見れば十分

# ============================== --to-down-stub ==============================

def to_down_stub(csv_path):
    """医師が『医師判定』列に採用(採用/○/yes/採)を記入したCSV → DOWN_FACILITIES用JSONスタブ。"""
    p = Path(csv_path)
    adopt = {"採用", "採", "○", "o", "yes", "y", "1"}
    out = []
    with p.open(encoding="utf-8-sig", newline="") as f:
        for i, r in enumerate(csv.DictReader(f)):
            judge = norm(r.get("医師判定", "")).lower()
            if judge not in adopt:
                continue
            caps = [c for c in (r.get("ナビイ抽出cap", "").split()) if c]
            out.append({
                "id": f"dn{900+i:03d}",  # 採番は貼り付け時に振り直し前提の仮ID
                "name": r.get("施設名", ""), "dir": "down", "real": True,
                "lat": float(r["lat"]) if r.get("lat") else None,
                "lng": float(r["lng"]) if r.get("lng") else None,
                "addr": r.get("住所", ""), "tel": r.get("電話", ""),
                "url": r.get("HP URL", ""),
                "station": "", "walk": None, "bus": "", "bf": True,
                "hours": "HP参照（要確認）", "caps": caps,
                "source": "医療情報ネット(ナビイ)＋公式サイト（医師確認済み）",
            })
    print(json.dumps(out, ensure_ascii=False, indent=1))
    print(f"\n# ↑ {len(out)}件。fetch_kosei.py の DOWN_FACILITIES に貼り、id/駅/徒歩などを整えてください。",
          file=sys.stderr)

# ============================== main ==============================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--local", nargs="*", help="手動DLCSV: 施設票 [診療科票]")
    ap.add_argument("--dump-columns", action="store_true", help="実データの列名・サンプルを出力")
    ap.add_argument("--hp-check", action="store_true", help="各院HPを突合してcap根拠を補完")
    ap.add_argument("--no-geocode", action="store_true", help="ジオコーディング省略（区中心）")
    ap.add_argument("--to-down-stub", metavar="CSV", help="医師確認済CSV→DOWN_FACILITIES用JSON")
    a = ap.parse_args()

    if a.to_down_stub:
        to_down_stub(a.to_down_stub)
        return

    if a.dump_columns:
        print("[dump] ナビイ実データのスキーマ確認")
        dump_columns(a.local)
        return

    print("[1/4] 泌尿器科を標榜する診療所と診療科を収集（診療科票）")
    urology_ids, specs = collect_specialties(a.local[1] if a.local and len(a.local) > 1 else None)
    print("[2/4] 施設票から対象エリアの候補を抽出")
    cands = extract_candidates(urology_ids, specs, a.local[0] if a.local else None)
    if a.hp_check:
        print("[3/4] HP巡回（cap突合＋確度シグナル採取）")
        cands = hp_check(cands)
    else:
        print("[3/4] HP巡回はスキップ（--hp-check で有効化。確度は名称＋科構成のみで算出）")
    print("[4/4] 確度スコア算出・ジオコーディング・確認表出力（確度の高い順）")
    rows = finalize(cands, do_geocode=not a.no_geocode)
    write_csv(rows)
    write_html(rows)
    print("完了。navii_candidates.csv/html を医師がレビュー → 採用行を DOWN_FACILITIES へ。")

if __name__ == "__main__":
    main()
