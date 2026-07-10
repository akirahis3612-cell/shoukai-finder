# 紹介先ファインダー（葛飾区周辺・泌尿器科）

医師外来向けの紹介先検索ツール。開発者は泌尿器科医（コードは Claude が支援）。
臨床判断に関わるデータ変更は必ず医師の確認を取ってから確定すること。

## 構成

| ファイル | 役割 |
|---|---|
| `shoukai-finder-katsushika-v3.1.html` | アプリ本体（単一HTML、Leaflet + 国土地理院淡色タイル） |
| `facilities.json` | 施設データ（アプリが raw URL から fetch。上り14＋下り47＝61件。下りは uro_level つき） |
| `fetch_kosei.py` | データ生成器。厚生局名簿＋down_facilities.json → facilities.json |
| `fetch_navii.py` | ナビイ診療所オープンデータ → 泌尿器科クリニックを抽出・確度スコア/専門度付与 → `navii_candidates.csv/html`（医師確認表）＋`down_facilities.json`（アプリ用下りデータ）を出力 |
| `down_facilities.json` | 下り施設データ（ナビイ44院・uro_level/caps/座標）。fetch_kosei が手動分とマージして facilities.json へ |
| `geocode_cache.json` | 住所→座標キャッシュ（GSIジオコーダ節約用。fetch_kosei/navii で共有） |
| `.github/workflows/update-data.yml` | 毎月5日 朝6時JST に自動更新（workflow_dispatch も可） |
| `.github/workflows/navii-candidates.yml` | ナビイ候補の医師確認表を手動生成（成果物はArtifact） |

- DATA_URL: `https://raw.githubusercontent.com/akirahis3612-cell/shoukai-finder/main/facilities.json`
- HTML内の旧 Three.js 3D地図コードはコメントアウトで保存（復元用）。

## データパイプライン（fetch_kosei.py）

1. 関東信越厚生局 `chousa/kijyun.html` から医科名簿ZIP（`shisetsu_ika_*.zip`）を取得
   - ページ構成は2026-07時点。都県別Excelは廃止され全都県一括ZIP（中に都県別xlsx、東京=prefix 13）
2. Excel解析：ヘッダ4行目、列=医療機関名称/医療機関所在地(住所)/電話番号/受理届出名称
   - 「所在地（郵便番号）」列があるため addr のヒントは「住所」優先（detect_columns はヒント優先順）
3. `CAP_RULES` で届出名称→cap 変換（AREA_FILTER: 葛飾・足立・江戸川・墨田・荒川）
4. `MANUAL_INFO` をマージ（施設名部分一致）：届出制度がない手技（HoLEP/TURP/TUL等）を
   各院公式サイトの手術実績から手動監修したもの＋公式URL。**自動再生成でも消えない設計**
5. `load_downs()`：`down_facilities.json`（ナビイ由来47院）＋手動オーバーレイ（駅/徒歩の上書き）
   ＋手動フル院（Navii未収載=立石駅前/新小岩/金町中央）をマージして連結
6. ジオコーディング：GSI AddressSearch（住所に「東京都」を補完、cache使用、失敗時は区中心+approx）

## cap 分類（2026-07 医師監修済み）

- 自動（届出あり）: robot, eswl, aus, mrifusion, snm, microwave, rfa, hydro, tese, lsc, rasc
- 手動（届出制度なし→HP実績ベース）: holep, turp, tueb, tul, pnl, pul, tot, rezum
- 定義済みだが該当ゼロ: pvp（PVP/CVP）, rrp（開腹前立腺全摘）
- 下り用: keizoku, lhrh, bcg, bscope, female, botox
- 注意: lsc/rasc の届出は産婦人科主体の施設がある（泌尿器科ルートで紹介可能かは未確認）

## 運用上の注意・ハマりどころ

- 厚生局サイトは一部環境からタイムアウトするが、**GitHub Actions ランナーからは取得成功**（実証済み）
- GitHub の Web エディタで cron 行を編集すると、インラインの説明ヒント
  （"Runs at ..."）が本文に混入して YAML が壊れることがある（一度発生・修正済み）
- raw.githubusercontent.com は叩きすぎると 429（1時間弱で解除）
- GSIジオコーダは 0.25s/req のスリープを入れて配慮。geocode_cache.json をコミットして再利用
- 名簿Excelの届出名称は NFKC 正規化＋空白除去（norm()）後にキーワード照合
- 施設名の短縮表示は HTML 内 `shortName()`（法人名prefix除去。慈恵/女子医大/都立の特例あり）

## ロードマップ（医師と合意済みの方向性）

1. **ナビイ半自動化**【概ね実装済】: 厚労省 医療機能情報提供制度のオープンデータから
   下りクリニックを抽出 → HP突合 → 確度スコア/専門度 → facilities.json 反映、まで完了。
   - `fetch_navii.py`：ナビイ診療所データ(02-1施設票/02-2診療科票, 全国一括CSV, UTF-8-BOM, 半年更新,
     最新20251201, https://www.mhlw.go.jp/content/11121000/) → 葛飾＋隣接の泌尿器科診療所を抽出。
     **初回は `--dump-columns` で実列確認**（膀胱鏡/LH-RH等の細かい列は無い＝HP突合で補う）。
   - **確度スコア/泌尿器専門度**：名称(泌尿器/腎泌尿・腎透析)・HP医師紹介の泌尿器科専門医記載・
     HP泌尿器濃度・診療科の絞り込みを合成。HP文言は否定・部分一致の偽陽性を除外。
     `uro_level`=専門/対応/処方・近隣（患者案内の軸）。既知紹介先(ばんどう/とねり/井口/金町腎)で較正。
   - **down外部化済**：fetch_navii が `down_facilities.json`(ナビイ44院・uro_level/caps/座標)を生成、
     fetch_kosei の `load_downs()` が `HAND_OVERRIDES`(駅/徒歩の上書き)＋`HAND_DOWN_FULL`
     (Navii未収載=立石駅前/新小岩/金町中央)をマージ → facilities.json(下り47)。
   - **アプリ対応済**：カードに専門度バッジ、プリセット「泌尿器専門で紹介」(uro_level=専門で絞る)。
   - 運用：`navii-candidates.yml`(手動)でHP巡回込み再生成。down_facilities.json はリポジトリにコミット。
     残:在宅/訪問診療の院が専門帯に混じる（外来紹介先とは用途違い）→ 医師判断 or 在宅フラグで別扱い。
2. **エリア拡大**: AREA_FILTER 変更だけで届出capは自動対応。都内全域だと泌尿器cap付き病院は138院
   （HP調査＝手動capの監修が主なコスト）。地図は Leaflet 化済みなので拡大に耐える
3. **クラスタリング**: 施設数が増えたら Leaflet.markercluster を追加
4. **公開**: GitHub Pages を有効化すれば院内どの端末からも URL で開ける（Settings > Pages、要オーナー操作）
5. holep/tul 等の手動capの年1回棚卸し（MANUAL_INFO のコメント参照）

## 検証方法

- ローカル: HTML をブラウザで開く（DATA_URL fetch は file:// からでも動く）
- コミット後: `https://rawcdn.githack.com/akirahis3612-cell/shoukai-finder/<sha>/shoukai-finder-katsushika-v3.1.html`
- データ再生成: `python3 fetch_kosei.py`（全自動）または `--local 名簿.xlsx`、`--dump-names` で届出名称一覧
