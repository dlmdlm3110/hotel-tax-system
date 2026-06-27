# 宿泊税スマート — CLAUDE.md

## プロジェクト概要

沖縄県北谷町の宿泊施設向け **宿泊税申告支援システム**。Streamlit で構築されたローカル動作の Web アプリ。  
`app.py` 1ファイルが本体（1,897行）。起動は `launch.command`（Mac）/ `launch.bat`（Windows）。

### 主な機能（タブ構成）

| タブ | 内容 |
|---|---|
| タブ1 データ読み込み | ホテルシステム CSV / Airbnb CSV を最大3ヶ月分アップロード |
| タブ2 計算確認 | 日別宿泊者数・金額の手動修正、予約別計算ステップの検証（4セクション） |
| タブ3 書類出力 | 納入申告書・月計表・納入書を Excel でダウンロード |
| タブ4 領収書発行 | タイプA〜D の領収書を1件ずつ発行 |

### 配布バリエーション（デスクトップに ZIP）

| ファイル名 | 内容 |
|---|---|
| `宿泊税スマートmax.zip` | 両 CSV 対応・全機能 |
| `宿泊税スマートねっぱんvr.zip` | ホテルシステム CSV 専用（`parse_airbnb_csv` 除去済み） |
| `宿泊税スマートAirbnbvr.zip` | Airbnb CSV 専用（`parse_hotel_csv`, `parse_hotel_total_csv`, `parse_hotel_csv_for_verification`, `AGODA_COMMISSION` 除去済み） |

**最終ビルド: 2026-06-13**

---

## 重要な定数・設定値

```python
TAX_RATE = 0.02               # 宿泊税率 2%
MAX_TAX_PER_PERSON = 2000     # 上限 2,000円/人
AGODA_COMMISSION = 0.12       # Agoda手数料 12%（CSV は手数料控除後の金額で入る）
```

課税標準額の計算式：`floor(宿泊料金/人 / 1000) * 1000 * 2% * 人数`（上限2,000円/人）  
**必ず予約1件・1泊ごとに個別計算すること**（日次合計の平均割りは過小計算になる）

---

## 対応 CSV フォーマット

### ① ホテルシステム CSV・日別明細形式（標準）

`detect_csv_format` 戻り値: `'hotel'` / `parse_hotel_csv` で処理。

| 項目 | 仕様 |
|---|---|
| 列数 | 43列 |
| エンコーディング | cp932（`errors='replace'`） |
| 日付形式 | MM/DD/YYYY または YYYY/M/D（`_parse_hotel_date` で両対応） |
| データ粒度 | **宿泊日基準**（1予約×1泊×1料金区分が1行） |
| 複数行/泊目 | 同一(予約ID col0, 泊目 col3)に大人行・子供行が分かれる → グループ化して金額を集約 |

**重要な列インデックス:**

| col | 内容 |
|---|---|
| parts[0] | 予約ID（グループキー） |
| parts[2] | 予約番号（表示用） |
| parts[3] | 泊目（1始まり） |
| parts[4] | チェックイン日（MM/DD/YYYY） |
| parts[7] | 総泊数 |
| parts[8] | 予約サイト（`'agoda'` 判定に使用、case insensitive） |
| parts[12] | 宿泊者名 |
| parts[18] | 大人人数 |
| parts[19] | 子供人数（幼児は別列で課税対象外） |
| parts[25] | 料金合計額（全泊合計・フォールバック用） |
| parts[36] | 大人合計額（**その夜1泊分の金額**・ダイナミックプライシング対応） |
| parts[37] | 子供合計額（その夜1泊分） |

**宿泊日の計算:**
```python
ci         = _parse_hotel_date(parts[4])   # MM/DD/YYYY と YYYY/M/D 両対応
night_date = ci + timedelta(days=night_num - 1)
```

**月またぎの扱い:**  
月をまたぐ予約は両月のCSVに全泊分の行が入る（宿泊日基準）。  
`if night_date.month != target_month: continue` でフィルタリングし対象月の夜だけ採用。月の前後どちらの月またぎも正しく動作する。

**金額フォールバック:**  
`parts[36] + parts[37] == 0` のとき `round(parts[25] / parts[7])` で均等按分。

**Agoda補正:**  
`parts[8].lower() == 'agoda'` のとき逆算: `amount = round(amount / (1 - AGODA_COMMISSION))`

### ①' ホテルシステム CSV・集計形式（ReservationTotalList）

`detect_csv_format` 戻り値: `'hotel_total'` / `parse_hotel_total_csv` で処理。**検証機能は非対応**（小計行がないため）。

| 項目 | 仕様 |
|---|---|
| 列数 | 75列（データ行の列数 > 60 で判定） |
| エンコーディング | cp932 |
| 日付形式 | YYYY/MM/DD（例: `2026/05/31`） |
| データ粒度 | **1行=1予約**（合計金額 ÷ 泊数 で1泊あたりを算出） |
| col[3] | **空（泊目なし）** ← 43列形式との判別ポイント |
| col[7] | 総泊数 |
| col[36]/[37] | **全泊分の合計額**（÷ 泊数 で1泊あたりを算出） |

ねっぱんの「予約一覧エクスポート」画面で誤って「ReservationTotalList」形式を選んだ場合に出力される。  
正しくは「ReservationList（日別明細）」を選ぶこと。集計形式でも宿泊税計算は可能だが検証機能は使えない。

### ② Airbnb CSV

- エンコーディング: UTF-8-sig / UTF-8 / Shift-JIS を順に試みる
- `'Confirmation code'` が1行目にあれば Airbnb と判定（`detect_csv_format` 関数）
- 1予約行を宿泊日ごとに展開、`Earnings ÷ 泊数` で1泊あたり金額を算出
- Agoda補正なし

---

## 日別集計の計算方式（重要・バグ修正済み 2026-06-07）

**問題**: 旧実装は常に `recompute_daily_from_edit`（日次合計÷日次人数）を使用。同日に複数の異なる料金予約がある場合、千円切捨てのタイミングがズレて課税標準額が過小になっていた。

**現在の正しいフロー:**
1. `compute_daily_agg(items)` → 予約1件ずつ `calc_tax` → 日別に加算 → `_daily_csv`
2. `_daily_csv` を data_editor の初期値として表示
3. ユーザー編集後: 編集前後で `guests` / `amount` が変化した日のみ `recompute_daily_from_edit` で再計算、変化がない日は `_daily_csv` の値をそのまま使用

---

## 検証セクション（タブ2）の4セクション

`parse_hotel_csv_for_verification(file_bytes, year, month, meal_per_person, tax_inclusive)` は **4値を返す**:
```python
return detail_rows, daily_subtotals, excluded_rows, agoda_rows
```

| セクション | 内容 |
|---|---|
| ① 日別集計 | 日ごとの人数・売上・課税標準額・宿泊税 |
| ② 予約別計算ステップ | 1予約×1泊ごとの計算過程（これが正規の数字） |
| ③ 月またぎ除外データ | 対象月外として除外した宿泊夜の一覧（監査用） |
| ④ Agoda補正明細 | CSV記載額→グロスアップ後金額と差額（監査用） |

検証Excelは4シート構成（`②予約別計算明細` / `①日別集計` / `③月またぎ除外` / `④Agoda補正`）。  
**集計形式（hotel_total）では検証不可**。タブ2に「日別明細形式のCSVをご利用ください」と表示される。

---

## ファイル構成

```
tax/
├── app.py                    ← メインアプリ（1,897行）
├── _build_zip.py             ← 配布ZIP生成スクリプト
├── requirements.txt          ← streamlit, openpyxl, pandas, Pillow
├── settings.json             ← 施設情報（配布ZIPには含めない）
├── 納入申告書テンプレ.xlsx
├── 宿泊税納入書テンプレ.xlsx
├── 宿泊税月計表テンプレ.xlsx
├── 領収書タイプA税込.xlsx
├── 領収書タイプB税別途.xlsx
├── 領収書タイプC税全コミ.xlsx
├── 領収書タイプD宿泊税（税のみ）.xlsx
├── setup.command / setup.bat  ← 初回セットアップ
└── launch.command / launch.bat ← 起動スクリプト
```

### Excel テンプレートのセルマッピング（generate_receipt）

| タイプ | 施設名セル | 宿泊者名 | 人数 | 税額 | 日付 |
|---|---|---|---|---|---|
| A（税込） | row20-22 col H | row12 col B | row12 col D | row12 col F | row15 col D/F/H |
| B（税別途） | D19 | D3 | D4 | D7/D11 | D16（文字列） |
| C（税全コミ） | D19 | D3 | D4 | D7/D11 | D16（文字列） |
| D（税のみ） | D24 ← 注意（D23ではない） | B4（merged） | D4 | D8/D12 | D3 |

月計表: B1=指定番号、B2=施設名、A3=令和年、C3=月、row7〜37（day1〜31）col B=宿泊者数/col C=基本料金

---

## バリアント生成方法

ZIP を作り直す際は `python3 _build_zip.py` を実行（`tax/` ディレクトリで）。  
`ast.parse()` で構文チェックしてから ZIP 化する。

### 除去すべき関数・定数

| バリアント | 除去する関数・定数 |
|---|---|
| ねっぱんvr | `detect_csv_format`, `parse_airbnb_csv` |
| Airbnbvr | `detect_csv_format`, `parse_hotel_csv`, `parse_hotel_total_csv`, `parse_hotel_csv_for_verification`, `AGODA_COMMISSION` 定数 |

`detect_csv_format` を除去した場合はアップロードループ内の分岐も削除し、単一形式に固定すること（`_build_zip.py` の `LOOP_NB` / `LOOP_AB` 定数を参照）。

---

## 現在の作業状況（2026-06-13 時点）

### 完了済み

- [x] 旧ねっぱん CSV 形式（summaryView・ReservationList）を完全削除し、新ホテルシステム CSV（43列）に一本化
- [x] 日付形式 MM/DD/YYYY と YYYY/M/D の両対応（`_parse_hotel_date` ヘルパー）
- [x] 宿泊日基準 CSV の月またぎ対応（`night_date` フィルターで両端とも正確に分割）
- [x] 検証セクションに③月またぎ除外データ・④Agoda補正明細を追加（4セクション・4シートExcel）
- [x] 日別集計バグ修正：予約別計算を基本値とし、手動編集日のみ日次平均で再計算（2026-06-07）
- [x] Agoda 手数料補正 12%
- [x] 手動修正機能（`st.data_editor` + session state）
- [x] 全Excelテンプレ対応（領収書A〜D・月計表）
- [x] Mac 起動スクリプト刷新（Homebrew PATH 自動検出・ブラウザ自動オープン）
- [x] `setup.command` で実行権限・デスクトップアイコンを自動生成
- [x] Windows 配布 ZIP 対応（`.bat` を CP932+CRLF でエンコード）
- [x] ReservationTotalList 形式（75列・1行=1予約）のパース対応（`parse_hotel_total_csv`）（2026-06-13）
- [x] ZIP 3本を 2026-06-13 版で再ビルド済み

### 未着手・検討中

- [ ] Airbnb CSV の検証機能（現在は検証セクション非表示）
- [ ] 複数月を一括で申告書に出力する際のページ区切り確認

---

## 重要な注意点

1. **ZIP を更新するたびに3バリアントすべてを再ビルド**する（`python3 _build_zip.py`）。

2. **`settings.json` は配布 ZIP に含めない**。

3. **テンプレート Excel を変更する場合は `generate_receipt` のセルマッピングも確認**（`wb.active` で1枚目シート取得）。

4. **バリアント生成時は `ast.parse()` で構文チェック必須**。

5. **ホテルシステム CSV のエンコーディングは cp932**（`errors='replace'`）。日本語ヘッダーは文字化けするがデータは正常。

6. **課税標準額は必ず予約単位で計算**すること。日次合計の平均割りは同日に料金差がある場合に過小計算になる。

7. **`parse_hotel_csv_for_verification` は4値返却**。呼び出し側は `v_detail, v_sub, v_excluded, v_agoda = ...` と4変数でアンパックすること。

8. **宿泊日基準CSVの月またぎ**: 月をまたぐ予約は両月のCSVに全行が入る。`night_date` フィルターで除外しており二重計上・漏れはない。

9. **ねっぱんのCSVエクスポートは「ReservationList（日別明細）」を選ぶこと**。「ReservationTotalList（集計形式）」でも読み込み可能だが、検証機能が使えない。
