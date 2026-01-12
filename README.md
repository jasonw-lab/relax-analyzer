# RelaxAnalyzer

楽天カード等のクレジットカード明細 CSV を月別シートへ集約し、消費種類を自動分類する Excel VSTO アドインです。

## 概要

RelaxAnalyzer は、複数のクレジットカード明細 CSV ファイルを一括取り込みし、月別シートに整理・集約します。キーワードベースの自動タイプ分類により、家計管理を効率化します。

### 主な機能

- **CSV 一括取込**: 複数の明細 CSV を選択し、ファイル名から月を自動抽出して対応シートへ書き込み
- **消費種類自動分類**: 利用店名に対してキーワードマッチングで消費種類 (食費、保険、投資等) を自動設定
- **Amazon CSV サマリ作成**: Amazon注文履歴CSVから購入サマリを生成し、amazonシートへ自動貼付
- **Amazon 照合**: カード利用明細のAmazon利用に対し、日付・金額で照合して商品名を自動記入
- **高速処理**: COM 相互運用最適化により、数百行のデータを高速処理
- **柔軟な設定**: `type.csv` またはワークブック内 `type` シートでキーワード定義を管理

## スクリーンショット

### リボンUI
Excel のリボンに「RelaxAnalyzer」タブが追加されます:
- **CSV取込** ボタン: 複数の CSV ファイルを一括取り込み
- **消費種類** ボタン: アクティブシートの消費種類列を一括更新
- **Amazon CSV** ボタン: Amazon注文履歴CSVからサマリを生成
- **Amazon Check** ボタン: カード明細のAmazon利用に商品名を自動記入

## システム要件

- **OS**: Windows 10/11
- **Excel**: Microsoft Excel 2016 以降 (Office 2016, 2019, 365)
- **.NET Framework**: 4.7.2 以降
- **Visual Studio Tools for Office (VSTO) Runtime**: インストール時に自動配布

## インストール

1. [Releases](https://github.com/jasonw-lab/relax-analyzer/releases) から最新版のインストーラーをダウンロード
2. `setup.exe` を実行してインストール
3. Excel を起動し、リボンに「RelaxAnalyzer」タブが表示されることを確認

## 使い方

### 初回セットアップ

1. **config.ini の配置** (オプション)
   - アドインの実行ファイルと同じディレクトリに `config.ini` を配置
   - 例:
     ```ini
     Project = E:\Project
     TypeCSV = relax-analyzer\rule\type.csv
     ```

2. **type.csv または type シートの準備**
   - ワークブック内に `type` シートを作成 (推奨)、または外部 `type.csv` を用意
   - A列: keyword (利用店名に含まれるキーワード)
   - B列: type (消費種類)
   - C列: comment (任意のメモ)

**例 (type シート)**:
   | keyword       | type | comment |
   |---------------|------|---------|
   | 楽天証券      | 投資 |         |
   | 朝日生命      | 保険 |         |
   | セブンイレブン | 食費 |         |
   | マクドナルド   | 食費 |         |

### CSV 取込

1. Excel で新規ブックを開く、または既存のブックを開く
2. リボンの **RelaxAnalyzer** タブ → **CSV取込** をクリック
3. 明細 CSV ファイルを複数選択 (Ctrl+クリックで複数選択)
4. 処理完了後、月別シート (1〜12) にデータが追記されます

**ファイル名形式**:
- `enaviYYMMDD(XXXX).csv` (例: `enavi250315(1234).csv` → 3月シート)
- `enaviYYYYMM(XXXX).csv` (例: `enavi202503(1234).csv` → 3月シート)

### 消費種類更新

CSV 取込後、または既存データに対して消費種類を一括更新:

1. 更新したいシートをアクティブにする (例: 「3」シート)
2. リボンの **RelaxAnalyzer** タブ → **消費種類** をクリック
3. K列 (消費種類) が自動更新されます

**動作**:
- B列 (利用店名・商品名) を読み取り、`type` シートまたは `type.csv` のキーワードと照合
- 最初に一致したキーワードの type を K列に設定

### Amazon CSV サマリ作成

Amazon注文履歴CSVから購入サマリを生成し、カード明細との照合に使用:

1. リボンの **RelaxAnalyzer** タブ → **Amazon CSV** をクリック
2. Amazon注文履歴CSV (`Retail.OrderHistory*.csv`) を選択
3. 出力先CSVファイル名を指定 (デフォルト: `amazon_order_summary.csv`)
4. 処理完了後、ワークブックに `amazon` シートが作成され、サマリデータが自動貼付されます

**Amazon注文履歴CSVの取得方法**:
1. Amazon.co.jp にログイン
2. アカウント＆リスト → 注文履歴レポート
3. 期間を指定して「レポートをリクエスト」
4. ダウンロードした `Retail.OrderHistory*.csv` を使用

**出力形式 (amazonシート)**:
| Order Date | Order ID | Item Short Name | 金額 | Order Status | Item Name | Short Name | Quantity |
|------------|----------|-----------------|------|--------------|-----------|------------|----------|
| 2025-11-17 | 503-xxx  | サランラップ... | 1490 | Authorized   | サラン... | サランラ... | 1        |

### Amazon 照合

カード利用明細のAmazon利用に対し、日付・金額で照合して商品名を自動記入:

1. 事前に **Amazon CSV** ボタンで `amazon` シートを作成
2. カード明細の月シート (1〜12) をアクティブにする
3. リボンの **RelaxAnalyzer** タブ → **Amazon Check** をクリック
4. L列 (コメント) にAmazon商品名が自動記入されます

**照合条件**:
- B列 (利用店名・商品名) に「AMAZON.」が含まれる行を対象
- **L列（コメント欄）が空の行のみ処理（既に入力済みの場合はスキップ）**
- A列 (利用日) の前後1週間以内のAmazon注文を検索
- **E列 (利用金額) と Amazon金額が一致するデータのみ抽出**
- 該当する商品名 (Item Short Name) をL列に記入
- 複数該当する場合は改行区切りで記入

**対象シート**:
- アクティブシートが月シート (1〜12) の場合: そのシートのみ処理
- それ以外のシートがアクティブの場合: 全月シート (1〜12) を処理するか確認ダイアログを表示

## データ形式

### 入力 CSV (楽天カード形式想定)

12列のカンマ区切り形式:
```csv
ご利用日,ご利用店名・商品名,利用者,支払方法,ご利用金額,手数料,支払総額,当月請求額,翌月繰越残高,備考,消費種類,メモ
2025/03/15,セブンイレブン,本人,1回払い,500,0,500,500,0,,食費,
```

### 出力シート形式

月別シート (1〜12) の 4 行目以降にデータが追記:
- A列: ご利用日
- B列: ご利用店名・商品名
- C〜E列: その他明細データ (E列: 利用金額)
- F〜J列: その他明細データ
- K列: 消費種類 (自動設定)
- L列: コメント (Amazon Check で自動記入)

**ファイル名行**:
- 背景色 `#E6F3FF` で各 CSV の開始位置を視覚的に区別

## 技術仕様

### 開発環境・技術スタック

- **言語**: C# 7.3
- **フレームワーク**: .NET Framework 4.7.2
- **プラットフォーム**: VSTO (Visual Studio Tools for Office)
- **ビルドツール**: MSBuild / Visual Studio 2022

### 主要ライブラリ

| ライブラリ | バージョン | 用途 |
|-----------|-----------|------|
| CsvHelper | 30.0.1 | CSV パース・読込 |
| Microsoft.Office.Interop.Excel | 15.0+ | Excel COM相互運用 |
| Microsoft.Office.Tools.Excel | 10.0+ | VSTO ランタイム |

### アーキテクチャ

```
analyzer/
├── Ribbon1.cs           # リボン UI イベントハンドラ
├── ThisAddIn.cs     # アドインエントリポイント・状態管理
└── Core/
    ├── RelaxAnalyzerConfig.cs      # config.ini 読込
    ├── TypeKeyword.cs          # キーワード・type ペアモデル
    ├── TypeMappingProvider.cs      # type シート/CSV 読込
    ├── TypeResolver.cs          # キーワード→type 解決ロジック
    ├── MonthExtractor.cs     # ファイル名→月抽出 (Regex)
    ├── CsvImportModels.cs          # データモデル (Batch, Chunk, Row)
    ├── CsvImportService.cs     # CSV 読込・正規化 (非同期)
    ├── SheetWriter.cs            # Excel シート書き込み (一括操作)
    ├── AmazonOrderSummaryService.cs # Amazon注文履歴CSV処理
    └── AmazonCheckService.cs       # Amazon照合・商品名記入
```

### 性能最適化

1. **COM 相互運用最適化**
   - 一括範囲操作 (`Range.Value2 = object[,]`) で個別セルアクセスを削減
   - 50行超のデータで画面更新・イベント・再計算を一時停止

2. **非同期処理**
   - CSV I/O は `Task.Run` でバックグラウンド実行
   - UI スレッドブロックを回避

3. **I/O 最適化**
   - FileStream バッファサイズ 32KB
   - SequentialScan オプション

4. **日付処理の最適化**
   - Excelシリアル値 (OLE Automation date) に対応
   - `DateTime.FromOADate()` による高速変換

**実測性能**:
- 消費種類更新: 数百行で **10倍以上高速化** (従来比)
- CSV 読込: 大容量ファイルで **10-20%高速化**
- Amazon照合: 一括範囲操作により高速処理

### 文字コード対応

- **UTF-8** (BOM あり/なし): 優先的に試行
- **Shift_JIS** (CP932): UTF-8 失敗時フォールバック
- 自動検出・警告表示

### 日付フォーマット対応

Amazon Check機能は以下の日付形式に対応:
- **Excelシリアル値** (例: `45978` → `2025-11-17`)
- `yyyy-MM-dd` (例: `2025-11-17`)
- `yyyy/MM/dd` (例: `2025/11/17`)
- `yyyy/M/d` (例: `2025/9/16` - 1桁の月・日)

## トラブルシューティング

### リボンタブが表示されない

- Excel を完全に終了し、再起動
- ファイル → オプション → アドイン → 「管理」で「COMアドイン」を選択 → 「RelaxAnalyzer」が有効か確認

### CSV 取込エラー

- ファイル名が `enaviYYMMDD` または `enaviYYYYMM` 形式か確認
- CSV がカンマ区切り形式か確認 (タブ区切りは非対応)
- 文字コードが UTF-8 または Shift_JIS か確認

### 消費種類が更新されない

- `type` シートまたは `config.ini` で指定した `type.csv` が存在するか確認
- キーワードが B列 (利用店名) に部分一致で含まれているか確認 (大文字小文字無視)

### Amazon Check が動作しない

- **「'amazon' シートが見つかりません」エラー**
  - 先に **Amazon CSV** ボタンで `amazon` シートを作成してください
  
- **「'amazon' シートにデータがありません」エラー**
  - Amazon注文履歴CSVの日付がExcelシリアル値または文字列として正しく読み込まれているか確認
  - デバッグメッセージで詳細を確認してください

- **商品名が記入されない**
  - カード明細のB列に「AMAZON.」が含まれているか確認（大文字小文字無視）
  - カード明細のE列（利用金額）とAmazon金額が一致しているか確認
  - 利用日とAmazon注文日が前後1週間以内か確認

## 変更履歴

### v1.2.0 (2025-XX-XX)
- ✨ Amazon CSV サマリ作成機能追加
- ✨ Amazon 照合機能追加（日付・金額で自動照合）
- 🔧 Excelシリアル値対応により日付処理を改善
- 📝 Amazon注文履歴との連携機能を実装

### v1.1.0 (2025-XX-XX)
- ✨ 消費種類一括更新ボタン追加
- ⚡ COM相互運用最適化 (10倍高速化)
- 🚀 CSV読込バッファサイズ拡大

### v1.0.0 (2025-XX-XX)
- 🎉 初回リリース
- CSV 一括取込機能
- キーワードベース自動分類

## サポート

- **Issues**: [GitHub Issues](https://github.com/jasonw-lab/relax-analyzer/issues)
- **Discussions**: [GitHub Discussions](https://github.com/jasonw-lab/relax-analyzer/discussions)

---

**開発者**: jasonw-lab  
**リポジトリ**: https://github.com/jasonw-lab/relax-analyzer
