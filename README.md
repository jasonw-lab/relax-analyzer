# RelaxAnalyzer

楽天カード等のクレジットカード明細 CSV を月別シートへ集約し、消費種類を自動分類する Excel VSTO アドインです。

## 概要

RelaxAnalyzer は、複数のクレジットカード明細 CSV ファイルを一括取り込みし、月別シートに整理・集約します。キーワードベースの自動タイプ分類により、家計管理を効率化します。

### 主な機能

- **CSV 一括取込**: 複数の明細 CSV を選択し、ファイル名から月を自動抽出して対応シートへ書き込み
- **消費種類自動分類**: 利用店名に対してキーワードマッチングで消費種類 (食費、保険、投資等) を自動設定
- **高速処理**: COM 相互運用最適化により、数百行のデータを高速処理
- **柔軟な設定**: `type.csv` またはワークブック内 `type` シートでキーワード定義を管理

## スクリーンショット

### リボンUI
Excel のリボンに「RelaxAnalyzer」タブが追加されます:
- **CSV取込** ボタン: 複数の CSV ファイルを一括取り込み
- **消費種類** ボタン: アクティブシートの消費種類列を一括更新

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
     TypeCSV = D:\Git\ross-dev2024\relax-analyzer\rule\type.csv
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
- C〜J列: その他明細データ
- K列: 消費種類 (自動設定)
- L列: メモ

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
├── Ribbon1.cs                   # リボン UI イベントハンドラ
├── ThisAddIn.cs                 # アドインエントリポイント・状態管理
└── Core/
    ├── RelaxAnalyzerConfig.cs   # config.ini 読込
    ├── TypeKeyword.cs           # キーワード・type ペアモデル
    ├── TypeMappingProvider.cs   # type シート/CSV 読込
    ├── TypeResolver.cs          # キーワード→type 解決ロジック
    ├── MonthExtractor.cs        # ファイル名→月抽出 (Regex)
    ├── CsvImportModels.cs       # データモデル (Batch, Chunk, Row)
    ├── CsvImportService.cs      # CSV 読込・正規化 (非同期)
    └── SheetWriter.cs           # Excel シート書き込み (一括操作)
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

**実測性能**:
- 消費種類更新: 数百行で **10倍以上高速化** (従来比)
- CSV 読込: 大容量ファイルで **10-20%高速化**

### 文字コード対応

- **UTF-8** (BOM あり/なし): 優先的に試行
- **Shift_JIS** (CP932): UTF-8 失敗時フォールバック
- 自動検出・警告表示

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


## 変更履歴

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
