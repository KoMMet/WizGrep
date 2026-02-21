# WizGrep 設計書

## 1. 概要

WizGrepは、Windows向けの高機能ファイル検索（Grep）アプリケーションです。テキストファイルだけでなく、Excel、Word、PowerPoint、PDFなどのOffice系ファイルも含めて横断的にキーワード検索を行うことができます。OOXML形式（.xlsx/.docx/.pptx等）だけでなく、旧バイナリ形式（.xls/.doc/.ppt）にも対応しています。

### 1.1 技術スタック

| 項目 | 技術 |
|------|------|
| フレームワーク | .NET 10 / WinUI 3 |
| 言語 | C# 14.0 |
| アーキテクチャ | MVVM (Model-View-ViewModel) |
| UIフレームワーク | Windows App SDK |
| MVVMライブラリ | CommunityToolkit.Mvvm（ObservableObject, ObservableProperty, RelayCommand） |
| OOXML処理 | DocumentFormat.OpenXml（.docx/.docm, .xlsx/.xlsm, .pptx/.pptm） |
| 旧Excel処理 | NPOI / HSSFWorkbook（.xls） |
| 旧Word/PPT処理 | OpenMcdf（.doc, .ppt） |
| PDF処理 | UglyToad.PdfPig（.pdf） |
| 文字コード処理 | System.Text.Encoding + CodePagesEncodingProvider |

### 1.2 システム要件

- Windows 10 バージョン 1809 以降
- Windows 11
- x64 / ARM64 アーキテクチャ

---

## 2. 機能一覧

| # | 機能名 | 説明 |
|---|--------|------|
| 1 | マルチキーワード検索 | 最大5つのキーワードでAND/OR検索 |
| 2 | マルチフォーマット対応 | Excel, Word, PowerPoint, PDF, テキストファイルに対応（OOXML＋旧バイナリ形式） |
| 3 | 正規表現検索 | 正規表現パターンによる高度な検索（事前バリデーション付き） |
| 4 | 大文字小文字区別 | 検索時の大文字小文字の区別設定 |
| 5 | リアルタイム表示 | 検索結果をリアルタイムで画面に表示 |
| 6 | インデックス機能 | 検索高速化のためのインデックスファイル生成・差分更新 |
| 7 | 検索結果エクスポート | 結果をTSV形式テキストファイルとして出力（タブ・改行エスケープ付き） |
| 8 | 設定永続化 | アプリケーション設定のローカル保存（JSON形式） |
| 9 | 除外拡張子 | 特定の拡張子をスキップする除外フィルタ |
| 10 | カスタム拡張子 | 標準以外の拡張子を追加して検索対象にする |
| 11 | 文字コード自動判定 | BOM検出 + UTF-8/Shift-JIS のstrict判定による自動検出 |

---

## 3. 処理概要

### 3.1 アプリケーション起動フロー

```
App.OnLaunched()
    └── MainWindow()
        ├── サービス初期化
        │   ├── SettingsService          … 設定の読み書き
        │   ├── FileReaderService        … ファイルリーダーの登録・管理
        │   │   ├── TextFileReader       … .txt
        │   │   ├── ExcelFileReader      … .xlsx, .xlsm
        │   │   ├── NpoiExcelFileReader  … .xls
        │   │   ├── WordFileReader       … .docx, .docm
        │   │   ├── DocFileReader        … .doc
        │   │   ├── PowerPointFileReader … .pptx, .pptm
        │   │   ├── PptFileReader        … .ppt
        │   │   └── PdfFileReader        … .pdf
        │   ├── IndexService             … インデックスの保存・読み込み・削除
        │   └── GrepService(FileReaderService, IndexService)
        ├── ViewModel初期化
        │   └── MainViewModel(GrepService, SettingsService)
        │       ├── 設定読み込み（GrepSettings, WizGrepSettings）
        │       └── ObservableCollection初期化
        └── InitializeComponent()
```

### 3.2 検索処理フロー

```
ユーザー操作: Grep設定ダイアログで「開始」クリック
    │
    ├── GrepSettingsDialogViewModel.Validate()
    │   ├── フォルダパスの空チェック
    │   ├── ファイル種別の選択チェック（HasAnyTargetFileSelected）
    │   └── 有効なキーワードの存在チェック
    │
    ▼
MainViewModel.StartSearchAsync()  ←  UIスレッド
    │
    ├── 設定をローカル変数にキャプチャ（スレッド安全のため）
    │   ├── grepSettings = GrepSettings
    │   ├── wizGrepSettings = WizGrepSettings
    │   ├── realTimeDisplay = grepSettings.RealTimeDisplay
    │   └── token = _cancellationTokenSource.Token
    │
    ├── Progress<GrepProgress> を UIスレッドで生成
    │
    └── Task.Run() ← バックグラウンドスレッド
        │
        └── GrepService.ExecuteGrepAsync()
            │
            ├── キーワード取得・空チェック
            │
            ├── 正規表現モード時：パターン事前検証
            │   └── 無効パターンは InvalidOperationException をスロー
            │
            ├── RebuildIndex時：既存インデックス削除
            │
            ├── 対象ファイル列挙
            │   ├── GetTargetExtensions()（ALLの場合は空リスト＝全ファイル対象）
            │   ├── GetExcludeExtensions()
            │   └── 拡張子が空でALLも未選択 → 即座に空結果を返却
            │
            ├── インデックス存在チェック
            │
            ├── ファイル処理ループ（CancellationToken対応）
            │   ├── 進捗報告 → Progress<T>経由でUIスレッドに通知
            │   ├── タイムスタンプ比較（差1秒未満で一致と判定）
            │   │   ├── 一致 → インデックスから内容取得
            │   │   └── 不一致 → FileReaderServiceで再読み込み
            │   ├── キーワードマッチング
            │   │   ├── AND検索: keywords.All(k => MatchesKeyword(...))
            │   │   └── OR検索:  keywords.Any(k => MatchesKeyword(...))
            │   ├── Excel改行削除（RemoveExcelLineBreaks + IsExcelFile判定）
            │   └── リアルタイム表示時は結果をProgressで即時報告
            │
            ├── インデックス保存（全ファイル内容 + タイムスタンプ）
            │
            └── 結果リストを返却

    ├── 非リアルタイム時：返却結果を一括でUIに追加
    │
    └── ステータスメッセージ更新
```

---

## 4. アーキテクチャ

### 4.1 レイヤー構成

```
┌──────────────────────────────────────────────────┐
│                     View                          │
│  MainWindow.xaml / .cs                            │
│  GrepSettingsDialog.xaml / .cs                    │
│  WizGrepSettingsDialog.xaml / .cs                 │
├──────────────────────────────────────────────────┤
│                  ViewModel                        │
│  MainViewModel                                    │
│  GrepSettingsDialogViewModel                      │
│  WizGrepSettingsDialogViewModel                   │
├──────────────────────────────────────────────────┤
│                   Services                        │
│  GrepService          … 検索オーケストレーション    │
│  FileReaderService    … ファイルリーダー管理        │
│  IndexService         … インデックス永続化          │
│  SettingsService      … 設定永続化                 │
├──────────────────────────────────────────────────┤
│                 FileReaders                       │
│  IFileReader (インターフェース)                     │
│  TextFileReader, ExcelFileReader,                  │
│  NpoiExcelFileReader, WordFileReader,              │
│  DocFileReader, PowerPointFileReader,              │
│  PptFileReader, PdfFileReader                      │
├──────────────────────────────────────────────────┤
│                   Models                          │
│  GrepResult, GrepSettings, WizGrepSettings,       │
│  SearchKeyword, FileTimestamp                      │
├──────────────────────────────────────────────────┤
│                   Helpers                         │
│  EncodingDetector                                 │
└──────────────────────────────────────────────────┘
```

### 4.2 プロジェクト構造

```
WizGrep/
├── App.xaml / App.xaml.cs              # アプリケーションエントリポイント
├── MainWindow.xaml / .cs               # メイン画面
├── Models/
│   ├── GrepResult.cs                   # 検索結果モデル（Location表示・インデックス変換）
│   ├── GrepSettings.cs                 # Grep設定モデル（拡張子管理・キーワード取得）
│   ├── WizGrepSettings.cs              # WizGrep設定モデル（インデックス設定）
│   ├── SearchKeyword.cs                # キーワード設定モデル（Keyword + IsEnabled）
│   └── FileTimestamp.cs                # タイムスタンプモデル（パイプ区切りシリアライズ）
├── ViewModels/
│   ├── MainViewModel.cs                # メイン画面ViewModel
│   ├── GrepSettingsDialogViewModel.cs  # Grep設定ダイアログViewModel
│   └── WizGrepSettingsDialogViewModel.cs # WizGrep設定ダイアログViewModel
├── Views/
│   ├── GrepSettingsDialog.xaml / .cs    # Grep設定ダイアログ
│   └── WizGrepSettingsDialog.xaml / .cs # WizGrep設定ダイアログ
├── Services/
│   ├── GrepService.cs                  # Grep検索サービス（GrepProgress進捗報告）
│   ├── FileReaderService.cs            # ファイル読み込み管理（Strategy pattern）
│   ├── IndexService.cs                 # インデックス管理（保存・読み込み・削除）
│   ├── SettingsService.cs              # 設定保存/読み込み（ApplicationData.LocalSettings）
│   └── FileReaders/
│       ├── IFileReader.cs              # ファイルリーダーインターフェース
│       ├── TextFileReader.cs           # テキストファイルリーダー（EncodingDetector連携）
│       ├── ExcelFileReader.cs          # OOXML Excelリーダー（セル・図形・コメント）
│       ├── NpoiExcelFileReader.cs      # 旧Excel（.xls）リーダー（NPOI/HSSF）
│       ├── WordFileReader.cs           # OOXML Wordリーダー（段落・図形・ヘッダー・フッター）
│       ├── DocFileReader.cs            # 旧Word（.doc）リーダー（OpenMcdf/FIB/CLX）
│       ├── PowerPointFileReader.cs     # OOXML PPTリーダー（スライド・図形・表・ノート）
│       ├── PptFileReader.cs            # 旧PPT（.ppt）リーダー（OpenMcdf/レコードツリー）
│       └── PdfFileReader.cs            # PDFリーダー（PdfPig/ページ単位テキスト抽出）
├── Helpers/
│   └── EncodingDetector.cs             # 文字コード自動判定（BOM + strict decode）
└── doc/
    ├── design.md                       # 設計書（本ファイル）
    └── manual.md                       # ユーザーマニュアル
```

---

## 5. 機能詳細仕様

### 5.1 マルチキーワード検索

#### 5.1.1 キーワード設定

| 項目 | 仕様 |
|------|------|
| 最大キーワード数 | 5個 |
| 個別有効/無効 | チェックボックスで切り替え |
| 空キーワード | 検索対象から除外（`GetActiveKeywords()` で `IsEnabled && !IsNullOrWhiteSpace` をフィルタ） |

#### 5.1.2 検索条件

| 条件 | 動作 |
|------|------|
| AND検索（デフォルト） | すべての有効キーワードが含まれる行/セルをマッチ |
| OR検索 | いずれかの有効キーワードが含まれる行/セルをマッチ |

#### 5.1.3 マッチング処理

```csharp
// AND検索
keywords.All(k => MatchesKeyword(content, k, settings))

// OR検索
keywords.Any(k => MatchesKeyword(content, k, settings))

// 各キーワードのマッチ判定
private bool MatchesKeyword(string content, string keyword, GrepSettings settings)
{
    if (settings.UseRegex)
    {
        var options = settings.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        return Regex.IsMatch(content, keyword, options);
    }
    var comparison = settings.CaseSensitive
        ? StringComparison.Ordinal
        : StringComparison.OrdinalIgnoreCase;
    return content.Contains(keyword, comparison);
}
```

#### 5.1.4 正規表現パターンの事前検証

正規表現モードの場合、検索ループに入る前に全キーワードのパターンを検証します。
無効なパターンが含まれている場合、`InvalidOperationException` をスローし、ステータスバーにエラーメッセージを表示します。

```csharp
// 事前検証
foreach (var keyword in keywords)
    Regex.Match("", keyword, regexOptions); // ArgumentException → InvalidOperationException
```

#### 5.1.5 バリデーション

Grep設定ダイアログの「開始」ボタン押下時に以下を検証します。
すべて満たさない場合、ダイアログは閉じずキャンセルされます。

| 検証項目 | 条件 |
|---------|------|
| 検索対象フォルダ | 空でないこと |
| 対象ファイル | 少なくとも1つのファイル種別が選択されていること |
| キーワード | 有効かつ空でないキーワードが少なくとも1つあること |

### 5.2 対応ファイル形式

#### 5.2.1 ファイルリーダー対応表

| リーダー | 対応拡張子 | ライブラリ | 読み取り対象 |
|----------|-----------|-----------|-------------|
| TextFileReader | .txt | 標準API + EncodingDetector | 行単位テキスト |
| ExcelFileReader | .xlsx, .xlsm | DocumentFormat.OpenXml | セル値・図形テキスト・コメント |
| NpoiExcelFileReader | .xls | NPOI (HSSFWorkbook) | セル値・図形テキスト・コメント（グループ図形再帰対応） |
| WordFileReader | .docx, .docm | DocumentFormat.OpenXml | 段落テキスト・VML図形テキスト・ヘッダー・フッター |
| DocFileReader | .doc | OpenMcdf | FIB/CLXピーステーブルからテキスト復元（Windows-1252/UTF-16LE対応）・フィールドコード除去 |
| PowerPointFileReader | .pptx, .pptm | DocumentFormat.OpenXml | スライドテキスト・図形テキスト・表/グラフテキスト・ノート |
| PptFileReader | .ppt | OpenMcdf | レコードツリーからテキスト抽出（TextCharsAtom/TextBytesAtom）・スライド/ノート区別 |
| PdfFileReader | .pdf | UglyToad.PdfPig | ページ単位テキスト抽出・行分割 |

#### 5.2.2 ファイルリーダーの選択ロジック（FileReaderService）

```
拡張子 → 登録済みリーダーを検索
    ├── 一致あり → 対応リーダーで読み込み
    └── 一致なし → TextFileReader（デフォルト）で読み込み
```

登録されたリーダーはStrategy patternで管理され、拡張子をキーとするDictionaryで解決されます。

#### 5.2.3 各ファイルリーダーの読み取り詳細

**TextFileReader:**
- `EncodingDetector.DetectEncoding()` でファイルの文字コードを自動判定
- `File.ReadAllLines()` で全行を読み込み
- 各行に行番号（1始まり）を付与

**ExcelFileReader (.xlsx/.xlsm):**
- `SpreadsheetDocument.Open()` で読み取り専用オープン
- SharedStringTableを利用してセル値を解決
- 全シート × 全行 × 全セルを走査
- 図形テキスト: `DrawingsPart` → `Shape` → `Paragraph` を走査
- コメント: `WorksheetCommentsPart` → `Comment` を走査

**NpoiExcelFileReader (.xls):**
- `HSSFWorkbook` で旧形式Excelを読み取り
- セル値: CellType に応じて文字列化（String/Numeric/Boolean/Formula/日付判定）
- 図形テキスト: `HSSFPatriarch` からグループ化された図形を再帰的に探索
- コメント: `ICell.CellComment` から取得

**WordFileReader (.docx/.docm):**
- `WordprocessingDocument.Open()` で読み取り専用オープン
- 本文段落: VML図形内の段落を除外して処理（重複防止）
- VML図形を含むRunも除外してテキスト取得
- 図形テキスト: `body.Descendants<Shape>()` で別途取得
- ヘッダー・フッター: `HeaderParts` / `FooterParts` から取得

**DocFileReader (.doc):**
- `RootStorage.OpenRead()` でCompound File Binaryを読み取り
- マジックナンバー（0xA5EC）の検証
- FIBフィールドからフラグ・テキスト長・CLXオフセットを取得
- テーブルストリーム（0Table/1Table）からCLX構造を解析
- ピーステーブル（PlcPcd）からテキストを復元（圧縮時Windows-1252/非圧縮時UTF-16LE）
- フィールドコード（\x13～\x15）をスタックベースで除去（ネスト対応）
- 制御文字を除去し、テーブルセル区切り（\x07）をタブに変換

**PowerPointFileReader (.pptx/.pptm):**
- `PresentationDocument.Open()` で読み取り専用オープン
- スライドごとにShape内のテキストを取得（shapeName or 「図形N」）
- GraphicFrame内のテキスト（表・グラフ）を取得
- ノートスライドのテキストを取得

**PptFileReader (.ppt):**
- `RootStorage.OpenRead()` → "PowerPoint Document" ストリームを読み取り
- レコードツリーを再帰的にスキャン（recVer=0x0F: コンテナ / それ以外: アトム）
- SlideListWithTextContainer はスキップ（SlideContainer内と重複するため）
- TextCharsAtom（UTF-16LE）/ TextBytesAtom（Windows-1252）からテキスト抽出
- SlideContainer/NotesContainer でスライド番号・ノート区別を管理

**PdfFileReader:**
- `PdfDocument.Open()` で読み込み
- ページごとにテキストを抽出し、改行で行分割
- 各行に行番号（1始まり）、SheetNameにページ番号（「ページN」）を設定

#### 5.2.4 カスタム拡張子

- チェックボックスで使用の有効/無効を切り替え
- チェックがOFFの場合、入力欄は無効化され拡張子は無視される
- ユーザーが任意の拡張子を追加可能
- カンマ、セミコロン、スペースで区切り
- ドットが省略された場合は自動付与（例: `csv` → `.csv`）
- 既に対象リストに含まれる拡張子は重複追加されない
- 追加された拡張子はFileReaderServiceでリーダーが見つからない場合TextFileReaderで処理

#### 5.2.5 除外拡張子

- チェックボックスで使用の有効/無効を切り替え
- カンマ、セミコロン、スペースで区切り
- ドットが省略された場合は自動付与
- 重複は除去される（`Distinct()`）
- ファイル列挙時に除外拡張子に一致するファイルをスキップ
- 除外は対象拡張子フィルタよりも優先される

### 5.3 検索オプション

#### 5.3.1 大文字小文字区別

| 設定 | 動作 |
|------|------|
| OFF（デフォルト） | 大文字小文字を区別しない（`StringComparison.OrdinalIgnoreCase` / `RegexOptions.IgnoreCase`） |
| ON | 大文字小文字を区別する（`StringComparison.Ordinal` / `RegexOptions.None`） |

#### 5.3.2 正規表現

| 設定 | 動作 |
|------|------|
| OFF（デフォルト） | `string.Contains()` による通常文字列検索 |
| ON | `Regex.IsMatch()` による正規表現パターンマッチ |

正規表現使用時:
- 大文字小文字区別設定に応じて `RegexOptions.IgnoreCase` を付与
- 検索開始前にパターンの事前検証を実施（無効パターンはエラー表示）

#### 5.3.3 リアルタイム表示

| 設定 | 動作 |
|------|------|
| ON（デフォルト） | `Progress<GrepProgress>` 経由で結果が見つかるたびにUIスレッドに通知・追加 |
| OFF | 検索完了後に結果リストを一括でUIに追加 |

検索処理は `Task.Run()` でバックグラウンドスレッドで実行されます。
`Progress<T>` はUIスレッドで生成されるため、コールバックは自動的にUIスレッドにマーシャリングされます。

#### 5.3.4 Excel改行削除

| 設定 | 動作 |
|------|------|
| OFF（デフォルト） | セル内改行をそのまま表示 |
| ON | 対象ファイルがExcel（.xlsx/.xlsm/.xls）の場合のみ、`\r\n`, `\n`, `\r` をスペースに置換 |

- 改行削除は表示用のContentに対してのみ適用されます
- インデックスには元のContent（改行を含む）が保存されます
- キーワードマッチングは元のContent（改行を含む）に対して行われます

### 5.4 インデックス機能

#### 5.4.1 インデックスの概要

インデックスは、ファイルの読み取り結果をキャッシュして次回検索を高速化する機能です。
インデックスにはキーワードに関係なく全ファイルの全内容が保存されるため、異なるキーワードでの再検索時にもファイル読み取りを省略できます。

#### 5.4.2 インデックスファイル構成

```
[IndexBasePath]/
└── [ドライブ文字（コロン除去）]/
    └── [フォルダパス]/
        ├── GrepIndex.txt      # 全ファイルの内容データ
        └── GrepTimestamp.txt  # ファイルのタイムスタンプ
```

例: ベースパス `C:\WizGrepIndex`、対象フォルダ `D:\Documents\Project`
→ `C:\WizGrepIndex\D\Documents\Project\GrepIndex.txt`

#### 5.4.3 インデックスファイル形式

**GrepIndex.txt（タブ区切り、UTF-8）:**
```
FilePath\tLineNumber\tSheetName\tCellAddress\tObjectName\tContent(エスケープ済み)
```

Content内のエスケープ規則:
| 元の文字 | エスケープ後 |
|---------|------------|
| `\` | `\\` |
| `\r` | `\r`（リテラル文字列） |
| `\n` | `\n`（リテラル文字列） |
| `\t` | `\t`（リテラル文字列） |

復元時はエスケープシーケンスを逆変換します。

**GrepTimestamp.txt（パイプ区切り、UTF-8）:**
```
FilePath|LastModifiedDateTime(ISO 8601 ラウンドトリップ形式)
```

例: `C:\Data\report.xlsx|2024-01-15T10:30:45.1234567+09:00`

#### 5.4.4 インデックス利用判定

```
IF インデックスベースパスが設定済み AND インデックスファイルが存在 THEN
    各ファイルについて:
        IF ファイルのLastWriteTimeとインデックスのタイムスタンプの差が1秒未満 THEN
            インデックスからファイルの内容を取得
        ELSE
            ファイルを再読み込み
        END IF
ELSE
    ファイルを読み込み
END IF

検索完了後:
    IF インデックスベースパスが設定済み THEN
        全ファイル内容とタイムスタンプをインデックスに保存
    END IF
```

#### 5.4.5 インデックス再構築

RebuildIndexフラグがONの場合、検索開始時に既存インデックスを削除してから処理を開始します。

### 5.5 検索結果

#### 5.5.1 GrepResult構造

| プロパティ | 型 | 説明 |
|-----------|-----|------|
| FilePath | string | ファイルの絶対パス |
| FileName | string | ファイル名のみ（読み取り専用、`Path.GetFileName()` で取得） |
| LineNumber | int | 行番号（テキスト/Word/PDF）またはExcelの行インデックス |
| SheetName | string? | シート名（Excel）、スライド名（PowerPoint）、ページ番号（PDF） |
| CellAddress | string? | セルアドレス（Excel、例: A1）、コメントの参照セル |
| ObjectName | string? | オブジェクト名（図形、コメント、ヘッダー、フッター、ノート、表/グラフなど） |
| Content | string | マッチした行/セルの内容 |
| Location | string | 表示用の位置情報（読み取り専用） |

#### 5.5.2 Location表示形式

| 条件 | 表示形式 | 例 |
|------|---------|-----|
| SheetName + CellAddress + ObjectName | `[シート名] セルアドレス オブジェクト名` | `[Sheet1] A1 コメント` |
| SheetName + CellAddress | `[シート名] セルアドレス` | `[Sheet1] A1` |
| SheetName + ObjectName | `[シート名] オブジェクト名` | `[スライド1] 図形1` |
| SheetName のみ | `[シート名] 行番号 行` | `[Sheet1] 5 行` |
| ObjectName のみ | `[オブジェクト] オブジェクト名` | `[オブジェクト] ヘッダー1` |
| いずれもなし | `行番号 行` | `15 行` |

#### 5.5.3 各ファイル種別のGrepResult設定

| ファイル種別 | SheetName | CellAddress | ObjectName | LineNumber |
|-------------|-----------|-------------|------------|------------|
| テキスト | (なし) | (なし) | (なし) | 行番号 |
| Excel セル | シート名 | セルアドレス | (なし) | 行インデックス |
| Excel 図形 | シート名 | (なし) | `図形N` | 0 |
| Excel コメント | シート名 | 参照セル | `コメント` | 0 |
| Word 段落 | (なし) | (なし) | (なし) | 段落番号 |
| Word 図形 | (なし) | (なし) | `図形N` | 0 |
| Word ヘッダー | (なし) | (なし) | `ヘッダーN` | 0 |
| Word フッター | (なし) | (なし) | `フッターN` | 0 |
| PowerPoint テキスト | `スライドN` | (なし) | 図形名 or `図形N` | 行番号 |
| PowerPoint 表/グラフ | `スライドN` | (なし) | `表/グラフ` | テキスト番号 |
| PowerPoint ノート | `スライドN` | (なし) | `ノート` | 行番号 |
| PDF | `ページN` | (なし) | (なし) | 行番号 |
| .doc (旧Word) | (なし) | (なし) | (なし) | 行番号 |
| .ppt (旧PPT) | `スライドN` or `文書情報` | (なし) | `ノート` or (なし) | 行番号 |
| .xls (旧Excel) セル | シート名 | セルアドレス | (なし) | 行インデックス |
| .xls 図形 | シート名 | (なし) | `図形N` | 0 |
| .xls コメント | シート名 | セルアドレス | `コメント` | 0 |

### 5.6 エクスポート機能

#### 5.6.1 ファイル一覧エクスポート

- 左ペインの「出力」ボタンで実行
- 1行に1ファイルパスを出力
- `FileSavePicker` でファイル保存先を選択

#### 5.6.2 検索結果エクスポート

- 右ペインの「出力」ボタンで実行
- TSV（タブ区切り）形式で出力
- ヘッダー行: `ファイルパス\tファイル名\t位置\t内容`
- 位置（Location）と内容（Content）内のタブ・改行はスペースに置換してエクスポート（TSV破損防止）
- `FileSavePicker` でファイル保存先を選択

### 5.7 設定永続化

#### 5.7.1 保存先

`Windows.Storage.ApplicationData.Current.LocalSettings`

#### 5.7.2 保存形式

JSON（`System.Text.Json.JsonSerializer` でシリアライズ/デシリアライズ）

#### 5.7.3 保存キー

| キー | 内容 |
|------|------|
| `GrepSettings` | Grep設定（キーワード、対象ファイル、オプション、カスタム/除外拡張子） |
| `WizGrepSettings` | WizGrep設定（インデックスベースパス、再構築フラグ） |

#### 5.7.4 保存タイミング

- GrepSettings: Grep設定ダイアログで「開始」ボタンを押下した際
- WizGrepSettings: WizGrep設定ダイアログで「OK」ボタンを押下した際

#### 5.7.5 エラー処理

- 保存失敗時: 例外を無視し、処理を継続
- 読み込み失敗時: デフォルト値の新しいインスタンスを返却

---

## 6. UI仕様

### 6.1 メイン画面構成

```
┌──────────────────────────────────────────────────────────┐
│ WizGrep                        [Grep設定] [WizGrep設定] [キャンセル] │
├──────────────────────────────────────────────────────────┤
│ ■■■■■■■■■■■■■■■■■（進捗バー: IsIndeterminate）          │
│ 処理中: report.xlsx (5/20)                                │
├─────────────────────┬────────────────────────────────────┤
│ マッチしたファイル [出力] │ 検索結果                  10件 [出力] │
├─────────────────────┼────────────────────────────────────┤
│ C:\...\report.xlsx   │ ファイル名  │  位置         │ 内容...  │
│ C:\...\manual.docx   │ report.xlsx │ [Sheet1] A1  │ 売上...  │
│                      │ manual.docx │ 15 行        │ 手順...  │
│                      │                                    │
└─────────────────────┴────────────────────────────────────┘
│ 検索完了: 10件見つかりました                                  │
└──────────────────────────────────────────────────────────┘
```

#### メイン画面レイアウト詳細

| 領域 | 説明 |
|------|------|
| タイトルバー | アプリ名「WizGrep」＋ボタン群（Grep設定・WizGrep設定・キャンセル） |
| 進捗表示 | 検索中のみ表示。プログレスバー（不定）＋ 処理中ファイル名 (処理済み/全体) |
| 左ペイン（幅300px固定） | マッチしたファイルの絶対パス一覧 ＋ 出力ボタン |
| 右ペイン（可変幅） | 検索結果詳細（ファイル名: 200px / 位置: 120px / 内容: 残り）＋ 件数 ＋ 出力ボタン |
| ステータスバー | 検索状態メッセージ（準備完了/検索中/検索完了/キャンセル/エラー） |

#### ボタン状態制御

| ボタン | 検索中 | 非検索時 |
|--------|--------|---------|
| Grep設定 | 無効 | 有効 |
| WizGrep設定 | 無効 | 有効 |
| キャンセル | 表示 | 非表示 |

#### 背景

Mica バックドロップ（`MicaBackdrop`）を使用。

### 6.2 Grep設定ダイアログ

ContentDialogとして表示されます。

| セクション | 設定項目 |
|-----------|---------|
| 検索対象フォルダ | フォルダパス入力欄、「参照...」ボタン（`FolderPicker`） |
| 検索キーワード | 5個のキーワード入力欄（各々に有効/無効チェックボックス付き） |
| 検索条件 | AND検索 / OR検索 ラジオボタン |
| 対象ファイル | Excel, Word, PowerPoint, PDF, Text, ALL チェックボックス |
| 追加拡張子 | 使用有無チェックボックス + カスタム拡張子入力欄 |
| 除外拡張子 | 使用有無チェックボックス + 除外拡張子入力欄 |
| オプション | 大文字小文字区別、正規表現、リアルタイム表示、Excel改行削除 |

**ALLチェックボックスの動作:**
- ONにすると、個別ファイル種別チェックボックスとカスタム拡張子入力欄が無効化される
- 拡張子に関係なく全ファイルが検索対象になる

**プライマリボタン（開始）の動作:**
1. `Validate()` を実行
2. 検証失敗時: ダイアログは閉じない（`args.Cancel = true`）
3. 検証成功時: 設定を保存 → 検索を開始

### 6.3 WizGrep設定ダイアログ

ContentDialogとして表示されます。

| セクション | 設定項目 |
|-----------|---------|
| インデックス保存先 | ベースパス入力欄、「参照」ボタン（`FolderPicker`） |
| インデックス再構築 | 「次回Grep時にインデックスを新規に作り直す」チェックボックス |

プライマリボタン（OK）で設定を保存します。

---

## 7. エラーハンドリング

### 7.1 ファイル読み込みエラー

- 全ファイルリーダーは `try-catch(Exception)` でラップ
- 読み込みに失敗したファイルはスキップし、空の結果を返却
- エラーメッセージの表示なし（サイレントスキップ）
- 処理は次のファイルに継続

### 7.2 正規表現パターンエラー

- 検索開始前に全キーワードのパターンを事前検証
- 無効パターンを検出した場合、`InvalidOperationException` をスロー
- ステータスバーにエラーメッセージを表示: `エラー: 正規表現パターンが無効です: "パターン" - 詳細`

### 7.3 検索キャンセル

- `CancellationTokenSource` / `CancellationToken` による協調キャンセル
- ファイル処理ループの先頭で `cancellationToken.ThrowIfCancellationRequested()` を呼び出し
- `OperationCanceledException` をキャッチし、「検索がキャンセルされました」を表示
- CancellationTokenSourceは `finally` ブロックで `Dispose()` → `null` に設定

### 7.4 設定保存/読み込みエラー

- 保存エラー: 例外を無視し処理継続
- 読み込みエラー: デフォルト値のインスタンスを返却

### 7.5 ダイアログエラー

- Grep設定ダイアログ・WizGrep設定ダイアログ表示中の例外は `ContentDialog` でエラーメッセージを表示

### 7.6 エクスポートエラー

- ファイル保存中の例外は `ContentDialog` でエラーメッセージを表示

---

## 8. 文字コード自動判定（EncodingDetector）

### 8.1 判定アルゴリズム

```
1. ファイル先頭4096バイトをバッファに読み込み
2. BOM判定（優先順位順）:
   a. UTF-8 BOM (EF BB BF)
   b. UTF-32 LE BOM (FF FE 00 00)  ← UTF-16 LEより先に判定
   c. UTF-32 BE BOM (00 00 FE FF)  ← UTF-16 BEより先に判定
   d. UTF-16 LE BOM (FF FE)
   e. UTF-16 BE BOM (FE FF)
3. BOMなしの場合:
   a. バッファ途中切断時はマルチバイト境界での誤判定を避けるため末尾4バイトを除外
   b. Strict UTF-8 デコードを試行
   c. Strict Shift-JIS デコードを試行
   d. UTF-8のみ有効 → UTF-8
   e. Shift-JISのみ有効 → Shift-JIS
   f. 両方有効（ASCII等） → UTF-8を優先
   g. データが2バイト未満 → UTF-8（デフォルト）
```

### 8.2 対応エンコーディング一覧

| エンコーディング | BOM検出 | ヒューリスティック判定 |
|----------------|---------|-------------------|
| UTF-8（BOM付き） | ? | ? |
| UTF-8（BOMなし） | ? | ? |
| UTF-16 LE | ? | ? |
| UTF-16 BE | ? | ? |
| UTF-32 LE | ? | ? |
| UTF-32 BE | ? | ? |
| Shift-JIS | ? | ? |
