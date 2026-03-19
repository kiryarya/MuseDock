# MuseDock

Windows 向けの WPF 製ファイル管理アプリです。ドライブ内のフォルダとファイルを一覧表示し、プレビュー、お気に入り、タグ、メモの管理を行えます。

## 主な機能

- ドライブ一覧の表示
- フォルダ移動と子フォルダの一覧表示
- ファイルとフォルダの検索
- 種類フィルタ
- お気に入り、タグ、メモの保存
- 画像、動画、テキストのプレビュー
- PDF の外部アプリ起動
- Explorer での表示
- 名前変更とごみ箱への削除
- `Plugins` フォルダからのプラグイン読み込み
- 右クリックメニューからの外部プラグイン実行

## 実行方法

```powershell
dotnet run --project .\src\MuseDock.Desktop\MuseDock.Desktop.csproj
```

## ビルド

```powershell
dotnet build .\src\MuseDock.Desktop\MuseDock.Desktop.csproj
```

## 配布用ビルド

```powershell
.\scripts\publish-release.cmd
```

配布用の出力先は `dist\MuseDock-win-x64` に固定しています。スクリプト実行時に古い `dist` 配下のバリエーション出力は自動で削除されます。

## メタデータ保存先

タグ、メモ、お気に入り情報はドライブ単位で次の場所へ保存されます。

```text
%LOCALAPPDATA%\MuseDock\
```

## プラグイン

MuseDock は外部プラグインを読み込めます。右クリックメニューの `プラグイン` から実行します。

読み込み先は次の 2 か所です。

```text
<MuseDock の実行フォルダ>\Plugins\
%LOCALAPPDATA%\MuseDock\plugins\
```

各プラグインはフォルダごとに配置し、その中に `plugin.json` を置きます。

```text
plugins\
  MyPlugin\
    plugin.json
    run.ps1
```

### `plugin.json` の例

```json
{
  "id": "my.image.ai",
  "name": "画像 AI ツール",
  "version": "1.0.0",
  "description": "選択画像を外部 AI ツールへ渡すプラグインです。",
  "commands": [
    {
      "id": "edit-image",
      "name": "画像を AI 編集",
      "runner": "powershell",
      "entry": "run.ps1",
      "arguments": [
        "-ContextPath",
        "{contextPath}"
      ],
      "requiresSelection": true,
      "waitForExit": true,
      "refreshAfterRun": true,
      "assetKinds": ["image"],
      "extensions": [".png", ".jpg", ".webp"]
    }
  ]
}
```

### 対応 runner

- `powershell`
- `pwsh`
- `exe`

`exe` は `python` や `node` のように PATH 上のコマンドも使えます。

### プラグインへ渡される情報

MuseDock は実行時に JSON コンテキストファイルを生成し、`{contextPath}` と環境変数 `MUSEDOCK_CONTEXT_PATH` でプラグインへ渡します。

主な内容は次の通りです。

- 現在開いているフォルダ
- 現在のドライブ
- 選択中ファイルのパス、種類、タグ、メモ、お気に入り状態
- 実行したプラグインとコマンドの情報

また次の置換トークンも使えます。

- `{contextPath}`
- `{pluginDirectory}`
- `{appDirectory}`
- `{currentDirectory}`
- `{currentDrive}`
- `{selectedPath}`

### サンプル

同梱サンプルは次にあります。

```text
src\MuseDock.Desktop\Plugins\Samples\SelectionContextExporter\
```

配布版にも `Plugins\Samples\SelectionContextExporter\` として同梱されます。選択ファイルのコンテキスト JSON を `%TEMP%\MuseDock\SamplePlugin\` にコピーするだけの最小サンプルです。

簡易の開発手順書は [docs/plugin-development-guide.md](C:/Users/tapih/Downloads/Programming/ファイル管理ソフト/docs/plugin-development-guide.md) にあります。

画像編集系の参考用サンプルは `src\MuseDock.Desktop\Plugins\Samples\SimpleImageTools\` に追加しています。

## 既知の制約

- PDF の埋め込みプレビューには未対応です
- サムネイル生成キャッシュは未実装です
- 大きなライブラリでは初回読み込みに時間がかかる場合があります
