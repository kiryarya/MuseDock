# MuseDock プラグイン開発ガイド

この資料は、MuseDock 向けのプラグインを最小構成で作るための簡易手順書です。  
まずは `plugin.json` と実行スクリプト 1 本で動くところまでを対象にしています。

## 1. できること

MuseDock のプラグインは、右クリックメニューの `プラグイン` から実行される外部コマンドです。

現時点では次のような用途を想定しています。

- 選択中画像を AI 編集ツールへ渡す
- 選択中ファイルを別形式へ変換する
- 選択中ファイルの情報を外部 API へ送る
- 現在フォルダの内容をまとめて処理する

MuseDock 本体は、選択中ファイルや現在フォルダなどの情報を JSON でプラグインへ渡します。  
プラグイン側はその JSON を読んで自由に処理します。

## 2. 配置場所

プラグインは次のどちらかに置きます。

```text
<MuseDock の実行フォルダ>\Plugins\
%LOCALAPPDATA%\MuseDock\plugins\
```

開発中は `%LOCALAPPDATA%\MuseDock\plugins\` を使うと差し替えが楽です。

例:

```text
%LOCALAPPDATA%\MuseDock\plugins\
  MyFirstPlugin\
    plugin.json
    run.ps1
```

## 3. 最小構成

最低限必要なのは次の 2 つです。

- `plugin.json`
- 実行ファイルまたはスクリプト

PowerShell 版の最小構成例です。

### `plugin.json`

```json
{
  "id": "my.first.plugin",
  "name": "最初のプラグイン",
  "version": "1.0.0",
  "description": "選択情報を受け取る最小サンプルです。",
  "commands": [
    {
      "id": "hello-context",
      "name": "選択情報を表示",
      "runner": "powershell",
      "entry": "run.ps1",
      "arguments": [
        "-ContextPath",
        "{contextPath}"
      ],
      "requiresSelection": true,
      "waitForExit": true,
      "refreshAfterRun": false,
      "assetKinds": ["image", "video", "text", "pdf", "other", "folder"]
    }
  ]
}
```

### `run.ps1`

```powershell
param(
    [Parameter(Mandatory = $true)]
    [string]$ContextPath
)

$context = Get-Content $ContextPath -Raw | ConvertFrom-Json

[System.Windows.MessageBox]::Show(
    "現在フォルダ: $($context.currentDirectory)`n選択数: $($context.selectedItems.Count)",
    "MyFirstPlugin"
)
```

## 4. `plugin.json` の主要項目

### プラグイン全体

- `id`
  一意な ID。重複不可です。
- `name`
  MuseDock 上に表示されるプラグイン名です。
- `version`
  文字列であれば任意形式で構いません。
- `description`
  メニューの補足説明に使われます。
- `commands`
  実行できるコマンド一覧です。

### command

- `id`
  コマンドの一意 ID です。
- `name`
  右クリックメニューに出る名前です。
- `runner`
  `powershell` / `pwsh` / `exe` に対応しています。
- `entry`
  実行ファイル名またはスクリプト名です。
- `arguments`
  実行時引数です。置換トークンを使えます。
- `requiresSelection`
  `true` ならファイル未選択時はメニューに出しません。
- `waitForExit`
  `true` なら処理完了まで待ちます。
- `refreshAfterRun`
  `true` なら実行後に現在フォルダを再読み込みします。
- `assetKinds`
  表示対象の種類です。`folder` / `image` / `video` / `text` / `pdf` / `other`
- `extensions`
  拡張子で対象を絞れます。例: `[".png", ".jpg"]`

## 5. 置換トークン

`arguments` では次のトークンを使えます。

- `{contextPath}`
  MuseDock が生成した JSON コンテキストファイルのパス
- `{pluginDirectory}`
  そのプラグインのフォルダ
- `{appDirectory}`
  MuseDock の実行フォルダ
- `{currentDirectory}`
  現在開いているフォルダ
- `{currentDrive}`
  現在のドライブ
- `{selectedPath}`
  代表の選択ファイルパス

## 6. 環境変数

MuseDock はプラグイン実行時に次の環境変数も渡します。

- `MUSEDOCK_CONTEXT_PATH`
- `MUSEDOCK_PLUGIN_DIRECTORY`
- `MUSEDOCK_APP_DIRECTORY`
- `MUSEDOCK_CURRENT_DIRECTORY`
- `MUSEDOCK_CURRENT_DRIVE`
- `MUSEDOCK_SELECTED_PATH`

`arguments` で渡すより、環境変数から取りたい場合はこちらを使えます。

## 7. コンテキスト JSON の例

プラグインが受け取る JSON は概ね次の形です。

```json
{
  "appName": "MuseDock",
  "invokedAt": "2026-03-19T18:20:00+09:00",
  "pluginId": "my.first.plugin",
  "pluginName": "最初のプラグイン",
  "commandId": "hello-context",
  "commandName": "選択情報を表示",
  "pluginDirectory": "C:\\Users\\...\\plugins\\MyFirstPlugin",
  "appDirectory": "C:\\Users\\...\\MuseDock-win-x64",
  "currentDirectory": "C:\\Images",
  "currentDrivePath": "C:\\",
  "selectedItems": [
    {
      "filePath": "C:\\Images\\sample.png",
      "name": "sample.png",
      "extension": ".png",
      "assetKind": "image",
      "isDirectory": false,
      "tags": ["sample"],
      "note": "",
      "isFavorite": false
    }
  ]
}
```

## 8. 開発の流れ

1. `%LOCALAPPDATA%\MuseDock\plugins\MyPlugin\` を作る
2. `plugin.json` を置く
3. `run.ps1` などの実行ファイルを置く
4. MuseDock を起動する
5. 対象ファイルを右クリックする
6. `プラグイン > 自分のプラグイン名 > コマンド名` を実行する
7. 反映されない場合は `プラグイン > プラグインを再読み込み` を使う

## 9. AI 画像編集プラグインの作り方

たとえば AI 画像編集プラグインなら、流れは次のようになります。

1. `assetKinds` を `["image"]` にする
2. `extensions` を `[".png", ".jpg", ".webp"]` などにする
3. `selectedItems[0].filePath` を AI ツールへ渡す
4. 編集結果を元フォルダに保存する
5. `refreshAfterRun` を `true` にして MuseDock 側を再読み込みさせる

## 10. まず見るべきサンプル

同梱サンプル:

```text
src\MuseDock.Desktop\Plugins\Samples\SelectionContextExporter\
```

このサンプルは、受け取ったコンテキスト JSON をそのまま `%TEMP%\MuseDock\SamplePlugin\` へコピーするだけです。  
最初はこれを複製して名前を変えるのが最短です。

画像編集プラグインの参考例:

```text
src\MuseDock.Desktop\Plugins\Samples\SimpleImageTools\
```

このサンプルには次の 2 コマンドが入っています。

- グレースケール複製を作成
- 50% 縮小複製を作成

どちらも元画像を上書きせず、同じフォルダに別名保存します。  
AI 画像編集プラグインを作る場合も、このサンプルの `plugin.json` と `run.ps1` を複製して、画像処理部分だけ差し替えるのが簡単です。

## 11. プラグインの追加手順

MuseDock に新しいプラグインを追加する最短手順です。

1. `%LOCALAPPDATA%\MuseDock\plugins\` を開く
2. その中に新しいフォルダを作る
3. `plugin.json` を置く
4. 実行スクリプトや実行ファイルを置く
5. MuseDock を起動する
6. 対象ファイルを右クリックして `プラグイン` を開く
7. 反映されない場合は `プラグインを再読み込み` を押す

フォルダ例:

```text
%LOCALAPPDATA%\MuseDock\plugins\
  MyImagePlugin\
    plugin.json
    run.ps1
```

最初は `src\MuseDock.Desktop\Plugins\Samples\SimpleImageTools\` を丸ごと複製して、`id` と `name` を変えるのが一番速いです。

## 12. 注意点

- 長時間かかる処理は外部プロセス側で進捗表示を出した方が安全です
- `waitForExit = false` の場合、MuseDock は処理完了を待ちません
- PowerShell 実行ポリシーの影響を避けるため、MuseDock は `-ExecutionPolicy Bypass` 付きで実行します
- プラグイン側のエラーは MuseDock 上では「実行失敗」として扱われます

## 13. 今後追加したいもの

現時点では簡易版です。今後の拡張候補は次の通りです。

- 複数選択を前提にした UI
- プラグインごとの設定画面
- 標準 API の追加
- プラグイン専用パネル
- 実行ログビューア
