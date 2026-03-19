# PaneNest

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

## 実行方法

```powershell
dotnet run --project .\src\PaneNest.Desktop\PaneNest.Desktop.csproj
```

## ビルド

```powershell
dotnet build .\src\PaneNest.Desktop\PaneNest.Desktop.csproj
```

## 配布用ビルド

```powershell
.\scripts\publish-release.cmd
```

配布用の出力先は `dist\PaneNest-win-x64` に固定しています。スクリプト実行時に古い `dist` 配下のバリエーション出力は自動で削除されます。

## メタデータ保存先

タグ、メモ、お気に入り情報はドライブ単位で次の場所へ保存されます。

```text
%LOCALAPPDATA%\PaneNest\
```

## 既知の制約

- PDF の埋め込みプレビューには未対応です
- サムネイル生成キャッシュは未実装です
- 大きなライブラリでは初回読み込みに時間がかかる場合があります
