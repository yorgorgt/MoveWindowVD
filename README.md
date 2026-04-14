# MoveWindowVD

Windows 11 でアクティブウィンドウを隣の仮想デスクトップへ移動するユーティリティです。

**Ctrl + Shift + Win + Left / Right** で、ウィンドウを左右の仮想デスクトップへ移動し、自動的にそのデスクトップへ切り替えます。

## Features

- 常駐型 (システムトレイアイコン)
- 単一インスタンス制御
- Windows 11 Build 22000 (21H2) / 26100 (24H2) 対応

## Usage

ビルドした `MoveWindowVD.exe` を起動するとシステムトレイに常駐します。

| Shortcut | Action |
|---|---|
| Ctrl + Shift + Win + ← | ウィンドウを左の仮想デスクトップへ移動 |
| Ctrl + Shift + Win + → | ウィンドウを右の仮想デスクトップへ移動 |

終了するにはトレイアイコンを右クリック → **終了** を選択してください。

## Notes

本ツールは Windows の非公開 COM インターフェース (`IVirtualDesktopManagerInternal`) を使用しています。Windows Update によりインターフェースの GUID や vtable レイアウトが変更される場合があります。

## License

[MIT](LICENSE)
