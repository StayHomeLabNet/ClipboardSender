# Clipboard Sync（Windows 常駐トレイアプリ）

> 旧名：**Clipboard Sender**（v1.0.0 で *受信* 対応により **Clipboard Sync** に改名）

**Clipboard Sync** は **Notemod-selfhosted** 用の軽量な Windows トレイアプリです。  
クリップボード共有が **双方向** になります。

- **送信**：Windows のクリップボード → Notemod-selfhosted（書き込み）
- **受信**：Notemod-selfhosted → Windows のクリップボード →（任意）**自動ペースト（Ctrl+V）**

さらに、INBOX の定期削除（Cleanup）と、サーバーに溜まりがちな **バックアップファイルの一括削除** も行えます。

## 主な機能

- トレイ常駐（メイン画面なし）
- トレイアイコン左クリックで「送信の有効/無効」を切り替え
- グローバルホットキーで送信の有効/無効切り替え
- **送信**：クリップボード文字列を Notemod-selfhosted に POST
- **受信**：`read_api.php` で最新ノートを取得してクリップボードへコピー
  - 受信後に **自動ペースト（Ctrl+V）** まで実行（任意）
  - ペースト前に「クリップボード安定待ち（ms）」を挟める（任意）
- 削除（Cleanup）
  - 手動：**INBOX 全削除**
  - 自動：**定時（1日1回）** または **X 分毎**
- バックアップ管理（Notemod-selfhosted 側でバックアップ作成を有効にしている場合）
  - **バックアップファイル数** の表示
  - **バックアップファイルの一括削除** ボタン
- BASIC 認証（任意）
- UI 言語：English / 日本語 / Türkçe

## 想定用途

Notemod-selfhosted（サーバー） + Clipboard Sync（Windows） + iPhone のショートカット/背面タップ/アクションボタン等で、

- iPhone でコピーした文字列を、Windows 側で **ホットキーを1回押すだけ** で貼り付けまで完了

という使い方を想定しています。

## 必要なもの

- Windows 10 / 11
- .NET 8 Desktop Runtime（配布 EXE 実行用）
- Notemod-selfhosted が下記 API を提供していること
  - `api.php`（送信）
  - `read_api.php`（受信）
  - `cleanup_api.php`（INBOX 全削除 + バックアップファイル一括削除）

> サーバーが BASIC 認証で保護されている場合は、設定画面で入力してください。

## インストール

1. GitHub Releases から最新版をダウンロード
2. EXE を起動
3. トレイメニューの **設定** から URL / token を入力

## 使い方

### 基本動作

- アプリはトレイに常駐します
- トレイアイコン左クリックで「送信」の ON/OFF
- 右クリックメニューから：
  - INBOX 全削除
  - 設定
  - ヘルプ / バージョン情報

### 設定画面（タブで切り替え）

- **送信/削除設定**
  - POST 先 URL + token（送信）
  - BASIC 認証（任意）
  - Cleanup_API URL + token（INBOX 全削除）
  - 自動削除スケジュール（定時 / X 分毎）
  - バックアップファイル数表示 + 一括削除ボタン
  - 言語
- **受信設定**
  - Read API URL
  - 受信ホットキー
  - ペーストまで含める（ON/OFF）
  - クリップボード安定待ち（ms）

> ※ 自動ペーストは、貼り付け先アプリが **管理者権限** で動いていると拒否されることがあります。  
> まずはメモ帳などで動作確認してください。

## 開発者向け（ビルド）

- Visual Studio 2022 + .NET 8 SDK + WinForms ワークロード

## セキュリティ / プライバシー

- 送信するのは「コピーした文字列」です（送信が ON のときのみ）
- token / パスワードは `settings.json` に保存しますが、**DPAPI（Windows ユーザースコープ）** で暗号化されます
- インターネット公開する場合は HTTPS + BASIC 認証等を推奨します

## Links

- Notemod-selfhosted: https://github.com/StayHomeLabNet/Notemod-selfhosted
- Help: https://stayhomelab.net/Clipboardsync
- GitHub: https://github.com/StayHomeLabNet/ClipboardSync
