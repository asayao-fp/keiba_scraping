# tools/jvlink32 – JV-Link デバッグツール群

## 概要

JV-Link COM サーバー（`JVDTLab.JVLink`）を Python/pywin32 から操作するデバッグ用スクリプトと、
`JVRead` クラッシュ（終了コード `0xC0000409`）を回避するための .NET ブリッジで構成されます。

---

## 前提条件

| 項目 | 内容 |
|---|---|
| OS | Windows 10 / 11 (x86 or x64) |
| JV-Link | バージョン 4.9.0 以降がインストール済み |
| .NET SDK | .NET 8 SDK 以降（`dotnet` コマンドが使えること） |
| Python | 3.11 以降（`jvread_via_bridge.py` を使う場合） |

> **注意**: JV-Link の COM サーバーは 32-bit (x86) のみです。
> ブリッジは `<PlatformTarget>x86</PlatformTarget>` でビルドします。

---

## JVLinkBridge（.NET ブリッジ）のビルド

```powershell
cd tools\jvlink32\JVLinkBridge
dotnet build -c Release
```

成功すると以下の場所に実行ファイルが生成されます:

```
tools\jvlink32\JVLinkBridge\bin\Release\net8.0-windows\JVLinkBridge.exe
```

---

## JVLinkBridge の直接実行（PowerShell）

```powershell
# 環境変数で制御する例
$env:JV_DATASPEC          = "RACE"
$env:JV_FROMDATE          = "20240101000000"
$env:JV_OPTION            = "1"
$env:JV_SAVE_PATH         = "C:\ProgramData\JRA-VAN\Data"
$env:JV_READ_MAX_WAIT_SEC = "60"
$env:JV_READ_INTERVAL_SEC = "0.5"

.\tools\jvlink32\JVLinkBridge\bin\Release\net8.0-windows\JVLinkBridge.exe
```

位置引数で渡すこともできます（`dataspec fromdate option`）:

```powershell
.\JVLinkBridge.exe RACE 20240101000000 1
```

### 出力 JSON の例

```json
{
  "ok": true,
  "setup": { "init": 0, "save_path": 0, "save_flag": 0, "pay_flag": 0 },
  "open": {
    "dataspec": "RACE",
    "fromdate": "20240101000000",
    "option": 1,
    "ret": 1,
    "readcount": 42,
    "downloadcount": 0,
    "lastfiletimestamp": "20240201120000"
  },
  "read": {
    "found": true,
    "ret": 0,
    "size": 512,
    "filename": "RA2024010101.jvd",
    "buff_head": "RA20240101...",
    "attempts_tail": [...]
  },
  "close": 0
}
```

JSONには少なくとも `open.ret`, `open.readcount`, `open.downloadcount`,
`read.ret`, `read.size`, `read.filename`, `close` が含まれます。

---

## Python ラッパー（jvread_via_bridge.py）の実行

ブリッジをビルドした後、Python から呼び出せます:

```powershell
# 環境変数で制御
$env:JV_DATASPEC = "RACE"
python tools\jvlink32\jvread_via_bridge.py

# 位置引数で制御
python tools\jvlink32\jvread_via_bridge.py RACE 20240101000000 1
```

---

## 既存デバッグスクリプトとの関係

| ファイル | 用途 |
|---|---|
| `jvlink_open_debug.py` | `JVRead` を **呼ばない** 安全版デバッグ（`JVOpen` まで確認） |
| `jvread_driver.py` | `JVRead` を直接呼ぶ（`0xC0000409` でクラッシュする可能性あり） |
| `jvread_via_bridge.py` | **推奨**: .NET ブリッジ経由で `JVRead` を安全に呼ぶ |
| `JVLinkBridge/Program.cs` | .NET ブリッジ本体 |

> `jvlink_open_debug.py` は現在 `JVRead` の実呼び出し行をコメントアウトし
> ダミー値を返す安全な状態になっています。
> `JVRead` の結果が必要な場合は `jvread_via_bridge.py` を使用してください。

---

## トラブルシューティング

### `JVDTLab.JVLink ProgID not found`

JV-Link がインストールされていないか、COM 登録が失われています。
JV-Link を再インストールして試してください。

### `JVOpen returned -201` など

JV-Link のサービスキーが未設定またはサブスクリプションが失効している可能性があります。
JV-Link の設定ダイアログを起動して確認してください。

### ビルドエラー `NETSDK1032: The RuntimeIdentifier platform ...`

x86 ターゲットの SDK が必要です。以下を確認してください:

```powershell
dotnet --list-sdks
```
