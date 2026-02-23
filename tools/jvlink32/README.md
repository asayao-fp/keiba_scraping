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

### 出力 JSON の例（正常時）

```json
{
  "ok": true,
  "stage": "close",
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

### 出力 JSON の例（エラー時）

```json
{
  "ok": false,
  "stage": "read",
  "hresult": "0x80010105",
  "error": "The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))",
  "setup": { "init": 0, "save_path": 0, "save_flag": 0, "pay_flag": 0 },
  "open": { "ret": 0, "readcount": 0, "downloadcount": 0, "lastfiletimestamp": "" },
  "status_poll": [
    { "timestamp": "2026-02-23T02:00:00.000Z", "status": 1 },
    { "timestamp": "2026-02-23T02:00:00.500Z", "status": 0 }
  ]
}
```

JSONには少なくとも `ok`, `stage`, `open.ret`, `open.readcount`, `open.downloadcount`,
`read.ret`, `read.size`, `read.filename`, `close` が含まれます。
エラー時は `stage`, `hresult`（COMException の場合）, `error` が追加されます。

---

## 環境変数一覧

| 変数名 | デフォルト | 説明 |
|---|---|---|
| `JV_DATASPEC` | `RACE` | JVOpen の dataspec |
| `JV_FROMDATE` | `20240101000000` | JVOpen の fromdate |
| `JV_OPTION` | `1` | JVOpen の option |
| `JV_SAVE_PATH` | `C:\ProgramData\JRA-VAN\Data` | JVSetSavePath のパス |
| `JV_READ_MAX_WAIT_SEC` | `60` | JVRead リトライの最大待機時間（秒） |
| `JV_READ_INTERVAL_SEC` | `0.5` | JVRead リトライ間隔（秒） |
| `JV_SLEEP_AFTER_OPEN_SEC` | `1.0` | JVOpen 後、JVRead 前のスリープ時間（秒）。`0` で無効化 |
| `JV_ENABLE_UI_PROPERTIES` | `0` | `1` にすると `JVSetUIProperties` をデフォルト値で呼び出す |
| `JV_ENABLE_STATUS_POLL` | `0` | `1` にすると JVOpen 後に `JVStatus` を繰り返しポーリングする |
| `JV_STATUS_POLL_MAX_WAIT_SEC` | `10` | `JVStatus` ポーリングの最大待機時間（秒） |
| `JV_STATUS_POLL_INTERVAL_SEC` | `0.5` | `JVStatus` ポーリング間隔（秒） |
| `JV_READ_REQUIRE_STATUS_ZERO` | `0` | `1` にすると `JVStatus()==0` が確認されるまで `JVRead` を呼ばない。ポーリング期間内に `0` にならない場合は `ok=false` / `stage="status_poll"` を出力して終了する |

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

`ok=false` のとき、ラッパーはエラー内容（stage, hresult, error）を stderr に出力し、
終了コード 1 を返します。JSON 本体は stdout に出力されます。

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

### `RPC_E_SERVERFAULT (0x80010105)` が `JVRead` で発生する

`JVRead` で `COMException: 0x80010105` が発生する場合、以下の手順で診断情報を収集してください。

> **JVRead の ByRef 呼び出しについて**
> ブリッジは `System.Reflection.ParameterModifier` を使用して `JVRead` の全引数を
> ByRef（参照渡し）としてマーシャリングします。これにより、レイトバインド反射での
> ByRef 引数の不正マーシャリングが原因で発生する `RPC_E_SERVERFAULT` を回避できます。
> 診断には各呼び出し試行の `read_args_types` と `read_args_values_preview` が含まれます。

#### ステップ 1: 診断モードで実行

```powershell
$env:JV_SLEEP_AFTER_OPEN_SEC      = "3"   # JVOpen 後に 3 秒待機
$env:JV_ENABLE_STATUS_POLL        = "1"   # JVStatus をポーリングして状態遷移を確認
$env:JV_STATUS_POLL_MAX_WAIT_SEC  = "15"
$env:JV_STATUS_POLL_INTERVAL_SEC  = "1"
$env:JV_ENABLE_UI_PROPERTIES      = "1"   # JVSetUIProperties を呼び出す（環境によって有効）
$env:JV_READ_REQUIRE_STATUS_ZERO  = "1"   # JVStatus==0 が確認されるまで JVRead を呼ばない
.\JVLinkBridge.exe RACE 20240101000000 1
```

出力 JSON の `status_poll` フィールドで、`JVStatus` の遷移を確認します。
`status=0` が現れる前に `JVRead` を呼び出すとサーバーフォルトが発生することがあります。
`JV_READ_REQUIRE_STATUS_ZERO=1` を設定すると、`status=0` が確認されるまで `JVRead` を呼ばないため、サーバーフォルトを回避できます。

#### ステップ 2: JSON のフィールドを確認

| フィールド | 意味 |
|---|---|
| `stage` | エラーが発生した段階（`init` / `open` / `status_poll` / `read` / `close`） |
| `hresult` | COMException の HRESULT 値（例: `0x80010105`） |
| `error` | エラーメッセージ |
| `open.readcount` | `JVOpen` が返したレコード数（0 の場合はデータなし） |
| `status_poll` | `JVStatus` のポーリング結果（`JV_ENABLE_STATUS_POLL=1` または `JV_READ_REQUIRE_STATUS_ZERO=1` のとき） |
| `read.attempts_tail[].read_args_types` | `JVRead` 呼び出し後の各引数の型名（ByRef マーシャリングの確認に使用） |
| `read.attempts_tail[].read_args_values_preview` | `JVRead` 呼び出し後の各引数の値プレビュー（filename, size, buff 先頭） |

#### ステップ 3: JV-Link の状態を確認

- JV-Link の設定ダイアログ（タスクトレイアイコン）を一度開いてサービスキーを確認する
- OS を再起動してから再試行する
- JV-Link を再インストールする
