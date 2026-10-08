# 電子カルテ端末 台帳自動収集ツール

インターネットに接続されていない院内ネットワーク（電子カルテ系）で、Windows 端末の情報を
**追加ソフトなし（Windows 標準の PowerShell だけ）** で収集し、Excel で開ける一覧 CSV にまとめます。

## 収集できる項目

| 項目 | 取得元 | 備考 |
|---|---|---|
| タグNo | 資産台帳 CSV ＞ レジストリ `HKLM\SOFTWARE\HospitalInventory\AssetTag` ＞ BIOS の Asset Tag | Windows 自体はタグNoを持たないので、**台帳との突合が基本** |
| 端末名 | `COMPUTERNAME` | |
| IPアドレス / MACアドレス | WMI (`Win32_NetworkAdapterConfiguration`) | NIC が複数あれば `;` 区切り |
| Windowsバージョン | WMI + レジストリ | 例: `Microsoft Windows 10 Enterprise LTSC 21H2 (64 ビット)` とビルド番号 |
| Officeバージョン | Click-to-Run 設定 / アンインストール情報 | 2007〜2024・Microsoft 365 |
| Officeライセンスキー | ライセンス WMI / レジストリ | **下記の制限あり** |
| 起動日 | 最終起動日時 (`LastBootUpTime`) | 連続稼働日数も出力 |
| 端末稼働日 / 稼働月数 | 台帳の「導入日」＞ OS インストール日 | どちらを使ったかを「稼働日の根拠」列に出力 |

### Office ライセンスキーの制限（重要）

- **Office 2013 以降（2016/2019/2021/2024/M365）は、PC 内にフルのプロダクトキーが保存されていません。**
  取得できるのは末尾 5 文字（`XXXXX-XXXXX-XXXXX-XXXXX-ABCDE`）と認証状態（Licensed など）のみです。
  末尾 5 文字を購入時のキー一覧と照合すれば、どのキーを使っているかは特定できます。
- Office 2010 / 2007 はレジストリから復号してフルキーを出力します。
- ボリュームライセンス (KMS/MAK) も末尾 5 文字での照合になります。

### 「全端末」について

- 情報を取れるのは **その時点で起動している Windows 端末だけ** です。
  停止中の端末は、方式 A（起動時スクリプト）なら次回起動時に自動で集まります。
- 未収集の端末は、資産台帳・ネットワークスキャン結果と突合して「未収集」行として一覧に出すので、漏れが分かります。
- プリンタ・医療機器・Linux 機器などはネットワークスキャンで IP/MAC/名前のみ一覧化されます。

## ファイル構成

| ファイル | 役割 | 実行場所 |
|---|---|---|
| `Get-DeviceInventory.ps1` | 端末 1 台分の情報を収集し `<端末名>.csv` を出力 | 各端末 |
| `run-inventory.bat` | 上記を起動するバッチ（GPO・タスクスケジューラ用） | 各端末 |
| `Invoke-RemoteInventory.ps1` | 管理 PC から WinRM で全端末に一括実行 | 管理 PC |
| `Find-NetworkHosts.ps1` | Ping スキャンで生きている IP / MAC / 名前を一覧化 | 管理 PC |
| `Merge-Inventory.ps1` | 全 CSV ＋ 台帳 ＋ スキャン結果を 1 つの一覧に統合 | 管理 PC |
| `sample/asset-master.csv` | 資産台帳のサンプル（列名はこのまま使ってください） | — |

## 手順

### 準備：収集用の共有フォルダ

ファイルサーバーに例として `\\fs01\inventory$` を作成し、

- `\\fs01\inventory$\Get-DeviceInventory.ps1`（スクリプト本体）
- `\\fs01\inventory$\raw\`（各端末の CSV 出力先）

を置きます。`raw` フォルダのアクセス権は

- **Domain Computers**（起動時スクリプトの場合）または **Domain Users**（ログオンスクリプトの場合）: 「ファイルの作成/データの書き込み」「変更」
- **情報システム担当者のみ**: 読み取り

としてください（ライセンス情報を含むため、一般利用者が読めないようにします）。

### 方式 A：起動時スクリプト（推奨・ドメイン環境）

1. グループポリシー管理で電子カルテ端末の OU に GPO を作成
2. 「コンピューターの構成 → ポリシー → Windows の設定 → スクリプト → スタートアップ」に
   `run-inventory.bat` を登録（中の `SHARE=` を自分の共有パスに変更）
3. 各端末が起動するたびに `raw\<端末名>.csv` が最新に上書きされます

ワークグループ環境の場合は、各端末のタスクスケジューラに「起動時」「SYSTEM で実行」で
`run-inventory.bat` を登録するか、USB メモリから `-OutputDir` を USB のフォルダにして 1 台ずつ実行してください。

### 方式 B：管理 PC から一括実行（WinRM が有効な場合）

```powershell
# AD の全コンピューターに実行
.\Invoke-RemoteInventory.ps1 -FromActiveDirectory -OutputDir C:\inventory\raw

# 端末名/IP のリストから実行（hosts.txt に 1 行 1 台）
.\Invoke-RemoteInventory.ps1 -ComputerListFile .\hosts.txt -OutputDir C:\inventory\raw -Credential (Get-Credential)
```

失敗した端末は `raw\_failed.txt` に出ます（電源 OFF・WinRM 無効・FW など）。

### 未収集の機器を洗い出す（任意）

```powershell
.\Find-NetworkHosts.ps1 -Subnet 192.168.10,192.168.11 -OutFile C:\inventory\discovery.csv
```

MAC アドレスは ARP で取るため、**スキャンする PC と同じセグメントの機器のみ** 取得できます。
セグメントが分かれている場合は各セグメントの PC で実行するか、L3 スイッチ / DHCP サーバーの ARP・リース表を使ってください。

### 一覧の作成

```powershell
.\Merge-Inventory.ps1 -RawDir \\fs01\inventory$\raw `
    -AssetMaster .\asset-master.csv `
    -Discovery C:\inventory\discovery.csv `
    -OutFile .\端末一覧.csv
```

`端末一覧.csv` をそのまま Excel で開けます（UTF-8 BOM 付き）。資産台帳は Excel で
「CSV (コンマ区切り)」保存したもの（Shift-JIS）でも UTF-8 でも読めます。

#### 出力列

`タグNo, 端末名, IPアドレス, MACアドレス, Windowsバージョン, Officeバージョン, Officeライセンスキー,
起動日, 連続稼働日数, 端末稼働日, 端末稼働月数, 稼働日の根拠, ライセンス状態, Windowsビルド,
メーカー, 機種, シリアル番号, 設置場所, 最終収集日時, 状態`

「状態」は `収集済` / `収集済(30日以上前)` / `未収集(台帳のみ)` / `未収集(ネットワーク検出のみ)` のいずれかです。

## 注意事項

- **導入前に電子カルテベンダーへ確認してください。** 端末へのスクリプト配布や GPO 変更が
  保守契約上ベンダー承認事項になっている場合があります。スクリプトは読み取りのみで設定変更は行いません。
- 「端末稼働日」に OS インストール日を使う場合、再セットアップや Windows の大型アップデートで日付が更新されます。
  正確な稼働月数が必要な場合は資産台帳の「導入日」を記入してください。
- 「起動日」は Windows の高速スタートアップが有効だと、シャットダウン→電源 ON では更新されません（再起動で更新）。
- Windows 7 (PowerShell 2.0) 以降で動作します。`Merge-Inventory.ps1` / `Find-NetworkHosts.ps1` は管理 PC で
  Windows PowerShell 5.1 以降を使ってください。
