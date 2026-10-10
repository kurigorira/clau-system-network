# 電子カルテ端末 台帳自動収集ツール

インターネットにつながっていない電子カルテ網（AD ドメイン環境）で、全 Windows 端末の情報を自動で集め、
**Excel 台帳（.xlsx）** を作ります。Windows 標準の PowerShell だけで動き、追加ソフトや外部モジュールは不要です。

- 端末側: Windows 7 / 8.1 / 10 / 11（PowerShell 2.0 以降）
- 管理 PC 側（集計）: Windows PowerShell 5.1 以降。**Excel が入っていなくても .xlsx を作れます**

## 全体の流れ

```
[サーバー] \\nagasakinet.local\dfsroot\newton\startup\inventory\  （terminal\ の 2 ファイルを置く）
    │ 方式A: 端末起動時のコピー用バッチで C:\inventory へコピー → 実行（GPO 追加不要）
    │ 方式B: GPO のスケジュールタスク（毎日＋起動5分後）
    ▼
[各端末] C:\inventory\Get-DeviceInventory.ps1 ──► \\<結果共有>\<端末名>.csv
[管理PC] admin\Merge-Inventory.ps1
    ├─ 結果共有の *.csv を全部読む
    ├─ 台帳.csv と突合（タグNo・導入日・設置場所は手入力分を使う / 新しい端末は行を追加）
    └─► 端末一覧.xlsx（端末一覧 / 未収集 / 集計 の3シート）
```

## 一覧に出る項目

| 列 | 取得方法 |
|---|---|
| タグNo | **台帳.csv に手入力**（BIOS に Asset Tag が設定済みなら初回に自動で入力） |
| 端末名 / IPアドレス / MACアドレス | 端末から取得（NIC が複数あれば `;` 区切り） |
| Windowsバージョン | 端末から取得（例: `Microsoft Windows 10 Enterprise LTSC 21H2 (64 ビット)`） |
| Officeバージョン | 端末から取得（2007〜2024・Microsoft 365） |
| Officeライセンスキー | 端末から取得。**Office 2013 以降は末尾 5 桁のみ**（下記参照） |
| 起動日 | 最後に起動した日時 |
| 端末稼働日(導入日) | 台帳.csv の「導入日」を優先。未入力なら端末から推定（下記参照） |
| 端末稼働月数 | 導入日から今日までの月数 |
| 導入日の根拠 | `台帳` または `推定(端末)` |
| 設置場所 | 台帳.csv に手入力 |
| ライセンス状態 / メーカー / 機種 / シリアル番号 / BIOS日付 / 最終収集日時 / 状態 | 参考情報 |

「状態」と行の色:

| 状態 | 色 |
|---|---|
| 収集済 | なし |
| 収集済(30日以上前) | 黄色（最近電源が入っていない・ネットワーク不通・撤去済みの可能性） |
| 未収集(台帳のみ) | 赤（台帳にはあるが一度も収集できていない） |
| 未収集(ネットワーク検出のみ) | 赤（`Find-NetworkHosts.ps1` で見つかったが CSV が無い。プリンタ・医療機器など） |

「集計」シートには、収集済・未収集の台数、Windows 別台数、Office 別台数、稼働月数の帯別台数（〜36 / 37〜60 / 61か月〜）が出ます。

### Office ライセンスキーの制限

- **Office 2013 以降（2016/2019/2021/2024/M365）は、PC の中に全桁のキーが保存されていません。**
  取れるのは末尾 5 桁（`XXXXX-XXXXX-XXXXX-XXXXX-ABCDE`）と認証状態だけです。購入時のキー一覧と末尾 5 桁を照合してください。
- Office 2010 / 2007 は、レジストリから全桁を復元して出力します。

### 導入日の推定方法（台帳に導入日が無い端末）

次の日付のうち**いちばん古いもの**を推定導入日とします。

1. Windows のインストール日
2. Windows 10/11 の大型アップデート前の元のインストール日（`HKLM\SYSTEM\Setup\Source OS (…)`）
3. 最も古いユーザープロファイル（`C:\Users\<ユーザー>`）の作成日

再セットアップした端末は、再セットアップした日になります。参考として BIOS日付（製造時期の目安）も出すので、
大きく食い違う端末は台帳.csv に正しい導入日を入力してください。

## ファイル構成

```
inventory\
├ terminal\                     … 中の ps1 と run-inventory.bat の 2 つだけを startup\inventory\ の直下に置く
│  ├ Get-DeviceInventory.ps1     端末 1 台分の情報を収集
│  ├ run-inventory.bat           ps1 を実行して結果共有へ CSV を書き込む（先頭の SHARE= を設定）
│  ├ check-inventory.bat         動かないときの診断用（ダブルクリックで各段階を確認。サーバーには置かない）
│  └ startup-snippet.bat         既存のスタートアップ用バッチに追加する 2 行の見本（サーバーには置かない）
├ admin\                        … 管理 PC で使う
│  ├ Merge-Inventory.ps1         全端末分を集計し、台帳.csv を更新、.xlsx を出力
│  ├ lib\Write-Xlsx.ps1          .xlsx の書き出し（Merge から読み込まれる）
│  ├ Find-NetworkHosts.ps1       Ping スキャンでネットワーク上の機器を洗い出す（任意）
│  └ Invoke-RemoteInventory.ps1  WinRM で今すぐ一括収集したいとき用（補助）
├ Deploy-Startup.md              方式A: 起動時コピーでの導入手順（推奨）
└ Install-InventoryTask.md       方式B: GPO スケジュールタスクでの導入手順
```

## 手順

### 1. 収集する

次のどちらかの方式で、各端末から結果共有フォルダに CSV を集めます。

| 方式 | 内容 | 手順書 |
|---|---|---|
| **A: 起動時コピー（GPO 追加不要・推奨）** | 既存のスタートアップ用バッチに 2 行を追加する。端末を起動するたびに `startup\inventory` を `C:\inventory` へコピーして実行する | [Deploy-Startup.md](Deploy-Startup.md) |
| B: GPO のスケジュールタスク | GPO で全端末にタスクを登録し、毎日と起動 5 分後に自動で収集する | [Install-InventoryTask.md](Install-InventoryTask.md) |

**どちらの方式でも、全台展開の前に Windows 7 と Windows 10/11 の端末で 1 台ずつ試験してください。**

### 2. 一覧を作る（いつでも何度でも）

管理 PC に `admin` フォルダごとコピーして実行します（`-RawDir` は `run-inventory.bat` の `SHARE=` に設定した結果共有）。

```powershell
cd C:\tools\admin
.\Merge-Inventory.ps1 -RawDir \\nagasakinet.local\dfsroot\newton\inventory_result `
                      -LedgerFile C:\tools\台帳.csv `
                      -OutFile C:\tools\端末一覧.xlsx
```

台帳.csv は、担当者だけが読み書きできる場所に置いてください（結果共有と同じフォルダには置かないでください）。

スクリプトの実行が禁止されている場合は、`powershell -ExecutionPolicy Bypass -File .\Merge-Inventory.ps1 …` で実行してください。

### 3. 台帳.csv にタグNo・導入日を入力する

初回の実行で `台帳.csv` が作られます。Excel で開いて、**タグNo・導入日・設置場所・備考** の 4 列を入力し、
「CSV（コンマ区切り）」形式で上書き保存してください（UTF-8 / Shift-JIS のどちらで保存しても読めます）。

| 列 | 誰が書くか |
|---|---|
| タグNo / 導入日 / 設置場所 / 備考 | **担当者が手入力**（スクリプトは上書きしません） |
| 端末名 / シリアル番号 / MACアドレス / 初回検出日 / 最終検出日 | スクリプトが毎回更新 |

- 導入日は `2021/4/1` のような形式で入力します。
- 端末は **シリアル番号** で照合します（シリアルが無い機種は端末名で照合）。端末名を変えてもタグNoは引き継がれます。
- 新しい端末は自動で行が追加されます。廃棄した端末は行を削除してください（残しておくと「未収集」として出続けます）。
- 実行のたびに前回分を `台帳.csv.bak` に残します。**台帳.csv を Excel で開いたまま実行すると更新できません。**

もう一度 `Merge-Inventory.ps1` を実行すると、入力内容が一覧に反映されます。

### 4.（任意）ネットワーク上の未収集機器も一覧に入れる

```powershell
.\Find-NetworkHosts.ps1 -Subnet 192.168.10,192.168.11 -OutFile C:\tools\discovery.csv
.\Merge-Inventory.ps1 -RawDir … -LedgerFile … -OutFile … -Discovery C:\tools\discovery.csv
```

MAC アドレスは ARP で取るため、スキャンする PC と同じセグメントの機器しか取れません。

## 注意事項

- スクリプトは**読み取りのみ**で、端末の設定は変更しません。ただし端末へのスクリプト配布や GPO の追加は、電子カルテベンダーとの保守契約に従って事前に連絡・承認を取ってください。
- 情報が取れるのは、電源が入っている端末だけです。方式 A は**端末の起動時**に収集するため、再起動しない端末は収集されません（「収集済(30日以上前)」で分かります。その端末で `C:\inventory\run-inventory.bat` を管理者権限で実行すれば収集されます）。
- 「起動日」は Windows 8 以降の高速スタートアップが有効だと、シャットダウン→電源 ON では更新されません（再起動で更新されます）。
- 出力（xlsx・台帳・raw）にはライセンス情報が含まれます。保存場所のアクセス権に注意してください。
