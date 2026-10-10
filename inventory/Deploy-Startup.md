# ログオン時コピー方式の導入手順（方式A・推奨）

各端末で**ユーザーがログオンしたとき**に動いている「サーバーからローカルへコピーするバッチ」に 2 行追加し、
配布フォルダを `C:\inventory` へコピーして収集スクリプトを実行します。GPO の追加は不要です。

```
サーバー  \\nagasakinet.local\dfsroot\newton\startup\inventory\
            ├ Get-DeviceInventory.ps1
            └ run-inventory.bat        ← SHARE= に結果の保存先を設定
              │  既存のログオン用バッチ（追加した 2 行）※ログオンしたユーザーの権限で動く
              ▼  ① robocopy でコピー  ② バックグラウンドで実行
端末      C:\inventory\                ← Windows 7 は事前準備が必要（手順 0）
            ├ Get-DeviceInventory.ps1
            ├ run-inventory.bat
            └ last-run.log             ← 実行結果のログ
              │
              ▼
結果共有  \\<結果の保存先>\<端末名>_<ユーザー名>.csv   ← 1 人 1 台につき 1 ファイル
              │
              ▼
管理PC    admin\Merge-Inventory.ps1 → 端末一覧.xlsx ＋ 台帳.csv   （手順 7）
```

## 0. Windows 7 端末の事前準備（1 回だけ）

Windows 7 の端末では、一般ユーザーが `C:\` に `inventory` フォルダを作れないため、robocopy が「アクセス拒否」になります。
管理者が一度だけ `C:\inventory` を作り、**Users グループに変更権限**を付けておけば、以後は一般ユーザーでもコピーできます。
（Windows 10/11 では不要です。実行しても問題はありません。）

### まとめて行う（管理用 PC から）

ドメイン管理者で PowerShell を開き、`admin` フォルダで実行します。管理共有（`\\端末\C$`）を通して各端末に作ります。

```powershell
# AD に登録されている Windows 7 の端末すべて
.\Prepare-InventoryFolder.ps1 -FromActiveDirectory

# 端末名を書いたテキストファイル（1 行 1 台）を使う場合
.\Prepare-InventoryFolder.ps1 -ComputerListFile .\win7.txt
```

結果は `prepare-result.csv` に出ます。

| Status | 意味 |
|---|---|
| OK | 準備完了 |
| Offline | 電源が切れていた。あとでもう一度実行する（何度実行しても問題ありません） |
| CreateFailed | `\\端末\C$` に入れない（管理共有が無効、またはファイアウォール）。下の「1 台ずつ行う」で対応する |
| IcaclsFailed | 権限の設定に失敗した（Detail 列にエラー内容） |

### 1 台ずつ行う

`terminal\prepare-inventory-folder.bat` を端末に持っていき、右クリック →「**管理者として実行**」します。

## 1. 結果の保存先（共有フォルダ）を用意する

各端末の CSV を集める共有フォルダを用意します。配布元の `startup\inventory` とは**別のフォルダ**にしてください。
（配布元と同じ場所にすると、CSV がほかの端末へコピーされてしまいます。）

収集は**ログオンしたユーザーの権限**で動くので、一般ユーザーが書き込めるようにします。

| 対象 | 権限 |
|---|---|
| **Domain Users** | このフォルダーのみ：「フォルダーの一覧/データの読み取り」「属性の読み取り」「ファイルの作成/データの書き込み」 |
| **CREATOR OWNER** | ファイルのみ：「変更」（自分が作った CSV だけを書き直せる） |
| 情報システム担当 | フルコントロール |

- 一般ユーザーは、**ほかの人の CSV の中身は読めません**（ファイル名の一覧は見えます）。CSV には Office ライセンス情報が入るためです。
- 1 台の端末を何人も使うので、ファイルは `<端末名>_<ユーザー名>.csv` と**人ごとに分かれます**。
  1 つのファイルを全員で上書きする作りにすると、ほかの人が作ったファイルを書き換える権限が必要になってしまうためです。
  集計のときに、端末ごとに最新の 1 件だけを使うので、一覧が重複することはありません（手順 7）。
- 同じ人が 20 時間以内にもう一度ログオンしたときは、収集せずにすぐ終わります（ログに `SKIP` と出ます）。

## 2. サーバーに配布ファイルを置く

### パスを確認する

エクスプローラーの表示が `newton (\\NAGASAKINET.local\dfsroot)` の場合、実際のパスは `\\nagasakinet.local\dfsroot\newton` です。
コマンドプロンプトで次を実行し、**ファイル一覧が出る方**のパスを使ってください（この手順書は `newton` で書いています）。

```bat
dir \\nagasakinet.local\dfsroot\newton\startup\inventory
dir \\nagasakinet.local\dfsroot\newtons\startup\inventory
```

### 置くファイル

リポジトリの `terminal\` フォルダの**中にある 2 ファイルだけ**を、`startup\inventory\` の**直下**に置きます。

```
正しい置き方                              よくある間違い（動きません）
startup\inventory\                        startup\inventory\
  ├ Get-DeviceInventory.ps1                 ├ admin\
  └ run-inventory.bat                       ├ terminal\
                                            │  ├ Get-DeviceInventory.ps1
                                            │  └ run-inventory.bat
                                            ├ README.md など
                                            └ network.bat（追記したログオン用バッチ）
```

- `startup\inventory\` の中身はすべて各端末の `C:\inventory` にコピーされます。
- `admin\`（集計用）、`.md`（手順書）、`startup-snippet.bat`（見本）、`check-inventory.bat` / `prepare-inventory-folder.bat`、ログオン用バッチ本体は**ここに置かないでください**。

## 3. run-inventory.bat に結果の保存先を設定する

サーバー上の `run-inventory.bat` をメモ帳で開き、次の行を手順 1 の共有フォルダに書き換えます。

```bat
rem ===== Result folder (CHANGE THIS to the shared folder for results) =====
set SHARE=\\nagasakinet.local\dfsroot\CHANGE_ME
```

例: `set SHARE=\\nagasakinet.local\dfsroot\newton\inventory_result`

`CHANGE_ME` のまま配布すると、端末は何もせずに終了し、ログに `ERROR: SHARE is not configured` と出ます。

## 4. 既存のログオン用バッチに 2 行を追加する

既存のコピー用バッチの**最後**に、`terminal\startup-snippet.bat` の 2 行を追加します。

```bat
robocopy "\\nagasakinet.local\dfsroot\newton\startup\inventory" "C:\inventory" /E /R:1 /W:1 /XF last-run.log /NP /NFL /NDL /LOG:C:\inventory-copy.log
if %ERRORLEVEL% lss 8 start "" /b cmd /c "C:\inventory\run-inventory.bat"
```

- 1 行目で `startup\inventory` を `C:\inventory` にコピーします（変更があったファイルだけ上書き）。結果は `C:\inventory-copy.log` に残ります。
- 2 行目で、コピーが成功したとき（終了コード 8 未満）だけ、収集をバックグラウンドで開始します。ログオンを待たせません。

## 5. 試験（全台展開の前に必ず）

### 診断用の bat で確認する

試験端末（**Windows 7 と Windows 10/11 を 1 台ずつ**）に `terminal\check-inventory.bat` をコピーし、
**一般ユーザーでログオンして**ダブルクリックします。次の 6 段階を順に調べ、`[OK]` / `[NG]` と対処方法を表示します。

| 段階 | 調べること |
|---|---|
| [1] | 配布元 `startup\inventory` に届くか（`newton` / `newtons` の違いもここで分かる） |
| [2] | 2 ファイルが配布元の直下にあるか |
| [3] | `C:\inventory` へのコピー（Windows 7 で失敗したら手順 0 を案内） |
| [4] | `run-inventory.bat` の `SHARE=` が設定済みか、その共有に届くか |
| [5] | 収集を実行し、終了コードと `last-run.log` を表示（診断では 20 時間の間引きをしない） |
| [6] | 結果共有に `<端末名>_<ユーザー名>.csv` ができたか |

配布元のパスが違う場合は、`check-inventory.bat` の先頭の `set SRC=` を書き換えてください。最後に `All steps OK.` と出れば完了です。

### ログオン時の動作を確認する

1. 試験端末のログオン用バッチにだけ、手順 4 の 2 行を追加する
2. 一般ユーザーでログオンし、2〜3 分待つ
3. 結果共有に `<端末名>_<ユーザー名>.csv` ができているか確認する
4. CSV を開き、IPAddress / MACAddress、OfficeProduct / OfficeLicenseKey、EstimatedStartDate（推定導入日）が妥当か確認する
5. できていなければ、端末の `C:\inventory\last-run.log` と `C:\inventory-copy.log` を確認する

```
[2026/10/10  8:31:02.15] start computer=EMR-PC001 user=nurse01 output=\\nagasakinet.local\dfsroot\newton\inventory_result
OK: \\nagasakinet.local\dfsroot\newton\inventory_result\EMR-PC001_nurse01.csv
[2026/10/10  8:31:20.48] end exit=0
```

| 症状・ログの内容 | 原因と対処 |
|---|---|
| `C:\inventory` が無い／`inventory-copy.log` に「アクセスが拒否されました」 | Windows 7 で事前準備をしていない（手順 0） |
| `C:\inventory\terminal\` がある | 配布元に `terminal` フォルダごと置いている（手順 2） |
| `last-run.log` が無い | 2 行目が実行されていない（追記した位置・内容を確認する） |
| `ERROR: SHARE is not configured` | 手順 3 の書き換えを忘れている |
| `ERROR: cannot reach ...` | 結果共有のパスの誤り、または Domain Users に「フォルダーの一覧」権限がない |
| `ERROR: could not write ... アクセスが拒否されました` | Domain Users に「ファイルの作成」、CREATOR OWNER に「変更」が無い（手順 1） |
| `SKIP: ... written 3.2 hours ago` | 正常。同じ人が 20 時間以内にすでに収集済み |
| `start` の行だけで `end` が無い | 下の「バックグラウンド実行が止まる場合」を参照 |

### バックグラウンド実行が止まる場合

環境によっては、ログオン用バッチが終わったときに、バックグラウンドの処理も止められることがあります。
その場合は 2 行目を次のように変えて、終わるまで待つ実行にしてください（ログオンが最大 1 分ほど遅くなります）。

```bat
if %ERRORLEVEL% lss 8 call "C:\inventory\run-inventory.bat"
```

## 6. 運用

- ユーザーがログオンするたびに収集されます（同じ人・同じ端末は 20 時間に 1 回まで）。
- **誰もログオンしない端末は収集されません。** 一覧で「収集済(30日以上前)」（黄色）や「未収集」（赤）に出ます。
- スクリプトを更新するときは、サーバーの `startup\inventory` のファイルを置き換えるだけです。次のログオン時に各端末へ反映されます。
- 新しく Windows 7 の端末を追加したときは、その端末に手順 0 を行ってください。

## 7. 結果の CSV をまとめる（Excel 一覧の作成）

結果共有にたまった CSV は、管理用 PC で `admin\Merge-Inventory.ps1` を実行すると、1 つの Excel ファイルにまとまります。
CSV を手で開いたり、コピーしてつなげたりする必要はありません。好きなときに何度でも実行できます。

```powershell
cd C:\tools\admin
.\Merge-Inventory.ps1 -RawDir \\nagasakinet.local\dfsroot\newton\inventory_result `
                      -LedgerFile C:\tools\台帳.csv `
                      -OutFile C:\tools\端末一覧.xlsx
```

| 引数 | 指定するもの |
|---|---|
| `-RawDir` | 結果共有（`run-inventory.bat` の `SHARE=` と同じパス） |
| `-LedgerFile` | 台帳.csv の保存先（初回は自動で作られます。担当者だけが読み書きできる場所に置く） |
| `-OutFile` | 作成する Excel ファイル |

このスクリプトは次のことを自動で行います。

1. 結果共有の **すべての CSV** を読み込む
2. 同じ端末の CSV が複数ある場合（複数のユーザーがログオンした端末）は、**最後に収集された 1 件だけ**を採用する
   （シリアル番号で同じ端末と判断します。シリアルが無い機種は端末名で判断します）
3. 台帳.csv と突き合わせ、手入力したタグNo・導入日・設置場所を反映する（新しい端末は台帳に自動で追加）
4. `端末一覧.xlsx` を作る。シートは「端末一覧」（1 台 1 行）、「未収集」、「集計」（OS 別・Office 別・稼働月数別の台数）の 3 枚

台帳.csv の入力方法など詳しくは [README](README.md) の「一覧を作る」「台帳.csv にタグNo・導入日を入力する」を参照してください。

## 8. 停止・撤去

ログオン用バッチから 2 行を削除します。端末に残った `C:\inventory` は、そのままにしても動作に影響はありません。
