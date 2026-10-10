# 起動時コピー方式の導入手順（方式A・推奨）

各端末が起動するときに動いている「サーバーからローカルへコピーするバッチ」（ログオン前・SYSTEM 権限）に 2 行追加し、
`inventory` フォルダを `C:\inventory` へコピーして収集スクリプトを実行します。GPO の追加は不要です。

```
サーバー  \\nagasakinet.local\dfsroot\newton\startup\inventory\
            ├ Get-DeviceInventory.ps1
            └ run-inventory.bat        ← SHARE= に結果の保存先を設定
              │  既存のスタートアップ用バッチ（追加した 2 行）
              ▼  ① robocopy でコピー  ② バックグラウンドで実行
端末      C:\inventory\
            ├ Get-DeviceInventory.ps1
            ├ run-inventory.bat
            └ last-run.log             ← 実行結果のログ
              │
              ▼
結果共有  \\<結果の保存先>\<端末名>.csv
              │
              ▼
管理PC    admin\Merge-Inventory.ps1 → 端末一覧.xlsx ＋ 台帳.csv
```

## 1. 結果の保存先（共有フォルダ）を用意する

各端末の CSV を集める共有フォルダを用意します。配布元の `startup\inventory` とは**別のフォルダ**にしてください。
配布元と同じ場所にすると、端末の CSV がほかの端末へコピーされてしまいます。

コピー用バッチは **SYSTEM 権限**で動くため、共有へは**端末のコンピューターアカウント**として書き込みます。

| 対象 | 権限 |
|---|---|
| **Domain Computers** | 「ファイルの作成/データの書き込み」（このフォルダーのみ）＋ CREATOR OWNER に「変更」（ファイルのみ） |
| 情報システム担当 | フルコントロール（読み取りは担当者だけに絞る） |

- 2 回目以降、各端末は自分の `<端末名>.csv` を削除してから書き直します。そのため、自分が作ったファイルを変更・削除できる権限（CREATOR OWNER: 変更）が必要です。
- CSV には Office ライセンス情報が入ります。一般利用者が読めないようにしてください。

## 2. サーバーに inventory フォルダを置く

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
                                            └ network.bat（追記したスタートアップ用バッチ）
```

- `startup\inventory\` の中身はすべて各端末の `C:\inventory` にコピーされます。`terminal\` の下に置くと、端末側は `C:\inventory\terminal\run-inventory.bat` になってしまい、実行されません。
- `admin\`（集計用）、`.md`（手順書）、`startup-snippet.bat`（見本）、スタートアップ用バッチ本体は、**ここに置かないでください**。管理 PC など別の場所に保管します。
- `check-inventory.bat`（診断用）もここには置かず、必要なときに試験端末へ持っていって使います（手順 5）。

## 3. run-inventory.bat に結果の保存先を設定する

サーバー上の `run-inventory.bat` をメモ帳で開き、次の行を手順 1 の共有フォルダに書き換えます。

```bat
rem ===== Result folder (CHANGE THIS to the shared folder for results) =====
set SHARE=\\nagasakinet.local\dfsroot\CHANGE_ME
```

例: `set SHARE=\\nagasakinet.local\dfsroot\newton\inventory_result`

`CHANGE_ME` のまま配布すると、端末は何もせずに終了し、ログに `ERROR: SHARE is not configured` と出ます。

## 4. 既存のスタートアップ用バッチに 2 行を追加する

既存のコピー用バッチの**最後**に、`terminal\startup-snippet.bat` の 2 行を追加します。

```bat
robocopy "\\nagasakinet.local\dfsroot\newton\startup\inventory" "C:\inventory" /E /R:1 /W:1 /XF last-run.log /NP /NFL /NDL /LOG:C:\inventory-copy.log
if %ERRORLEVEL% lss 8 start "" /b cmd /c "C:\inventory\run-inventory.bat"
```

- 1 行目で `startup\inventory` を `C:\inventory` にコピーします（フォルダが無ければ作られます。変更があったファイルだけ上書きします）。
  コピーの結果は `C:\inventory-copy.log` に残ります。
- 2 行目で、コピーが成功したとき（robocopy の終了コードが 8 未満）だけ、収集をバックグラウンドで開始します。
  Windows 7 ではライセンスの照会に 1 分前後かかることがあるので、起動を待たせないようにバックグラウンドで動かします。
- 既存のバッチで変数 `%ERRORLEVEL%` を別の目的に使っている場合は、2 行をまとめて最後に置けば影響しません。

## 5. 試験（全台展開の前に必ず）

### まず診断用の bat で確認する

`terminal\check-inventory.bat` を試験端末にコピーし、ダブルクリックで実行します。
次の 6 段階を順に調べ、`[OK]` / `[NG]` と対処方法を画面に表示して、問題があった段階で止まります。

| 段階 | 調べること |
|---|---|
| [1] | 配布元 `startup\inventory` に届くか（パスの `newton` / `newtons` 違いもここで分かる） |
| [2] | 2 ファイルが配布元の直下にあるか（`terminal\` の下に入っている場合はそう表示する） |
| [3] | `C:\inventory` へのコピー |
| [4] | `run-inventory.bat` の `SHARE=` が設定済みか、その共有に届くか |
| [5] | 収集を実行し、終了コードと `last-run.log` を表示 |
| [6] | 結果共有に `<端末名>.csv` ができたか |

配布元のパスが違う場合は、`check-inventory.bat` の先頭の `set SRC=` を書き換えてください。
最後まで `All steps OK.` と出れば準備は完了です。止まった場合は、その画面を見て対処してください。

> 診断用の bat は**ダブルクリックしたユーザーの権限**で動きます。起動時のバッチは **SYSTEM（コンピューターアカウント）** で動くので、
> 診断ではうまくいくのに起動時は書き込めない場合は、結果共有に Domain Computers の書き込み権限があるか確認してください。

### 起動時の動作を確認する

1. 手順 2〜3 を済ませる
2. 試験用の端末（**Windows 7 と Windows 10/11 を 1 台ずつ**）のスタートアップ用バッチにだけ、手順 4 の 2 行を追加する
   （難しければ、その端末で管理者のコマンドプロンプトを開き、2 行を手で実行してもかまいません）
3. 端末を再起動し、2〜3 分待つ
4. 結果共有に `<端末名>.csv` ができているか確認する
5. CSV を開き、次の点を確認する
   - IPAddress / MACAddress に電子カルテ網の NIC が入っているか
   - OfficeProduct / OfficeLicenseKey が入っているか（Office が無い端末は空欄）
   - EstimatedStartDate（推定導入日）が妥当な日付か
6. できていなければ、端末の `C:\inventory\last-run.log` を確認する

```
[2026/10/10  8:31:02.15] start computer=EMR-PC001 user=EMR-PC001$ output=\\nagasakinet.local\dfsroot\newton\inventory_result
OK: \\nagasakinet.local\dfsroot\newton\inventory_result\EMR-PC001.csv
[2026/10/10  8:31:20.48] end exit=0
```

| ログの内容 | 原因と対処 |
|---|---|
| `C:\inventory` が無い | robocopy が失敗している。`C:\inventory-copy.log` を見る（配布元のパス、またはコピー用バッチへの追記を確認する） |
| `C:\inventory\terminal\` がある | 配布元に `terminal` フォルダごと置いている（手順 2 の「よくある間違い」） |
| `last-run.log` が無い | 2 行目が実行されていない（追記した位置・内容を確認する） |
| `ERROR: SHARE is not configured` | 手順 3 の書き換えを忘れている |
| `ERROR: cannot reach ...` | 結果共有のパスの誤り、または起動直後にネットワークへつながらない（30 秒待っても届かなかった） |
| `ERROR: could not write ... アクセスが拒否されました` | Domain Computers に書き込み権限がない（手順 1） |
| `start` の行だけで `end` が無い | 収集の途中で止められた可能性がある。下の「バックグラウンド実行が止まる場合」を参照 |

### バックグラウンド実行が止まる場合

環境によっては、スタートアップ用バッチが終わったときに、バックグラウンドで起動した処理も止められることがあります。
その場合は、2 行目を次のように変えて、終わるまで待つ実行にしてください（起動が最大 1 分ほど遅くなります）。

```bat
if %ERRORLEVEL% lss 8 call "C:\inventory\run-inventory.bat"
```

## 6. 運用

- 端末を起動するたびに、その端末の CSV が最新に置き換わります。
- 管理 PC で、好きなタイミングで `admin\Merge-Inventory.ps1` を実行し、一覧を作ります（[README](README.md) の「一覧を作る」を参照）。
- **再起動しない端末は収集されません。** 一覧で「収集済(30日以上前)」と黄色で表示されたり、「未収集」に出たりします。
  そのような端末では、管理者権限で `C:\inventory\run-inventory.bat` を一度実行すれば収集されます。
- スクリプトを更新するときは、サーバーの `startup\inventory` のファイルを置き換えるだけです。各端末の次の起動時に反映されます。

## 7. 停止・撤去

スタートアップ用バッチから 2 行を削除します。端末に残った `C:\inventory` は、そのままにしても動作に影響はありません。
