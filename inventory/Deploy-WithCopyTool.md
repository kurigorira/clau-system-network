# 既存のコピーツールで収集する手順（GPO 不要）

院内にすでにある「サーバーから各端末へファイルをコピーして実行する仕組み」を使って、収集スクリプトを配って実行します。
各端末は結果を **サーバーの共有フォルダへ直接** 書き込みます。GPO の追加は不要です。

```
[サーバー] 既存のコピーツール
   ├─ コピー: Get-DeviceInventory.ps1 と run-inventory.bat → 端末の C:\ProgramData\HospitalInventory\
   └─ 実行  : C:\ProgramData\HospitalInventory\run-inventory.bat
[端末]     → \\fs01\inventory$\raw\<端末名>.csv を書き込む（結果ログは端末内の last-run.log）
[管理PC]   Merge-Inventory.ps1 → 端末一覧.xlsx ＋ 台帳.csv
```

## 1. bat の出力先を設定する

`run-inventory.bat` を開き、13 行目の出力先を自分の共有フォルダに書き換えます。

```bat
set SHARE=\\fs01\inventory$\raw
```

書き換えずに、実行時の引数で指定することもできます（`run-inventory.bat \\fs01\inventory$\raw`）。

## 2. コピーツールに登録する内容

| 項目 | 設定値 |
|---|---|
| コピーするファイル | `Get-DeviceInventory.ps1`、`run-inventory.bat`（**必ず 2 つセットで、同じフォルダに置く**） |
| コピー先 | `C:\ProgramData\HospitalInventory\`（別の場所でも可） |
| 実行するコマンド | `C:\ProgramData\HospitalInventory\run-inventory.bat` |
| 成功の判定 | 終了コード `0` = 共有への書き込み成功、`1` = 書き込み失敗 |

1 回の実行は通常 10〜30 秒で終わります。Windows 7 では、Office ライセンスの照会に 1 分前後かかることがあります。

## 3. 実行アカウントと共有フォルダの権限

**実行したアカウントで** 共有フォルダ（`raw`）へ書き込みます。コピーツールがどのアカウントで実行するかによって、権限が必要な相手が変わります。

| コピーツールの実行アカウント | 共有フォルダ `raw` に書き込み権限が必要な相手 |
|---|---|
| SYSTEM（ローカルシステム） | 端末のコンピューターアカウント（例: **Domain Computers** グループ） |
| ドメインの管理用アカウント | そのアカウント |
| ログオン中のユーザー | そのユーザー（例: **Domain Users** グループ） |

- 同じ端末が 2 回目以降に実行すると、自分の `<端末名>.csv` を削除してから書き直します。そのため、**作成だけでなく削除（変更）の権限** も必要です。
- 結果にはOffice ライセンス情報が入ります。`raw` の **読み取り** は情報システム担当だけに絞ってください。

## 4. 試験（全台展開の前に必ず）

1. Windows 7 の端末と Windows 10/11 の端末を 1 台ずつ選び、コピーツールで配布・実行する
2. 共有の `raw` に `<端末名>.csv` ができているか確認する
3. CSV を開き、次の点を確認する
   - IPAddress / MACAddress に電子カルテ網の NIC が入っているか
   - OfficeProduct / OfficeLicenseKey が入っているか（Office が無い端末は空欄）
   - EstimatedStartDate（推定導入日）が妥当な日付か
4. CSV ができていない場合は、端末の `C:\ProgramData\HospitalInventory\last-run.log` を見る

```
[2026/10/10 10:00:01.23] start computer=EMR-PC001 user=SYSTEM output=\\fs01\inventory$\raw
OK: \\fs01\inventory$\raw\EMR-PC001.csv
[2026/10/10 10:00:18.45] end exit=0
```

| ログの内容 | 原因と対処 |
|---|---|
| `ERROR: could not write ... Access is denied` / `アクセスが拒否されました` | 実行アカウントに共有への書き込み権限がない → 3. の表を確認する |
| `ERROR: could not write ... network path was not found` / `ネットワーク パスが見つかりません` | 共有パスの誤り、またはサーバーに到達できない |
| `start` の行しかない | PowerShell を起動できていない、または ps1 が同じフォルダに無い |
| ログ自体が無い | bat が実行されていない（コピーツール側の実行設定を確認する） |

## 5. 運用

1. 月 1 回など定期的に、コピーツールで配布・実行する（ファイルは上書きで問題ありません）
2. 管理 PC で `Merge-Inventory.ps1` を実行し、`端末一覧.xlsx` を作る（[README](README.md) の「一覧を作る」を参照）
3. 「未収集」シートや「収集済(30日以上前)」の端末は、電源が切れていた可能性があります。次回の実行で拾えるか確認してください

> 配布・実行できるのは、その時点で電源が入っている端末だけです。
> 夜間や休日に電源を落とす端末がある部署は、日中に実行してください。
