# RankingQuagga
[Quagga](https://quagga.studio) APIと連携して成績情報を反映するPowerPointプラグインです。

## 機能
- ピリオド成績
- 総合成績

## ファイル
- srcフォルダ: RankingQuagga.pptmを解凍したもの
- imagesフォルダ: README用画像
- RankingQuagga.ppam: **プラグイン本体**
- RankingQuagga.pptm: 編集用

## インストール
1. Github上の[RankingQuagga.ppam](RankingQuagga.ppam)のページに飛ぶ
2. "Download raw file"をクリックしてプラグイン本体をダウンロード  
![download-raw-file](images/download-raw-file.png)  
これ以降はOfficeバージョンによって差異がある可能性アリ
3. "%APPDATA%\Microsoft\AddIns"にコピー
4. "ファイル"→"オプション"→"アドイン"→"管理"項目の"PowerPointアドイン"を選択し"設定"をクリック
5. "新規追加"より3.でコピーしたプラグイン本体を選択
6. "RankingQuagga"にチェックがついていることを確認し、設定ウィンドウを閉じる

## 使い方(共通)
1. リボンの"RankingQuagga"タブを選択
2. "設定"をクリック
3. イベントIDと秘密の暗号を入力  
\(サーバーは"テストサーバー"を選択する必要はない。テストサーバーについては＠＠＠フォルダを参照)
4. "保存"をクリック
5. 成績を反映させるオブジェクト名を下図のように編集(階層はこの通りでなくてよい、"番号"はランクの意)  
![object-name](images/object-name.png)

### ピリオド成績
1. 成績を反映させる予定のスライドを選択した状態で"ピリオド成績を反映"をクリック

### 総合成績
総合成績は2枚スライドが必要になる。下位を表示する用のスライドと、トップ10を表示する用のスライドである。  
1. 下位を表示する用のスライドを選択した状態で"総合成績を反映"をクリック

## Dependencies
- VBA-JSON (MIT License): https://github.com/VBA-tools/VBA-JSON
- VBA-Dictionary (MIT License): https://github.com/VBA-tools/VBA-Dictionary
