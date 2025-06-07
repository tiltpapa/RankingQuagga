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
3. "%APPDATA%\Microsoft\AddIns"にコピー
これ以降はOfficeバージョンによって差異がある可能性アリ
4. "ファイル"→"オプション"→"アドイン"→"管理"項目の"PowerPointアドイン"を選択し"設定"をクリック
5. "新規追加"より3.でコピーしたプラグイン本体を選択する
6. "RankingQuagga"にチェックがついていることを確認し、設定ウィンドウを閉じる

## Dependencies
- VBA-JSON (MIT License): https://github.com/VBA-tools/VBA-JSON
- VBA-Dictionary (MIT License): https://github.com/VBA-tools/VBA-Dictionary
