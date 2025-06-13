# Quagga APIモック

テスト目的のCloudflare Workersを使用したAPIモックサービスです。  
`https://ranking-quagga.tiltpapa.workers.dev`にてデプロイしています。

## 機能

以下の3つのAPIエンドポイントを提供します：

1. 直近に集計された問題情報: `/programs/{event_id}/questions/last_aggregate`
2. 特定のピリオドの成績取得: `/programs/{event_id}/result/period`
3. 暫定総合成績取得: `/programs/{event_id}/result/total?top={number}`

## セットアップ方法

1. 必要なパッケージをインストールします：

```bash
npm install
```

2. ローカルで開発サーバーを起動：

```bash
npm run dev
```

3. Cloudflare Workersにデプロイ：

```bash
npm run deploy
```

## APIの詳細

### 直近に集計された問題情報
- エンドポイント: `/programs/{event_id}/questions/last_aggregate`
- questionは固定です
- 15件の回答情報を返します
- 6-15人が正解します

### 特定のピリオドの成績
- エンドポイント: `/programs/{event_id}/result/period`
- 15件の成績情報を返します

### 暫定総合成績
- エンドポイント: `/programs/{event_id}/result/total`
- パラメータ:
  - `top`: 返す成績情報の数（本家には無いため省略可能、デフォルト: 35）

## 注意点
- テスト目的のため、`event_id`は実際には使用されません
- 名前はfakerを使用してランダムに生成されます
- 10%の確率で名前の後ろに括弧書きでグループ名が付きます
- テスト目的のため、パラメータを削っています
