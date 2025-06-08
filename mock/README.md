# ランキングAPIモック

テスト目的のCloudflare Workersを使用したAPIモックサービスです。

## 機能

以下の2つのAPIエンドポイントを提供します：

1. 特定のピリオドの成績取得: `/programs/{event_id}/result/period`
2. 暫定総合成績取得: `/programs/{event_id}/result/total?top={number}`

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

### 特定のピリオドの成績
- エンドポイント: `/programs/{event_id}/result/period`
- 15件の成績情報を返します
- ポイントは0-10のランダムな整数
- タイムは5.000-99.000のランダムな小数

### 暫定総合成績
- エンドポイント: `/programs/{event_id}/result/total`
- パラメータ:
  - `top`: 返す結果の数（デフォルト: 35）
- ポイントは1-50のランダムな整数
- タイムは5.000-499.000のランダムな小数

## 注意点

- テスト目的のため、`event_id`は実際には使用されません
- 名前はfaker-jsを使用してランダムに生成されます
- 10%の確率で名前の後ろに括弧書きでグループ名が付きます 