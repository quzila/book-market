# Apps Script セットアップ

## 1. Google スプレッドシートを作成
1. 新規の Google スプレッドシートを1つ作成。
2. `拡張機能 > Apps Script` を開く。

## 2. スクリプトを配置
1. `/Users/tatsuru/Desktop/book/apps-script/Code.gs` の内容を `Code.gs` に貼り付ける。
2. 保存する。

## 3. 初期化
1. Web アプリとしてデプロイ後、以下へアクセスして初期化を確認する。
   - `.../exec?action=initSheets`
2. `listings` / `interests` シートが作成される。

## 4. Web アプリとしてデプロイ
1. `デプロイ > 新しいデプロイ`。
2. 種類は `ウェブアプリ`。
3. 実行ユーザー: `自分`。
4. アクセス権: `全員`（または運用方針に合わせる）。
5. 発行された URL を確認する。

## 5. フロントとの接続
- フロントは以下URLを固定で使用する実装。
- `https://script.google.com/macros/s/AKfycbxEkgzuv0id6ahjSORd87458qTVSlFpVAw4Yixea3PCa3c90DcZN3iqNF4LmIkLnarF/exec`
- 別URLを使う場合のみ `/Users/tatsuru/Desktop/book/config.js` で `APPS_SCRIPT_BASE_URL` を上書き。

## 6. シート仕様
### `listings`
`created_at, listing_id, item_type, category, title, description, image_urls_json, jan, subject_tags_json, author, publisher, published_date, seller_name, seller_id, status`

### `interests`
`created_at, interest_id, listing_id, viewer_id, viewer_name`

## 7. API
### POST `action=loginByRoom`
request:
```json
{ "roomNumber": "105" }
```

### GET `action=listListings&viewerId=...&category=...&subject=...&q=...`

### GET `action=listListingsV2&viewerId=...&limit=80&cursor=0`
- 初回ロード向けの軽量一覧API（`wantedBy` を含まない）。
- サーバー側の一覧スナップショット JSON からページ単位で返す。
- レスポンスに `version` を含む（クライアント側のキャッシュ判定用）。
- `nextCursor` が空文字の場合、次ページなし。

### GET `action=getListingsVersion`
- 一覧スナップショットの `version` と `totalCount` を返す軽量API。
- クライアントは `version` が同じ場合に一覧再取得をスキップできる。

### GET `action=getListingDetailV2&listingId=...&viewerId=...`
- 商品詳細を遅延取得するAPI。

### GET `action=listWishersV2&listingId=...&viewerId=...`
- 希望者一覧を遅延取得するAPI。
- 未ログイン (`viewerId` なし) の場合は `wantedBy` は空配列、`wantedCount` のみ返す。

### POST `action=uploadImage`
request:
```json
{ "fileName": "photo.jpg", "mimeType": "image/jpeg", "dataUrl": "data:image/jpeg;base64,..." }
```

### POST `action=addListing`
request:
```json
{
  "sellerId": "resident_105",
  "sellerName": "坪田　晴琉",
  "listing": {
    "itemType": "book",
    "category": "書籍",
    "title": "線形代数",
    "description": "書き込み少なめ",
    "imageUrls": ["https://..."],
    "jan": "9784410105784",
    "subjectTags": ["数学"],
    "author": "著者名",
    "publisher": "出版社",
    "publishedDate": "2023-04-01"
  }
}
```

### POST `action=addListingsBatch`
request:
```json
{
  "sellerId": "resident_105",
  "sellerName": "坪田　晴琉",
  "listings": [
    {
      "itemType": "goods",
      "category": "生活用品",
      "title": "延長コード",
      "description": "使用1年",
      "imageUrls": ["https://..."]
    },
    {
      "itemType": "book",
      "category": "書籍",
      "title": "力学",
      "description": "",
      "imageUrls": [],
      "jan": "9784410105784",
      "subjectTags": ["物理"]
    }
  ]
}
```

### POST `action=updateListing`
request:
```json
{
  "sellerId": "resident_105",
  "listingId": "listing_xxx",
  "listing": {
    "itemType": "book",
    "category": "書籍",
    "title": "線形代数（第2版）",
    "description": "状態更新済み",
    "imageUrls": ["https://..."],
    "jan": "9784410105784",
    "subjectTags": ["数学"],
    "author": "著者名",
    "publisher": "出版社",
    "publishedDate": "2023-04-01"
  }
}
```
- 自分の有効な出品 (`status=AVAILABLE`) のみ更新可能。

### POST `action=addInterest`
request:
```json
{ "viewerId": "resident_105", "viewerName": "坪田　晴琉", "listingId": "listing_xxx" }
```

### POST `action=removeInterest`
request:
```json
{ "viewerId": "resident_105", "listingId": "listing_xxx" }
```

### POST `action=cancelListing`
request:
```json
{ "sellerId": "resident_105", "listingId": "listing_xxx" }
```

### GET `action=listMyPage&viewerId=resident_105`

## 8. 権限について
- `uploadImage` は Google Drive にファイルを保存するため、初回実行時に Drive 権限を承認する。

## 9. GitHub Pages 公開
1. `/Users/tatsuru/Desktop/book` を GitHub に push。
2. `Settings > Pages` で公開ブランチを設定。
3. 公開URLで Home / Plus / My の動作を確認。
