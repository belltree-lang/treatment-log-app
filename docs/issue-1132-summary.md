# Issue #1132 / Issue #8 対応サマリー

## 変更点
- PDF生成の入口で prepared payload をシートから読み込み、PDF生成中は runtime cache（in-memory）で参照する導線を追加。
- PDF生成に必要な過去月を洗い出して一括ロードし、PDF生成フェーズではチャンク化キャッシュの merge を発生させない構成に変更。
- monthCache 参照時に runtime prepared payload を優先して使い、明細計算や合算の参照元を runtime cache に統一。

## 変更理由
- PDF生成中の `loadBillingCachePayload_ merged chunked cache` ログを排除し、チャンク化キャッシュの結合処理を避けるため。
- PDF生成の性能改善（読み込み回数とI/O削減）を狙うため。

## 影響範囲
- PDF生成系の入口関数（対象月・過去月の prepared payload 読み込みと runtime cache 参照）。
- 永続キャッシュ（CacheService / Drive）や prepared payload スキーマには変更なし。
- 請求書/領収書の金額計算、合算、表示仕様は変更なし。
