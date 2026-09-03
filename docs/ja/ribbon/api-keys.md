# API キー管理

**場所**: リボン `EGTools` タブ → `EGTools` グループ → スプリットボタンのドロップダウンメニュー

公共 API 関数・リボン機能に使用する API キーを登録/削除するダイアログです。

リストにはサービスが **3 つ**あります。

| サービス | 用途 | 関連機能 |
|---|---|---|
| juso | 道路名住所の検索 | [SEARCHADDRESS](../functions/korea/SEARCHADDRESS.md) |
| data.go.kr | 祝日・事業者登録状態 | [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md) · [カレンダー/予定表](calendar.md#祝日の案内) · [BRNSTATUS](../functions/korea/BRNSTATUS.md) |
| vworld | 地図/座標 | [GEOSEARCH](../functions/korea/GEOSEARCH.md) · [GEOCODER](../functions/korea/GEOCODER.md) · [GEOADDRESS](../functions/korea/GEOADDRESS.md) · [ADDRESSMAP](../functions/korea/ADDRESSMAP.md) |

- サービスを選択すると、キーの登録状況は**マスク（●●●●●●）のみ**で表示されます。
  保存されたキーの原文は再表示されません。
- キーは**ユーザーアカウントに保存**され、**ブックや数式には残りません**。
- 関数の**最後の引数としてキーを直接渡す**従来の方式とも共存します
  — 引数で渡したキーは自動的に同じ保存先へ保存され、以後は省略できます。
- **キーは EGTools++ と一緒には配布されません。**各サービス（juso.go.kr、
  data.go.kr、vworld.kr）でユーザー自身が発行を受ける必要があります。

## data.go.kr のキーは 1 つです

公共データポータル（data.go.kr）は**アカウントごとに認証キーを 1 つだけ**発行し、活用申請した
すべてのサービスでそのキーを共通に使います。そのため EGTools++ もサービスごとに分けず、
`data.go.kr` の 1 枠で管理します。

- この 1 つのキーを [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md)（祝日等の特日情報）、
  [カレンダー/予定表](calendar.md#祝日の案内)の臨時祝日の反映、
  [BRNSTATUS](../functions/korea/BRNSTATUS.md)（国税庁の事業者登録状態）が**すべて共有**します。
- 旧バージョンで**事業者登録用に別途登録しておいたキー**は、最初の使用時に
  統合された枠へ**自動的に移行**されます — ユーザーが登録し直したり整理したりする
  必要はありません。
- ポータルに表示されるエンコード（Encoding）/デコード（Decoding）キーのどちらを貼り付けても
  構いません。必要に応じて内部でエンコードして使用します。
- 関数が照会に失敗して**キーが拒否されたと確認された場合は、保存されたキーを自動的に
  削除**し、再発行の案内を返します。新しいキーを発行して再度登録してください。
