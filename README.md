# 受付管理ツール

これはGoogleフォームを利用した受付管理をサポートするGASスクリプトです。

次のような機能があります。

・予約受付フォームの自動生成

・注文数量に達したら自動的に受付終了

・注文状況を管理者へメール通知

## システム構成
![image](https://user-images.githubusercontent.com/11681096/130345647-a78bd285-604c-425b-b470-b286dd2ddf61.png)

## 必要なリソースとソースコードの組み込み方
### １．受付シート（Google Spreadsheet）

https://docs.google.com/spreadsheets/d/1_7gn1R2TBbweu9Fhy9Uz-nh2s0-9Tznd87B8kk59d-Q/edit?usp=sharing

スクリプトエディタに「管理メニュー.gs」を追加

スクリプトエディタに「フォーム１.gs」を追加

必要に応じて「フォーム２.gs」「フォーム３.gs」も追加

### ２．予約受付フォーム（Google Form）

https://forms.gle/BoYNguWb8aCY1Cv36

スクリプトエディタに「フォーム１送信時.gs」を追加

必要に応じて同じフォームを作成、それぞれのスクリプトエディタに「フォーム２送信時.gs」「フォーム３送信時.gs」を追加

### ３．IDの設定（スクリプトエディタ）
設置したすべてのソースコードにて、openByIdで作成したSpreadsheet・FormのファイルIDを記入

「フォームN送信時.gs」にて、送信先メールアドレスを記入


