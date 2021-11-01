# 目的
講師が何を教えれるかの情報を検索できるシステムがなかったため。

そのようなシステムがないので、誰が何をできるかは知っている人に聞くか、やったことあるかどうかを過去に遡って調べる必要があった。

これはとても無駄だと感じたのでこのようなシステムを作った。

また、VBAやGASを書いたことがなないので書いてみようと思ったことや、自分以外の多数が使うシステムを開発してみたかったことも理由の一つである。

# 使用言語の決定理由など
私は半年後に退職する。そのため、私の借りているサーバ上にシステムを構築すると、退職者が個人情報にアクセスできてしまうことになるので個人情報流出の可能性がある。また、講師情報をネット上に置くことは教室長も許可してくれないため、検索システムはVBAで構築することにした。

また、講師番号と科目の情報はGoogleスプレッドシート上においてもいいという許可が出たので、データの収集、ある程度のデータの整形はGoogleスプレッドシート上で行うことにした。

# データの収集から検索までの流れ

![haichi_sarch全体図](https://ateruimashin.com/diary/wp-content/uploads/2021/11/94a83d59d90e26fd0b06e6dd46fc9d69.png)

# 検索システム外観と機能説明

## 外観

![プロトタイプ画像](https://ateruimashin.com/diary/wp-content/uploads/2021/10/eba3211de2501239529b7359e038b853.png)  

## 機能説明

1. 科目検索: 科目から教務可能な講師を検索できる
2. 講師検索: 講師名からその講師の教務可能/不可能を一覧表示する
3. 講師教務可能判定: 講師名と科目からその講師がその科目を教務可能か表示する

# 環境

## 開発環境

- Windows10またはWindows11
- Microsoft 365 Apps for enterprose - Excel version 2110

## 運用環境

- WIndows10
- Excel 2016

# 実装した機能

- 検索システム(VBA)
- ローカルにある講師名簿とGoogleスプレッドシートから取得したデータを一つの表にする機能(VBA)
- 退職者を表から削除する(GAS)
- Googleスプレッドシート上の名簿にいない講師が新しくフォームを送信したとき、Googleスプレッドシート上の名簿に講師を追加する(GAS)
- いつでも自分の情報を更新できる(GAS)

# 仕様であり不具合でないもの

- 講師教務可能判定をしたいとき、科目を最後まで埋めずに検索ボタンを押すと講師検索になる



# 既知の不具合と修正時期

- Excel2016ではPower QueryのクエリをVBAからRefreshしようとするとエラーが出る。そのため現在その機能を削除している。修正時期は未定。詳しくは[release v1.0.1](https://github.com/ateruimashin/haichi_sarch/releases/tag/v1.0.1)を参照してください。

# 今後の予定

- リリースする

# お賃金
タダ働きである。お金がほしい。  
｡оО(｡´•ㅅ•｡)Оо｡おじは貧乏...

# More...

- [VBAを使って検索システムを作ってみた](https://ateruimashin.com/diary/2021/09/vba-sarch/)
- [退職者リストに沿って名簿から退職者を削除するスクリプトをGASを使って書く](https://ateruimashin.com/diary/2021/10/gas_delete_list/)
- [Googleフォームの回答から名簿や必要な表を作るスクリプトをGASを使って書く](https://ateruimashin.com/diary/2021/10/gas-script/)
- [講師の情報を検索するシステムをVBAで作った](https://ateruimashin.com/diary/2021/10/haichi_sarch_userform/)
- [VBAで2つのシートの結合、マクロ実行前のパスワード認証などを実装した](https://ateruimashin.com/diary/2021/10/vba_merge_etc/)

