# ListAuthors
学会講演オーサーリスト生成マクロ
xlwingsを使ったエクセル用マクロとオーサー情報をまとめたエクセルファイルのセットです。python3およびxlwingsをインストールして使用してください。
## 使い方
<img width="1398" alt="Screen Shot 2021-12-05 at 19 40 01" src="https://user-images.githubusercontent.com/18593190/144743639-33546ddd-92a8-4cb6-b6cf-6923ab0f9d25.png">
オレンジ色のセルを編集することで出力されるオーサーリストを変更できます。Groupの行にはオーサーリストに含めたいメンバーが所属しているGroupのIDを１セルに１つずつ並べます。importantの行には、オーサーの並び順において優先したい人（登壇者、責任著者など）を１セルに一人ずつ入力しておきます。

エクセルのリボンからxlwingsを選び、Run mainをクリックしてマクロを走らせます。するとJPSの様式に整形された日本語と英語のオーサーリストが表示されます。
<img width="1398" alt="Screen Shot 2021-12-05 at 20 07 31" src="https://user-images.githubusercontent.com/18593190/144744004-577c13f6-5cca-434e-a789-9159ee6b5fda.png">

### Member
Memberのシートには、オーサーリストに含まれる可能性がある人の情報を入力します。Groups, Affiliationsの列には、後述のGroupシートとAffiliationシートで記述されるグループ、機関のIDをカンマで区切って並べます。Name_Sort_Japanese_1, Name_Sort_English_1, Name_Print_English_1は必ず入力してください。
<img width="1398" alt="Screen Shot 2021-12-05 at 19 45 59" src="https://user-images.githubusercontent.com/18593190/144743743-1dd7acc9-c4b2-48d1-991c-0b2b85174064.png">

### Group
Groupはオーサーに含まれるか否かを判断する基準として使われる人の集合です。
<img width="1398" alt="Screen Shot 2021-12-05 at 19 46 03" src="https://user-images.githubusercontent.com/18593190/144743755-fc503ccd-ecc7-4a51-8f0c-bfd097c9c2ff.png">

### Affiliation
Affiliationはオーサーの所属機関です。全てのメンバーは１つかそれ以上のAffiliationを持っていなければなりません。
<img width="1398" alt="Screen Shot 2021-12-05 at 19 46 06" src="https://user-images.githubusercontent.com/18593190/144743859-5b8facb1-7b41-4e9d-acb8-37226c99672f.png">
