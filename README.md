# VBA-Dictionary
# 連想配列生成サポート用のVBA

- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

その他、実行環境など報告していただくと感謝感激雨霰。

# 使い方

## 「実行サンプル 連想配列作成.xlsm」の使い方

「実行サンプル 連想配列作成.xlsm」には「ModDictionary.bas」内のプロシージャの実行サンプルプロシージャのボタンが登録してある。

各ボタンを押して使用を確かめていただきたし
![実行サンプル中身](https://user-images.githubusercontent.com/73621859/130730462-d00a6218-3777-4ae1-b83c-cded7dceaad6.jpg)


## 設定

実行サンプル「実行サンプル 連想配列作成.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModDictionary.bas

### 設定2（参照ライブラリ）

特になし

## 現在「Dictionary.bas」にて使用できるプロシージャ一覧

- MakeDictFromArray		:Keyとなる階層状の配列（二次元配列）と、Itemとなる一次元配列を入力して連想配列を作成する
- MakeDictFromArrayWithItem	:MakeDictFromArrayで生成するような連想配列のItemが連想配列となる連想配列を生成する。

「MakeDictFromArray」の使い方
![MakeDictFromArrayの使い方](https://user-images.githubusercontent.com/73621859/128442700-97bba6a0-c109-487a-9f8e-79fe7de18d0a.jpg)


「MakeDictFromArrayWithItem」の使い方
![MakeDictFromArrayWithItemの使い方](https://user-images.githubusercontent.com/73621859/128448180-2f5dc674-cdea-4001-b24e-56ddc9dee756.jpg)