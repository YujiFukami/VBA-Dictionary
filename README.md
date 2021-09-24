# VBA-Dictionary
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

実行環境など報告していただくと感謝感激雨霰。

# 説明
配列から連想配列を作成できる。

## 活用例
簡単に連想配列を作成できる。

# 使い方
実行サンプル「Sample-Dictionary.xlsm」の中の使い方は以下の通り。

##  サンプル中身

各サンプルの実行ボタンとして、

「MakeDictFromArrayのテスト」「MakeDictFromArrayWithItemのテスト」「MakeDictFromArray1Dのテスト」が設定してある。

![実行前](Readme用/実行前.jpg)

各実行ボタンを押した後。

![実行後](Readme用/実行後.jpg)


##  それぞれのプロシージャ中身


「MakeDictFromArrayのテスト」登録のプロシージャの中身

![中身1](Readme用/中身1.jpg)

入力の詳細

![入力1](Readme用/入力1.jpg)


「MakeDictFromArrayWithItemのテスト」登録のプロシージャの中身

![中身2](Readme用/中身2.jpg)

入力の詳細

![入力2](Readme用/入力2.jpg)

「MakeDictFromArray1Dのテスト」登録のプロシージャの中身

![中身3](Readme用/中身3.jpg)

入力の詳細

![入力3](Readme用/入力3.jpg)


##  各プロシージャの紹介

プロシージャ名：MakeDictFromArray

階層型配列から連想配列を作成する。


引数

-  KaisoArray2D  階層型配列となる二次元配列。連想配列のKeyになる。

-  ItemArray1D   連想配列のItemになる一次元配列。


プロシージャ名：MakeDictFromArrayWithItem

階層型配列から連想配列を作成する。

連想配列の末端要素がさらに連想配列（要素連想配列）となっていて、もととなるキー配列、アイテム配列を入力する。


引数

-  KaisoArray2D  階層型配列となる二次元配列。連想配列のKeyになる。

-  KeyArray1D	末端要素になる連想配列のKey。一次元配列。

-  ItemArray1D   末端要素になる連想配列のItem。二次元配列。



プロシージャ名：MakeDictFromArray1D

KeyとItemのそれぞれ一次元配列階層型配列から連想配列を作成する。


引数

-  KeyArray1D	連想配列のKey。一次元配列。

-  ItemArray1D   連想配列のItem。一次元配列。



## 設定
実行サンプル「Sample-Dictionary.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModTest.bas
-  ModMakeDictFromArray.bas
-  ModMakeDictFromArrayWithItem.bas
-  ModMakeDictFromArray1D.bas

### 設定2（参照ライブラリ）
なし

