# SpreadSheet_to_PowerPoint

## 1. 概要
Googleスプレッドシートに記載されたデータを元に、PowerPointのスライドを作成する。

## 2. システム構成
- Googleスプレッドシート : 問題文と答えのデータを格納。Google Formからアンケートを取得するケースでも簡単に適用可能。
- Python : Googleスプレッドシートからデータを取得し、PowerPointのスライドを編集。
- PowerPointテンプレート: スライドのデザインの雛形となるテンプレートファイル。

## 3. 準備
- 3-1. Googleスプレッドシートに、パワポに埋め込む元データを用意

- 3-2. 雛形となるPowerPointファイルを用意
Python pptxライブラリではアニメーション設定はおろか、スライドのコピペすらできない。
が、「特定のレイアウト(「タイトルとコンテンツ」とか「白紙」とか)を指定して新しいスライドを作成」はできるので、PowerPoint側でスライドマスター機能で雛形となるレイアウトを作成または編集しておく。

- 3-3. Google APIの設定
まず、事前にGoogle Cloud ConsoleでAPIキーを取得し、サービスアカウントを設定する。
キーは json で保存し、pythonファイルと同じディレクトリに置く。
また、Google Cloud Console から Google Drive API および Google Sheets API を有効化しておく。

## 4. システム動作
- 4-1. Googleスプレッドシートからのデータ取得:
Google SheetsAPIを利用して、指定されたスプレッドシートから問題文と答えのデータを抽出する。

- 4-2. PowerPointテンプレートの編集:
指定されたPowerPointテンプレートファイルを読みこむ。
適宜スライドを追加し、問題文と答えの文字列を埋め込む。

- 4-3. PowerPointファイルの保存:
作成したスライドを新しいPowerPointファイルとして保存。
