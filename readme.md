# NetSuiteの総勘定元帳から弥生会計の仕訳データを作成するツール

## 概要

## 開発環境
* Python 3
* debian 11.7
* docker 20.10.6

## 機能一覧

### 伝票番号ごとのソート、貸借チェック

### 仕訳フラグの自動付与
識別フラグ(1行の伝票データ：2111、複数行の伝票データ1行目：2110、中間の行：2100、最終行：2101)を自動で付与する。

### 消費税自動調整
以下の場合に限り消費税を自動調整する。
* 消費税の科目が仕訳に含まれる
* すべての借方(または貸方)に消費税がかかる
* 1行のみに消費税がかかっている場合
* 各税率で判断している

### 弥生会計の形式に合わせた仕訳データをcsvに出力
勘定科目については24バイトの制限かつ先頭末尾に空白がないことが条件のため、それに合わせて整形する。

### 勘定科目比較
NetSuiteと弥生会計の勘定科目を比較する。弥生会計に登録されていない場合は事前に弥生会計に登録してもらう。
