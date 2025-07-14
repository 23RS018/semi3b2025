# pdfから表形式データの抽出・解析

## 1. 目的

- **PDFの表データを自動で抽出し、集計・可視化・Excel出力する**  
  - PDFファイルから表形式のデータを自動で検出・解析・構造化し,Excelなどで出力・保存することで、再利用・分析を容易にする。

---

## 2. プログラムの機能要件

- PDFファイルからテーブルを抽出
- 商品コードを含むテーブルを自動検出
- 列名を整形し、以下の必須列を自動検出
  - 数量列
  - 単価列
  - 金額列
- 数値データ（数量・単価・金額）の文字列整形と数値変換
- Excelファイル（.xlsx）への書き出し
- 商品別の数量を棒グラフ化し画像保存
- 集計結果を表示
  - 総数量
  - 総金額（税抜・税込）
  - 数量上位5品目
  - 金額上位5品目

---

## 3. 主な技術・開発環境

- **言語・ライブラリ**
  - Python(3.10.18)
  - pdfplumber
  - pandas
  - matplotlib
  - japanize-matplotlib
  - openpyxl
- **開発環境**
  - jupyter notebook
  - Anaconda Prompt

- **出力形式**
  - Excel（.xlsx）
  - PNG画像（グラフ）

---

## 4. 使い方

### 4.1 事前準備

#### パソコンに Python と Jupyter Notebook をインストールする

- このプログラムは「Python」と「Jupyter Notebook」というソフトを使って動きます。
#### 仮想環境をする(Anaconda Prompt)
```
- conda create -n 名前 python=バージョン
- conda activate 名前
```
#### jupyterを開く
```
- jupyter notebook
```
---

### 4.2 Jupyter Notebook でライブラリをインストールする

Jupyter Notebook 上で以下のセルを実行してください。  

```python
# ── セル 1: ライブラリのインストール ──────────────
!pip install -q pdfplumber pandas openpyxl
!pip install -q japanize-matplotlib
```


## 5.今後の予定
- 柔軟な表形式データの抽出・解析を可能にする。
- 画像形式のpdfの判別とその抽出・解析
- セル結合、縦文字などを含む複雑な表形式データの抽出・解析(罫線データの抽出・解析)
