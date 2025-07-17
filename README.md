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

#### 1. パソコンに Python と Jupyter Notebook をインストールする

- このプログラムは「Python」と「Jupyter Notebook」というソフトを使って動く。
#### 2. C:\Users\User\Documents\GitHub\semi3b2025のようにディレクトリを準備しておく

- このプログラムは以上のディレクトリ内にExcel（.xlsx）、PNG画像（グラフ）が保存される。


#### 3. 仮想環境を作成する(Anaconda Prompt)
```
 conda create -n 名前 python=バージョン
 conda activate 名前
```
#### 4. カーネルを作成する
```
pip install ipykernel
python -m ipykernel install --user --name 名前
```
#### 5. jupyterを開く
```
 jupyter notebook
```
#### 6. 作成したカーネルの名前の新しいnotebookを作る


---

### 4.2 jupyter notebookでプログラムを実行する。

1. ライブラリをインストールする  

```python
!pip install -q pdfplumber pandas openpyxl japanize-matplotlib
```
2. インポートをする
```python
import pdfplumber, pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import japanize_matplotlib
```
3. pdfからテーブルを抽出
```python
pdf_path = Path("商品発注表.pdf")
assert pdf_path.exists(), f"PDF が見つかりません: {pdf_path.resolve()}"
 
tables = []
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        tables.extend(page.extract_tables())
 
# ヘッダーに「商品コード」があるテーブルを採用
df = None
for tbl in tables:
    if tbl and any("商品コード" in (cell or "") for cell in tbl[0]):
        df = pd.DataFrame(tbl[1:], columns=tbl[0])
        break
if df is None:
    raise ValueError("対象となるテーブルが見つかりません。")
 
print("抽出列:", df.columns.tolist())  # 列名を確認
```
4. 列名を標準化し、必要列を自動検出
```python 
df = df.rename(
    columns=lambda c: (
        c.strip()
         .replace("（", "(")
         .replace("）", ")")
         .replace(" ", "")
         .replace("\t", "")
    )
)
 
# キーワードで列名を見つける
qty_col   = next((c for c in df.columns if "数量" in c),  None)
price_col = next((c for c in df.columns if "単価" in c),  None)
amt_col   = next((c for c in df.columns if "金額" in c),  None)
 
print("数量列:", qty_col, "| 単価列:", price_col, "| 金額列:", amt_col)
if None in (qty_col, price_col, amt_col):
    raise KeyError("数量・単価・金額に該当する列が見つかりません。")
```
5. 数値列の整形
```python
for col in (qty_col, price_col, amt_col):
    df[col] = (
        df[col]
          .astype(str)
          .str.replace(",", "", regex=False)
          .str.replace("円", "", regex=False)
          .str.strip()
          .astype(float)
    )
df[qty_col]   = df[qty_col].astype(int)
df[price_col] = df[price_col].astype(int)
df[amt_col]   = df[amt_col].astype(int)
 
display(df.head())  # プレビュー
```
  ```
 6. Excel へ書き出し
 ```python
 # 保存先ディレクトリ
save_dir = Path(r"C:\Users\User\Documents\GitHub\semi3b2025")

# Excelファイルのパスを作成
excel_path = save_dir / "商品発注表.xlsx"

# Excelファイルを書き出す
df.to_excel(excel_path, index=False, engine="openpyxl")

# 絶対パスを表示
print(f"Excel ファイルを保存しました → {excel_path.resolve()}")
 ```
 7. 数量を棒グラフで可視化
 ```python
 # 保存先ディレクトリ
save_dir = Path(r"C:\Users\User\Documents\GitHub\semi3b2025")

# ディレクトリが存在しない場合は作成する（任意）
save_dir.mkdir(parents=True, exist_ok=True)

# グラフに使うキー（商品名があれば優先、無ければ商品コード）
group_key = "商品名" if "商品名" in df.columns else "商品コード"

# 品目ごと数量を合計して多い順に並べ替え
qty_by_item = (
    df.groupby(group_key)[qty_col]
      .sum()
      .sort_values(ascending=False)
)

plt.figure(figsize=(10, 6))
qty_by_item.head(20).plot(kind="bar")          # 上位 20 品目を表示
plt.title("商品の数量（上位 20 品目）")
plt.ylabel("数量")
plt.xlabel(group_key)
plt.xticks(rotation=45, ha="right")
plt.tight_layout()

# 保存パスを作成
image_path = save_dir / "graph.png"

# 画像ファイルとして保存
plt.savefig(image_path, dpi=300)

print(f"グラフ画像を保存しました → {image_path.resolve()}")

plt.show()
 ```
 8. 追加で解析
 ```python
 print("▼ 集計レポート")
total_qty = df[qty_col].sum()
total_amt = df[amt_col].sum()
total_amt_tax = total_amt * 1.10

print(f"  総数量          : {total_qty:,}")
print(f"  総金額 (税抜)   : {total_amt:,}")
print(f"  総金額 (税込)   : {total_amt_tax:,.0f}")

print("\n▼ 数量 上位 5 品目")
display(qty_by_item.head(5).to_frame(name="数量"))

print("\n▼ 金額 上位 5 品目")
amt_by_item = (
    df.groupby(group_key)[amt_col]
      .sum()
      .sort_values(ascending=False)
)
display(amt_by_item.head(5).to_frame(name="金額 (円)"))
 ```

## 5.今後の予定
- 柔軟な表形式データの抽出・解析を可能にする。
- 画像形式のpdfの判別とその抽出・解析を可能にする。
- セル結合、縦文字などを含む複雑な表形式データの抽出・解析を罫線データの抽出・解析により可能にする。
- 今回試すことのできなかったライブラリを用いて実験する。
