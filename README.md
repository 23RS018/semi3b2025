# pdfから表形式データの抽出・解析

## 1. 目的

- **PDFの表データを自動で抽出し、数値列、文字列を自動で判定しそれぞれに応じた解析を行う**  
  - PDFファイルから表形式のデータを自動で検出・解析・構造化し,Excelなどで出力・保存することで、再利用・分析を容易にする。

---

## 2. プログラムの機能要件

- PDFファイルから複数の抽出戦略（罫線ベース/文字配置ベース/混合モード）でテーブルを自動抽出
- 表の品質を10点満点でスコアリングし、7.0点未満の無効な表を自動除外
- 数値列・数値行を5段階の厳密なチェックで自動判定
  - 除外キーワードフィルタリング（合計、小計など）
  - 数値変換可能率チェック（80%以上）
  - テキスト文字含有率チェック（20%以下）
  - 時間表記除外（コロン含む）
  - 異常値検出（1e15超）
- データ表とテキスト表を自動分類
- 数値データの自動クリーニング（カンマ・円マーク除去、整数変換）
- 結合セルの空白自動補完
- HTML形式での色分け表示（データ表：青色、テキスト表：緑色）
- 集計結果の自動計算
  - 各数値列の合計・件数・平均
  - 集計行（合計/小計/平均）の自動検出と除外

---

## 3. 主な技術・開発環境

- **言語・ライブラリ**
  - Python (3.10.18)
  - pdfplumber
  - pandas
  - openpyxl
  - japanize-matplotlib
  - IPython.display
  
- **開発環境**
  - Jupyter Notebook
  - Anaconda Prompt

- **出力形式**
  - HTML形式での表示（Jupyter Notebook内）
  - 構造化されたDataFrame

---

## 4. 使い方

### 4.1 事前準備

#### 1. パソコンに Python と Jupyter Notebook をインストールする

- このプログラムは「Python」と「Jupyter Notebook」というソフトを使って動く。

#### 2. PDFファイルを配置するディレクトリを準備する
```
./Sample Date/
  └── timetable.pdf  （または任意のPDFファイル）
```

- プログラムは `./Sample Date/` ディレクトリ内のPDFファイルを自動的に処理する。
- ディレクトリが存在しない場合は自動作成される。

#### 3. 仮想環境を作成する（Anaconda Prompt）
```bash
conda create -n pdf_analysis python=3.10.18
conda activate pdf_analysis
```

#### 4. カーネルを作成する
```bash
pip install ipykernel
python -m ipykernel install --user --name pdf_analysis
```

#### 5. Jupyter を開く
```bash
jupyter notebook
```

#### 6. 作成したカーネルの名前の新しいnotebookを作る

---

### 4.2 Jupyter Notebookでプログラムを実行する

#### セル 1: ライブラリのインストール
```python
!pip install -q pdfplumber pandas openpyxl japanize-matplotlib
```

#### セル 2: インポート
```python
import pdfplumber
import pandas as pd
from pathlib import Path
import re
from typing import List, Dict, Any, Tuple
from IPython.display import display, HTML
```

#### セル 3: 定数定義
```python
PDF_DIR = Path("./Sample Date")
EXAMPLE_PDF_NAME = "timetable.pdf"

REQUIRED_META_COLS = ['_SourceFile', '_PageNum', '_TableIndex', 
                      '_NumericCols', '_NumericRows', '_TableType', 
                      '_TableQualityScore', '_IsValidTable']

# 数値列判定のための定数
NUMERIC_RATIO_THRESHOLD = 0.80   # 数値変換可能な要素の最小比率
TEXT_RATIO_THRESHOLD = 0.20      # テキスト文字が含まれる要素の最大比率
ABNORMAL_VALUE_THRESHOLD = 1e15  # 異常値と見なす絶対値の閾値

# 表の妥当性判定のための定数
MIN_TABLE_ROWS = 2               # 最小行数（ヘッダー含まず）
MIN_TABLE_COLS = 2               # 最小列数
MIN_QUALITY_SCORE = 7.0          # 表として認識する最小品質スコア
MAX_EMPTY_CELL_RATIO = 0.7       # 空セルの最大比率
```

#### セル 4: ヘルパー関数定義
```python
def ensure_pdf_files_exist(pdf_dir: Path) -> List[Path]:
    """PDFディレクトリの存在確認とPDFファイル一覧取得"""
    if not pdf_dir.exists():
        print(f"⚠️ ディレクトリ '{pdf_dir}' が見つからなかったため作成しました。")
        pdf_dir.mkdir(parents=True, exist_ok=True)
        return []
    
    pdf_files = list(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"⚠️ '{pdf_dir}' 内にPDFファイルが見つかりません。'{EXAMPLE_PDF_NAME}'などのファイルを配置してください。")
    return pdf_files

def _is_purely_numeric_column(s: pd.Series) -> bool:
    """列または行が純粋な数値データかどうかを判定する（ヘルパー関数）"""
    
    s_str = s.astype(str).str.strip().replace('', pd.NA).dropna()
    
    if s_str.empty:
        return False
    
    # 1. 除外キーワードのフィルタリング
    excluded_keywords = ['合計', '小計', '総計', '計', 'total', 'sum', 'subtotal']
    s_filtered = s_str[~s_str.str.lower().isin(excluded_keywords)]
    
    if len(s_filtered) == 0:
        return False
    
    # 2. 数値変換可能率のチェック
    # カンマ、円マーク、空白を除去
    s_cleaned = s_filtered.str.replace('[,\s¥円]', '', regex=True) 
    
    numeric_count = pd.to_numeric(s_cleaned, errors='coerce').notna().sum()
    numeric_ratio = numeric_count / len(s_filtered)
    
    if numeric_ratio < NUMERIC_RATIO_THRESHOLD:
        return False
    
    # 3. テキスト文字含有率のチェック
    text_pattern = r'[a-zA-Zぁ-んァ-ヶー一-龯]'
    text_count = s_filtered.str.contains(text_pattern, regex=True, na=False).sum()
    text_ratio = text_count / len(s_filtered)
    
    if text_ratio > TEXT_RATIO_THRESHOLD:
        return False
    
    # 4. 時間表記チェック
    if s_filtered.str.contains(':', regex=False, na=False).any():
        return False
    
    # 5. 異常値チェック
    try:
        numeric_values = pd.to_numeric(s_cleaned, errors='coerce').dropna()
        if len(numeric_values) > 0 and (numeric_values.abs() > ABNORMAL_VALUE_THRESHOLD).any():
            return False
    except Exception:
        return False
    
    return True

def _is_valid_table(df: pd.DataFrame) -> Tuple[bool, float]:
    """表の妥当性を判定し、品質スコアを返す
    
    Returns:
        (is_valid, quality_score): 妥当性の真偽値と品質スコア（0-10）
    """
    data_cols = [c for c in df.columns if not c.startswith('_')]
    
    if len(df) < MIN_TABLE_ROWS or len(data_cols) < MIN_TABLE_COLS:
        return False, 0.0
    
    quality_score = 0.0
    
    # 1. 行数・列数チェック（最大2点）
    if len(df) >= 3:
        quality_score += 1.0
    if len(data_cols) >= 3:
        quality_score += 1.0
    
    # 2. 空セル比率チェック（最大2点）
    total_cells = len(df) * len(data_cols)
    empty_cells = (df[data_cols].astype(str).apply(lambda x: x.str.strip()).eq('').sum().sum())
    empty_ratio = empty_cells / total_cells if total_cells > 0 else 1.0
    
    if empty_ratio <= 0.3:
        quality_score += 2.0
    elif empty_ratio <= MAX_EMPTY_CELL_RATIO:
        quality_score += 1.0
    else:
        return False, quality_score  # 空セルが多すぎる場合は無効
    
    # 3. 数値列の存在チェック（最大2点）
    numeric_col_count = sum(1 for col in data_cols if _is_purely_numeric_column(df[col]))
    if numeric_col_count >= 2:
        quality_score += 2.0
    elif numeric_col_count >= 1:
        quality_score += 1.0
    
    # 4. ヘッダーの妥当性チェック（最大2点）
    # ヘッダーに重複が少なく、意味のある名前がついているか
    unique_headers = len([c for c in data_cols if not c.startswith('列')])
    if unique_headers == len(data_cols):
        quality_score += 2.0
    elif unique_headers >= len(data_cols) * 0.5:
        quality_score += 1.0
    
    # 5. データの一貫性チェック（最大2点）
    # 各列のデータ型が一貫しているか
    consistent_cols = 0
    for col in data_cols:
        col_data = df[col].astype(str).str.strip().replace('', pd.NA).dropna()
        if len(col_data) > 0:
            # 列の80%以上が同じパターン（数値 or テキスト）であれば一貫性あり
            numeric_count = pd.to_numeric(col_data.str.replace('[,\s¥円]', '', regex=True), errors='coerce').notna().sum()
            if numeric_count / len(col_data) >= 0.8 or numeric_count / len(col_data) <= 0.2:
                consistent_cols += 1
    
    consistency_ratio = consistent_cols / len(data_cols) if len(data_cols) > 0 else 0
    if consistency_ratio >= 0.8:
        quality_score += 2.0
    elif consistency_ratio >= 0.5:
        quality_score += 1.0
    
    is_valid = quality_score >= MIN_QUALITY_SCORE
    return is_valid, round(quality_score, 1)

def _analyze_table_structure(df: pd.DataFrame) -> Dict[str, Any]:
    """表の構造を分析し、数値列と数値行のインデックスを特定する"""
    result = {
        'numeric_cols': [],
        'numeric_rows': [],
        'is_data_table': False
    }
    
    # メタデータ列を除外
    data_cols = [c for c in df.columns if not c.startswith('_')]
    
    # 1. 列ごとに数値判定
    for col in data_cols:
        if _is_purely_numeric_column(df[col]):
            result['numeric_cols'].append(col)
    
    # 2. 行ごとに数値判定
    for idx in df.index:
        row_data = df.loc[idx, data_cols]
        if _is_purely_numeric_column(row_data):
            result['numeric_rows'].append(idx)
    
    # 3. データ表判定: 少なくとも1つの数値列がある、または行の半分以上が数値行である
    numeric_row_ratio = len(result['numeric_rows']) / len(df) if len(df) > 0 else 0
    result['is_data_table'] = len(result['numeric_cols']) > 0 or numeric_row_ratio > 0.5
    
    return result

def _clean_numeric_column(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """数値列をクリーニングし、整数型に変換する"""
    
    # 文字列に変換し、カンマ、円マーク、空白を除去
    cleaned_series = df[col].astype(str).str.strip().str.replace('[,\s¥円]', '', regex=True)
    
    # 数値（小数点、マイナス含む）以外の文字を削除
    cleaned_series = cleaned_series.str.replace(r'[^0-9\.\-]', '', regex=True)
    
    # 空白やハイフンを欠損値に変換
    cleaned_series = cleaned_series.replace({'': pd.NA, '-': pd.NA})
    
    # 数値に変換し、四捨五入して整数型（欠損値対応のInt64）に格納
    df[col] = pd.to_numeric(cleaned_series, errors='coerce').round().astype('Int64')
    return df

def fill_merged_cells(table: List[List], merge_info: List[Dict]) -> List[List]:
    """結合セル情報を使って、テーブルの空白セルを埋める"""
    if not table or not merge_info:
        return table
    
    # テーブルのサイズを取得
    num_rows = len(table)
    num_cols = len(table[0]) if table else 0
    
    # 結果用のテーブルをコピー
    result = [row[:] for row in table]
    
    for merge in merge_info:
        try:
            top = merge.get('top', 0)
            bottom = merge.get('bottom', 0)
            left = merge.get('left', 0)
            right = merge.get('right', 0)
            text = merge.get('text', '').strip()
            
            # インデックスの範囲チェック
            if top < 0 or bottom >= num_rows or left < 0 or right >= num_cols:
                continue
            
            # 結合範囲が不正な場合はスキップ
            if top > bottom or left > right:
                continue
            
            # 結合セルの範囲を埋める
            for row_idx in range(top, bottom + 1):
                # 行が存在するか確認
                if row_idx >= len(result):
                    break
                
                for col_idx in range(left, right + 1):
                    # 列が存在するか確認
                    if col_idx >= len(result[row_idx]):
                        break
                    
                    # Noneまたは空文字列の場合のみ埋める
                    if result[row_idx][col_idx] is None or str(result[row_idx][col_idx]).strip() == '':
                        result[row_idx][col_idx] = text
                        
        except Exception as e:
            # 個別の結合セル処理でエラーが発生しても続行
            print(f"    ⚠️ 結合セル処理エラー（スキップ）: {e}")
            continue
    
    return result
```

#### セル 5: PDFからテーブルを抽出
```python
def extract_all_tables_from_pdf(pdf_path: Path) -> List[pd.DataFrame]:
    """PDFファイルから全ての表を抽出し、データ表かテキスト表かを判定する"""
    print(f"📄 処理中: {pdf_path.name}")
    dataframes = []
    
    # 抽出設定のリスト: 複数の戦略で試行し、最も多く抽出できたものを採用
    table_settings_list = [
        {"name": "罫線ベース", "vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 3, "intersection_tolerance": 5},
        {"name": "文字配置ベース", "vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 5, "intersection_tolerance": 5},
        {"name": "混合モード", "vertical_strategy": "lines", "horizontal_strategy": "text", "snap_tolerance": 3, "intersection_tolerance": 5}
    ]
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                best_tables = []
                best_table_count = -1
                best_setting_name = ""
                
                # 最適な抽出設定を選択
                for settings_data in table_settings_list:
                    settings = {k: v for k, v in settings_data.items() if k != "name"}
                    setting_name = settings_data["name"]
                    try:
                        tables = page.extract_tables(table_settings=settings)
                        # 1行を超える（ヘッダーとデータ行がある）有効なテーブルのみをカウント
                        valid_tables = [t for t in tables if t and len(t) > 1]
                        
                        if len(valid_tables) > best_table_count:
                            best_tables = valid_tables
                            best_table_count = len(valid_tables)
                            best_setting_name = setting_name
                    except Exception:
                        continue
                
                if best_tables:
                    print(f"  ページ {page_num}: {best_setting_name}で {len(best_tables)} 個の表を抽出")
                
                for table_index, tbl in enumerate(best_tables):
                    header = tbl[0]
                    data = tbl[1:]
                    
                    # ヘッダーのクリーニングと重複列名対応
                    clean_header = []
                    for i, h in enumerate(header):
                        clean_h = str(h).strip().replace("\n", " ") if h and str(h).strip() else f"列{i+1}"
                        # 列名の重複を避ける
                        original_h = clean_h
                        count = 1
                        while clean_h in clean_header:
                            count += 1
                            clean_h = f"{original_h}_{count}"
                        clean_header.append(clean_h)
                    
                    df = pd.DataFrame(data, columns=clean_header)
                    # 全てのセルが空の行を削除
                    df = df.loc[~(df.astype(str).apply(lambda x: x.str.strip()).eq('').all(axis=1))]
                    
                    if df.empty:
                        continue
                    
                    # メタデータの追加と文字列の整形
                    df['_SourceFile'] = pdf_path.name
                    df['_PageNum'] = page_num
                    df['_TableIndex'] = table_index + 1
                    df = df.fillna("")
                    
                    for col in df.columns:
                        if col not in REQUIRED_META_COLS:
                            # 複数行を改行区切りに整形
                            df[col] = df[col].astype(str).apply(
                                lambda x: '\n'.join([line.strip() for line in str(x).split('\n') if line.strip()])
                            )
                    
                    # 表の妥当性判定
                    is_valid, quality_score = _is_valid_table(df)
                    
                    # 妥当でない表はスキップ
                    if not is_valid:
                        continue
                    
                    # 表の構造分析と数値列のクリーニング
                    structure = _analyze_table_structure(df)
                    numeric_cols = structure['numeric_cols']
                    numeric_rows = structure['numeric_rows']
                    is_data_table = structure['is_data_table']
                    
                    for col in numeric_cols:
                        df = _clean_numeric_column(df, col)
                    
                    # メタデータの最終調整
                    df['_TableType'] = "データ表" if is_data_table else "テキスト表"
                    df['_NumericCols'] = ",".join(numeric_cols)
                    df['_NumericRows'] = ",".join([str(r) for r in numeric_rows])
                    df['_IsValidTable'] = "はい"
                    df['_TableQualityScore'] = quality_score

                    dataframes.append(df)
        
        print(f"  ✓ 最終的に {len(dataframes)} 個の表を抽出\n")
        return dataframes

    except Exception as e:
        print(f"🚨 エラー発生 ({pdf_path.name}): {e}")
        return []
```

#### セル 6: 表示と解析
```python
def display_and_analyze_tables(all_table_dfs: List[pd.DataFrame]):
    """抽出した表を整形して表示し、数値列の集計を行う"""
    
    if not all_table_dfs:
        print("\n⚠️ 有効な表が見つかりませんでした。")
        return

    print(f"✅ 合計 {len(all_table_dfs)} 個の表を抽出しました。\n")
    print("=" * 70 + "\n")
    
    for idx, df in enumerate(all_table_dfs, 1):
        source_file = str(df['_SourceFile'].iloc[0]) 
        page_num = df['_PageNum'].iloc[0]
        table_index = df['_TableIndex'].iloc[0]
        table_type = df['_TableType'].iloc[0]
        quality_score = df['_TableQualityScore'].iloc[0]
        numeric_cols_str = df['_NumericCols'].iloc[0]
        numeric_rows_str = df['_NumericRows'].iloc[0]
        numeric_cols = [c for c in numeric_cols_str.split(',') if c] if numeric_cols_str else []
        numeric_rows = [int(r) for r in numeric_rows_str.split(',') if r] if numeric_rows_str else []
        
        print(f"【表 {idx}/{len(all_table_dfs)}】")
        print(f"  📄 ファイル: {source_file}")
        print(f"  📑 ページ: {page_num} / 表番号: {table_index}")
        print(f"  📊 種類: {table_type}")
        
        if numeric_cols:
            print(f"  🔢 数値列: {', '.join(numeric_cols)} ({len(numeric_cols)}列)")
        if numeric_rows:
            print(f"  🔢 数値行: {len(numeric_rows)}行")
        print()
        
        df_display = df.drop(columns=REQUIRED_META_COLS, errors='ignore').copy()
        
        if df_display.empty:
            print("  ℹ️ 表示可能なデータがありません\n")
            continue
        
        print("▼ 抽出データ:")
        
        # HTML表示のために改行を  に変換
        df_html = df_display.copy()
        for col in df_html.columns:
            df_html[col] = df_html[col].astype(str).str.replace('\n', '', regex=False)
        
        # テーブルの種類に応じて色を変更
        header_color = "#2196F3" if table_type == "データ表" else "#4CAF50"
        header_border = "#1976D2" if table_type == "データ表" else "#388E3C"
        
        # HTMLの表示スタイルをシンプルに
        html_table = df_html.to_html(index=False, escape=False, classes=f"pdf-table-{idx}")
        
        styled_html = f"""
        
            .pdf-table-{idx} {{
                border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 13px;
            }}
            .pdf-table-{idx} th {{
                background-color: {header_color}; color: white; padding: 10px; border: 2px solid {header_border};
                text-align: center; font-weight: bold;
            }}
            .pdf-table-{idx} td {{
                padding: 8px; border: 1px solid #ddd; text-align: center; vertical-align: top; min-width: 80px;
            }}
        
        {html_table}
        """
        
        display(HTML(styled_html))
        
        # 数値集計
        if not numeric_cols:
            print("  ℹ️ 数値列が検出されませんでした（集計スキップ）")
        else:
            print("\n📊 数値データの集計:")
            
            # 集計行を検出するためのキーワード
            aggregation_keywords = ['合計', '小計', '総計', '計', '平均', '平均値', 'total', 'sum', 'subtotal', 'average', 'avg', 'mean']
            
            # 各行が集計行かどうかを判定
            is_aggregation_row = pd.Series([False] * len(df), index=df.index)
            
            # 全ての列をチェックして、集計キーワードが含まれている行を特定
            data_cols = [c for c in df.columns if not c.startswith('_')]
            for col in data_cols:
                for idx_row in df.index:
                    cell_value = str(df.loc[idx_row, col]).strip().lower()
                    if any(keyword in cell_value for keyword in aggregation_keywords):
                        is_aggregation_row.loc[idx_row] = True
                        break
            
            # 集計行を除外したデータフレームを作成
            df_without_aggregation = df[~is_aggregation_row].copy()
            
            excluded_count = is_aggregation_row.sum()
            if excluded_count > 0:
                print(f"   ℹ️ {excluded_count}行の集計行を検出し、計算から除外しました")
            
            for col in numeric_cols:
                try:
                    # 集計行を除外したデータで計算
                    numeric_series = pd.to_numeric(df_without_aggregation[col], errors='coerce')
                    total = numeric_series.sum()
                    count = numeric_series.notna().sum()
                    
                    if pd.notna(total) and count > 0:
                        avg = total / count
                        print(f"   【{col}】")
                        print(f"     • 合計: {total:,.0f}")
                        print(f"     • 件数: {count}")
                        print(f"     • 平均: {avg:,.1f}")
                except Exception as e:
                    print(f"   【{col}】集計エラー: {e}")

        print("\n" + "=" * 70 + "\n")
```

#### セル 7: メイン処理
```python
def main():
    """メイン実行関数"""
    
    pdf_files = ensure_pdf_files_exist(PDF_DIR)

    if not pdf_files:
        return

    print(f"📂 {len(pdf_files)} 個のPDFファイルを発見\n")
    
    all_table_dfs = []
    for pdf_path in pdf_files:
        all_table_dfs.extend(extract_all_tables_from_pdf(pdf_path))

    display_and_analyze_tables(all_table_dfs)

if __name__ == "__main__":
    main()
```

---

## 5. 今後の予定

- **画像形式PDF（スキャンPDF）への対応**
  - OCR（光学文字認識）機能の統合
  - 画像内の表を検出・抽出
  
- **より複雑な表構造への対応強化**
  - 多階層ヘッダー（2行以上のヘッダー）の検出
  - 縦書きテキストの処理改善
  - 不規則な結合セルパターンへの対応
  
- **機械学習による表領域検出の検討**
  - より複雑なレイアウトのPDFへの対応
  - 表/非表の自動判別精度の向上
