# 医療費控除集計ツール

毎年の確定申告（医療費控除）に必要な明細の作成を、Pythonで自動化する実用ツールです。散らばりがちな領収書のデータをCSVとしてまとめておくことで、openpyxlを用いて関数埋め込み済みの綺麗なxlsxファイルを生成します。

## ⚡ 特徴
* ファイル名から「インデックス」「誰が」「どこに」支払ったかを自動抽出
* 保険金などの補填金をインデックスベースで自動的に紐付け・合算
* Excel側には値ではなく合計数式を直接埋め込み

## 🛠️ システム要件およびインストール
* Python 3.x
* pandas
* openpyxl

```bash
pip install pandas openpyxl
```

## 📋 データの準備
(対象の年度フォルダ（例：`sample`）を作成し、以下のように配置します)

```text
📁 sample/
├── 📄 001_名前_医療機関名.csv   # Format: 日付, 金額
├── 📄 002_名前_医療機関名.csv
└── 📄 refund.csv             # Format: インデックス, 補填金額
```

## 🚀 Usage (使い方)
1. `med.ini` に住所や氏名、パスを設定
2. 集計年(フォルダ名)指定し実行

```bash
python med.py 2020 -o 2020.xlsx
```

## 📄 License
This project is licensed under the MIT License.
