# gui_controller.py

## 目的
OCRプロセスの開始/停止と稼働状況の確認をGUIで行います。

## 動作概要
- `.env` を読み込み、STOP/STATUS/LOGファイルの場所を取得
- `ocrStart.py` をサブプロセスで起動
- STOPファイル作成で安全停止
- `ocr.status` のハートビートで稼働判定
- `ocr.latest.txt` をGUIに表示

## 使い方
1. `python gui_controller.py`
2. `Start` で起動
3. `Stop` で停止

## 注意
OCRがGUI外で動いていても、状態表示は更新されます。
