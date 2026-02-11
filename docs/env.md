# .env / .env.example

## 目的
実行設定をまとめるファイルです。`.env.example` を `.env` にコピーして使います。

## 重要項目
- `WEBHOOK_URL`: Discord Webhook URL
- `LINE_WINDOW_TITLE`: LINEウィンドウ名
- `CROP_RECT`: チャット欄の矩形
- `POLL_SEC`: 取得間隔
- `TESSERACT_CMD`: tesseract.exe のパス（任意）

## GUI関連
- `STOP_FILE`: 停止用ファイル名
- `STATUS_FILE`: 稼働状況ファイル名
- `LOG_FILE`: 最新OCRテキスト

## OCR精度調整
- `PREPROCESS`: 前処理の有無
- `THRESHOLD`: 2値化閾値
- `OCR_SCALE`: 拡大率
- `MEDIAN_SIZE`: ノイズ除去
- `SHARPEN`: シャープ化
