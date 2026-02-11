# ocrStart.py

## 目的
LINE PC版の画面をOCRで読み取り、Discord Webhookへ送信します。

## 挙動
1. LINEウィンドウをタイトル検索
2. 画面をキャプチャ
3. 必要に応じてCROP_RECTで切り取り
4. 前処理（2値化など）
5. OCR実行
6. 変化があればDiscord送信
7. 状態/ログファイルを書き出し

## 主な設定（.env）
- `WEBHOOK_URL`: Discord Webhook URL
- `LINE_WINDOW_TITLE`: LINEウィンドウタイトル（部分一致OK）
- `CROP_RECT`: `left,top,right,bottom`
- `POLL_SEC`: 取得間隔（秒）
- `OCR_LANG`: `jpn+eng` など
- `PREPROCESS`: 前処理の有無
- `THRESHOLD`: 2値化閾値
- `OCR_SCALE`: OCR用の拡大率
- `MEDIAN_SIZE`: ノイズ除去（0で無効）
- `SHARPEN`: シャープ化の有無
- `ADD_TIMESTAMP`: 投稿にタイムスタンプ付与
- `STOP_FILE`: 安全停止用ファイル名
- `PID_FILE`: PIDファイル
- `STATUS_FILE`: 稼働状況ファイル
- `HEARTBEAT_SEC`: ハートビート間隔
- `LOG_FILE`: 最新OCRテキスト
- `LOG_MAX_CHARS`: ログ最大文字数
- `TESSERACT_CMD`: tesseract.exe のフルパス（任意）

## 出力物
- `ocr.pid`: 実行中PID
- `ocr.status`: 稼働状況
- `ocr.latest.txt`: 最新OCR結果
- `screenshots/`: スクショ保存先
