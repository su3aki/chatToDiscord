# 概要

このプロジェクトは、LINE PC版のオープンチャット画面をOCRで取得し、Discord Webhookへ投稿します。
加えて、GUIで稼働状況の確認と安全停止ができます。

## 主な機能
1. LINEウィンドウのキャプチャとOCR
2. チャット欄のトリミング（CROP_RECT）
3. 画像前処理（2値化など）
4. Discord Webhook送信
5. GUIで開始/停止/状態表示
6. STOPファイルによる安全停止
7. スクショ出力と座標決定ツール

## 最小起動
1. `.env.example` を `.env` にコピー
2. `WEBHOOK_URL` を設定
3. `python ocrStart.py` または `python gui_controller.py` を実行

## 主要ファイル
- `ocrStart.py`: OCR本体とDiscord送信
- `gui_controller.py`: GUI制御
- `coord_picker.py`: クリックでCROP_RECTを出力
- `.env`: 実運用の設定
- `.env.example`: 設定テンプレート
# 注意点

1. CROP_MODE=screen のときは画面座標固定です。LINEウィンドウを別モニタに移動したら再設定してください。
2. マルチモニタでDPI倍率が違うとズレます。可能なら倍率を統一してください。
3. STOPファイルが残っていると即停止します（f:\Coding\discordBot\STOP）。
4. SAVE_SCREENSHOT_ONCE=true のときは1回で終了します。常時運用前にコメントアウト。
5. ONLY_ON_CHANGE=true だと同一内容は送信されません。動作確認中は false が分かりやすい。
6. 背景色で最適設定が変わります。白背景は INVERT=false、黒背景は INVERT=true 推奨。
7. LINEの表示変更（フォント/レイアウト）で精度が落ちることがあります。THRESHOLD と OCR_SCALE を再調整してください。
8. GUIはCtrl+Cで落とさないでください（KeyboardInterrupt）。×で閉じる。
9. TESSERACT_CMD が無効だとOCRが動きません。
