# 📸 MK Image Converter Expert v1.0.1

複数画像の一括リサイズ、形式変換、透かし合成をシンプルかつパワフルに。
v1.0.1では、待望の**日本語透かし**と**PNGロゴ合成**に対応しました。



## ✨ 主な機能 / Key Features
- **ドラッグ&ドロップ**: フォルダから画像を放り込むだけで一括読み込み。
- **一括リサイズ & 変換**: 指定サイズへのリサイズ、PNG/JPEG/WebP等への一括変換。
- **高度な透かし (Watermark)**:
  - **日本語対応**: メイリオフォント採用により、日本語テキストも綺麗に合成。
  - **PNGロゴ対応**: お手持ちのロゴ画像を好きな位置に重ね合わせ可能。
- **設定プリセット**: よく使うリサイズ設定や保存名を自動保存。
- **多言語対応**: 日本語と英語をワンクリックで切り替え。

## 🚀 アップデート内容 (v1.0.1)
- ✅ **日本語テキスト透かしの文字化けを解消**
- ✅ **PNG画像によるロゴ透かし機能を追加**
- ✅ **インストーラー作成スクリプトの最適化**

## 📥 ダウンロード / Download
右側の [Releases] ページから最新のインストーラー（`.exe`）をダウンロードしてください。
面倒な環境構築なしで、すぐにWindowsでお使いいただけます。

[![GitHub release (latest by date)](https://img.shields.io/github/v/release/maakura/MK-Image-Converter?style=for-the-badge&color=orange)](https://github.com/maakura/MK-Image-Converter/releases)
[![GitHub All Releases](https://img.shields.io/github/downloads/maakura/MK-Image-Converter/total?style=for-the-badge&color=blue)](https://github.com/maakura/MK-Image-Converter/releases)

---

## 🛠 開発者向け (Build from source)
自身でビルドする場合は、以下のライブラリが必要です：
```bash
pip install Pillow tkinterdnd2
pyinstaller --onefile --noconsole --collect-all tkinterdnd2 --icon=icon.ico imageconverter.py
```

## 📄 ライセンス / License
このプロジェクトは MIT License のもとで公開されています。
個人・商用問わず、自由にご利用・改変・再配布いただけます。
Details are described in the LICENSE file.

Copyright (c) 2026 maakura
