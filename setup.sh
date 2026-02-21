#!/bin/bash
# Download Noto Sans JP font (static TTF) for Unicode support (Japanese + Vietnamese)
echo "Downloading Noto Sans JP font (static TTF from fontsource)..."
curl -L -o NotoSansJP-Regular.ttf \
  "https://cdn.jsdelivr.net/fontsource/fonts/noto-sans-jp@latest/japanese-400-normal.ttf"
echo "Done! Font size: $(du -h NotoSansJP-Regular.ttf | cut -f1)"
echo "Font type: $(file NotoSansJP-Regular.ttf)"
