#!/bin/bash
# Download Noto Sans JP font for Unicode support (Japanese + Vietnamese)
echo "Downloading Noto Sans JP font..."
curl -L -o NotoSansJP-Regular.ttf \
  "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf"
echo "Done! Font size: $(du -h NotoSansJP-Regular.ttf | cut -f1)"
