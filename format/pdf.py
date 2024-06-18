# coding: utf-8

# PDF ファイルのテキストを抽出するためのクラスを提供します。
import struct
import xml.etree.ElementTree as ET

class PDF:
    def read(self, file):
        # PDF ファイルのヘッダーを読み込む
        data = file.data
        
        pages = []
        # ページを読み込む
        while True:
            # ページの開始位置を検索
            start = data.find(b'BT')
            if start < 0:
                break
            data = data[start:]
            
            # ページの終了位置を検索
            end = data.find(b'ET')
            if end < 0:
                break
            data = data[:end]
            
            # ページのテキストを抽出
            text = self.extract_text(data)
            pages.append(text)
            
            # 次のページを読み込む
            data = data[end:]