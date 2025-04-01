# pip install streamlit python-pptx

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
import re

def markdown_to_pptx(markdown_text):
    # 新しいプレゼンテーションを作成
    prs = Presentation()
    
    # マークダウンの内容をスライドごとに分割（---でスライドを区切る）
    slides_content = re.split(r'^---\s*$', markdown_text, flags=re.MULTILINE)
    
    for content in slides_content:
        if not content.strip():
            continue
            
        # スライドを追加（タイトルとコンテンツのレイアウト）
        slide_layout = prs.slide_layouts[1]  # タイトルとコンテンツ
        slide = prs.slides.add_slide(slide_layout)
        
        # スライドのタイトルとコンテンツを抽出
        lines = content.split('\n')
        title = ""
        body = []
        
        for line in lines:
            if line.startswith('# '):
                title = line[2:].strip()
            else:
                if line.strip():
                    body.append(line.strip())
        
        # タイトルを設定
        if title:
            slide.shapes.title.text = title
        else:
            slide.shapes.title.text = "New Slide"
        
        # コンテンツを設定
        if body:
            content_shape = slide.placeholders[1]
            tf = content_shape.text_frame
            tf.text = ""  # デフォルトのテキストをクリア
            
            for line in body:
                # 箇条書きの処理
                if line.startswith('- '):
                    p = tf.add_paragraph()
                    p.text = line[2:].strip()
                    p.level = 0
                elif line.startswith('  - '):
                    p = tf.add_paragraph()
                    p.text = line[4:].strip()
                    p.level = 1
                else:
                    p = tf.add_paragraph()
                    p.text = line
    
    return prs

def save_pptx(prs):
    # PowerPointファイルをバイナリデータとして保存
    file_stream = BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    return file_stream

# StreamlitアプリのUI
st.title('マークダウンからPowerPointへ変換')
st.write('マークダウン形式のテキストを入力し、PowerPointファイルとしてダウンロードできます。')
st.write('スライドは「---」で区切ってください。見出し（#）がスライドのタイトルになります。')

# テキスト入力エリア
default_markdown = """# 最初のスライド

これは最初のスライドの内容です。
- 箇条書き1
- 箇条書き2

---
# 2番目のスライド

- トップレベル項目
  - サブ項目
  - サブ項目
- 別のトップレベル項目
"""

markdown_text = st.text_area(
    'マークダウンテキストを入力してください',
    height=300,
    value=default_markdown
)

# 変換ボタン
if st.button('PowerPointファイルに変換'):
    if markdown_text:
        try:
            prs = markdown_to_pptx(markdown_text)
            pptx_file = save_pptx(prs)
            
            st.success('変換が完了しました！')
            st.download_button(
                label='PowerPointファイルをダウンロード',
                data=pptx_file,
                file_name='presentation.pptx',
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        except Exception as e:
            st.error(f'エラーが発生しました: {e}')
    else:
        st.warning('マークダウンテキストを入力してください')