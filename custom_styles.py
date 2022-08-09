from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from program import doc


if __name__ == "__main__":
	# создаем пользовательский стиль заголовка, с именем UserHead1
	style = doc.styles.add_style('UserHead1', WD_STYLE_TYPE.PARAGRAPH)
	style.font.name = 'Times New Roman'
	style.font.size = Pt(16)
	style.font.underline = False
	style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
