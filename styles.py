import docx
from docx.enum.style import WD_STYLE_TYPE


directory =		 'Styles'
docx_file_name = 'HSE_styles.docx'
styles_doc = docx.Document(f'{directory}/{docx_file_name}')

HSE_STYLES = styles_doc.styles
'''
Normal
Heading 1
Heading 2
Heading 3
List Paragraph
Before List 1 HSE
Table Name 1 HSE
Picture Name 1 HSE
Picture 1 HSE
Listing 1 HSE
'''
if __name__ == '__main__':
	paragraph_styles = [
		s for s in HSE_STYLES if s.type == WD_STYLE_TYPE.PARAGRAPH
	]
	for style in paragraph_styles:
		print(
			f'name: {style.name}',
			# f'font: {style.font}',
			# f'nps: {style.next_paragraph_style}',
			# f'pf: {style.paragraph_format}'
		)