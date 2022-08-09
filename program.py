import docx
import app_logger
from Standards import Standards
from styles import HSE_STYLES
from docx.enum.style import WD_STYLE_TYPE

'''
TODO:
	- Указывать на несоответствие ФИО документа с ФИО пользователя
	- Сделать отдельные параметры отступа для марк-х и нум-х списков
	- Определять названия таблиц и рисунков 12 кегля и правильно их оформлять
	- Сделать костыль для установки уровня параграфа? Потому что в docx такой возможности нет
README:
	- Редактор анализирует текст на основе размера Кегля, поэтому, пожалуйста, убедитесь, что в начале каждого параграфа хотя бы у одного символа стоит необходимый кегль. 
	- Пожалуйста не пишите заголовки с нумерацией так, чтобы между ними был таб, только пробел
'''

directory =		 'docx'
docx_file_name = 'ex2.docx'
out_docx_file_name = 'out.docx'
doc = docx.Document(f'{directory}/{docx_file_name}')

# logger = app_logger.get_logger(__name__)

if __name__ == '__main__':
	for i, paragraph in enumerate(doc.paragraphs):
		first_run = None
		try:
			first_run = paragraph.runs[0]
		# В некоторых параграфах Run вообще нету
		except IndexError:
			continue

		match first_run.font.size:
			case Standards.FontSizes.level_1_header:
				paragraph.style = HSE_STYLES['Heading 1']

			case Standards.FontSizes.level_2_header:
				paragraph.style = HSE_STYLES['Heading 2']

			case Standards.FontSizes.standard:
				# Если 1-й Run стандартного кегля и с полужирным начертанием
				if first_run.bold == Standards.Bolds.header:
					paragraph.style = HSE_STYLES['Heading 3']

				else:
					paragraph.style = HSE_STYLES['Normal']

			case Standards.FontSizes.reduced:
				...
			case _:
				...
		
		for i2, run in enumerate(paragraph.runs):
			# Пропускаем 1-й Run, тк все его параметры мы установили выше
			if i2 == 0:
				continue
			
			# Все остальные Run-ы делаем полностью идентичными 1-му
			run.font.name = 			first_run.font.name
			run.font.size = 			first_run.font.size
			run.font.bold = 			first_run.font.bold
			run.font.italic = 			first_run.font.italic
			run.font.underline = 		first_run.font.underline
			run.font.strike = 			first_run.font.strike
			run.font.color.rgb = 		first_run.font.color.rgb
			run.font.highlight_color =	first_run.font.highlight_color


	doc.save(f'{directory}/{out_docx_file_name}')
