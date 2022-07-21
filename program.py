import docx
import app_logger
from Standards import Standards

'''
TODO:
	- Указывать на несоответствие ФИО документа с ФИО пользователя
	- Сделать отдельные параметры отступа для марк-х и нум-х списков
	- Определять названия таблиц и рисунков 12 кегля и правильно их оформлять
README:
	- Редактор анализирует текст на основе размера Кегля, поэтому, пожалуйста, убедитесь, что в начале каждого параграфа хотя бы у одного символа стоит необходимый кегль. 
	- Пожалуйста не пишите заголовки с нумерацией так, чтобы между ними был таб, только пробел
'''


directory =		 'docx'
docx_file_name = 'ex2.docx'
out_docx_file_name = 'out.docx'
doc = docx.Document(f'{directory}/{docx_file_name}')

# logger = app_logger.get_logger(__name__)

for i, paragraph in enumerate(doc.paragraphs):

	try:
		first_run = paragraph.runs[0]
	# В некоторых местах Run вообще нету
	except IndexError:
		continue

	fmt = paragraph.paragraph_format
	# OPTIMIZE: может стоить сделать всё с if-ами? Тогда не возникнет проблем с логированием
	
	# Оформление параграфа: здесь установлены те параметры, которые встречаются чаще всего, если потом их нужно изменить, то в match они переназначаются
	fmt.widow_control = Standards.WidowControl.standard
	fmt.keep_with_next = Standards.KeepWithNext.header
	fmt.keep_together = Standards.KeepTogether.standard
	fmt.page_break_before = Standards.PageBreakBefore.standard

	fmt.alignment = Standards.ParagraphAlignments.header
	# Уровень
	
	fmt.left_indent =		Standards.HorizontalSpaces.indents_standard['left_indent']
	fmt.right_indent = Standards.HorizontalSpaces.indents_standard['right_indent']
	fmt.first_line_indent =	Standards.HorizontalSpaces.indents_standard['first_line_indent']

	fmt.space_before = Standards.VerticalSpaces.standard['space_before']
	fmt.space_after = Standards.VerticalSpaces.standard['space_after']
	fmt.line_spacing_rule = Standards.VerticalSpaces.line_spacing_single

	# Оформление 1-го Run: здесь установлены те параметры, которые встречаются чаще всего, если потом их нужно изменить, то в match они переназначаются 
	first_run.font.name = Standards.FontNames.standard 
	# NOTE: bold нельзя здесь задавать, тк это приведёт к ошибкам: заголовок 3-го уровня определяется по bold
	first_run.font.italic = Standards.Italics.standard
	first_run.font.underline = Standards.Underlines.standard
	first_run.font.strike = Standards.Strikes.standard
	first_run.font.color.rgb = Standards.Colors.standard
	first_run.font.highlight_color = Standards.HighlightColors.standard

	# Проверка параграфа и первого Run по параметрам, которые отличаются в зависимости от кегля
	# TODO: везде написать изменение уровней параграфов
	match first_run.font.size:
		case Standards.FontSizes.level_1_header:
			# Уровень
			
			fmt.space_before = Standards.VerticalSpaces.level_1_header['space_before']
			fmt.space_after = Standards.VerticalSpaces.level_1_header['space_after']

			first_run.font.bold = Standards.Bolds.header

		case Standards.FontSizes.level_2_header:
			# Уровень

			fmt.space_before = Standards.VerticalSpaces.level_2_header['space_before']
			fmt.space_after = Standards.VerticalSpaces.level_2_header['space_after']
			
			first_run.font.bold = Standards.Bolds.header
			
		case Standards.FontSizes.standard:
			# Уровень

			
			# Если 1-й Run стандартного кегля и с полужирным начертанием
			if first_run.bold == Standards.Bolds.header:
				fmt.alignment = Standards.ParagraphAlignments.header

				fmt.left_indent =		Standards.HorizontalSpaces.indents_standard['left_indent']
				# Правый отступ у всех стандартный
				fmt.first_line_indent =	Standards.HorizontalSpaces.indents_standard['first_line_indent']
				
				fmt.space_before = Standards.VerticalSpaces.level_3_header['space_before']
				fmt.space_after = Standards.VerticalSpaces.level_3_header['space_after']

				first_run.font.bold = Standards.Bolds.header

			else:
				fmt.keep_with_next = Standards.KeepWithNext.standard

				fmt.left_indent =		Standards.HorizontalSpaces.indents_for_standard_size['left_indent']
				# Правый отступ у всех стандартный
				fmt.first_line_indent =	Standards.HorizontalSpaces.indents_for_standard_size['first_line_indent']

				fmt.alignment = Standards.ParagraphAlignments.standard
				
				fmt.space_before = Standards.VerticalSpaces.standard['space_before']
				fmt.space_after = Standards.VerticalSpaces.standard['space_after']
				fmt.line_spacing_rule = Standards.VerticalSpaces.line_spacing_one_point_five

				first_run.font.bold = Standards.Bolds.standard
				
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
