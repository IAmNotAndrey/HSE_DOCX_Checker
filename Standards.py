from dataclasses import dataclass, field
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.shared import Pt, RGBColor, Length, Cm


#@dataclass
class Standards:
	'''Стандарты оформления ВКР, установленные ВШЭ'''
	#@dataclass
	class FontNames:
		""" Шрифты """
		standard: str = 'Times New Roman'
	
	#@dataclass
	class FontSizes:
		"""Кегли"""
		level_1_header: Length = Pt(16)
		level_2_header: Length = Pt(14)
		# standard: None | Length = Pt(13)
		standard: None | Length = None
		reduced: Length = Pt(12)

	#@dataclass
	class VerticalSpaces:
		'''Вертикальные интервалы'''
		# Интервалы перед/после (по вертикали)
		level_1_header: dict[str, None | Length] =	{'space_before': Pt(0),	'space_after': Pt(12)}
		level_2_header: dict[str, None | Length] =	{'space_before': Pt(12),'space_after': Pt(6)}
		level_3_header: dict[str, None | Length] =	{'space_before': Pt(8),	'space_after': Pt(4)}
		standard: 		dict[str, None | Length] = 	{'space_before': Pt(0),	'space_after': Pt(0)}
		table_name:		dict[str, None | Length] =	{'space_before': Pt(6),	'space_after': Pt(0)}
		table_text: 	dict[str, None | Length] =	{'space_before': Pt(2),	'space_after': Pt(2)}
		picture: 		dict[str, None | Length] = 	{'space_before': Pt(6),	'space_after': Pt(0)}
		picture_name: 	dict[str, None | Length] = 	{'space_before': Pt(0),	'space_after': Pt(6)}

		# Интервалы междустрочные
		line_spacing_single: 		 WD_LINE_SPACING =	WD_LINE_SPACING.SINGLE
		line_spacing_one_point_five: WD_LINE_SPACING =	WD_LINE_SPACING.ONE_POINT_FIVE

	#@dataclass
	class HorizontalSpaces:
		'''Горизонтальные интервалы'''
		# Отсутпы первой строки / слева / справа
		indents_standard: dict[str, None | Length] = {
		'first_line_indent': Cm(0),
		'left_indent': Cm(0),
		'right_indent': Cm(0)
		}
		indents_for_standard_size: dict[str, None | Length] = {
		'first_line_indent': Cm(1.25),
		'left_indent': Cm(0),
		'right_indent': Cm(0)
		}
		indents_list: dict[str, None | Length] = {
			'first_line_indent': Cm(1.5),
			'left_indent': Cm(2),
			'right_indent': Cm(0)
		}

	#@dataclass
	class Bolds:
		'''Полужирное начертание'''
		header: bool = True
		picture_name: bool = True
		standard: bool = False

	#@dataclass
	class Italics:
		'''Курсив'''
		standard: bool = False
		picture_name: bool = True

	#@dataclass
	class Underlines:
		'''Подчёркивание'''
		standard: bool = False

	#@dataclass
	class Strikes:
		'''Зачёркивания'''
		standard: bool = False

	#@dataclass
	class Colors:
		'''Цвет текста'''
		standard: RGBColor = RGBColor(0, 0, 0)

	#@dataclass
	class HighlightColors:
		'''Цвет заливки'''
		standard: RGBColor = WD_COLOR_INDEX.WHITE
		
	#@dataclass
	class KeepWithNext:
		'''Не отрывать от следующего абзаца'''
		header: bool = True
		standard: bool = False
		table_name: bool = True
		picture: bool = True
		picture_name: bool = False
					
	#@dataclass
	class KeepTogether:
		'''Не разрывать абзац'''
		standard: bool = False

	#@dataclass
	class PageBreakBefore:
		'''Абзац с новой страницы'''
		standard: bool = False

	#@dataclass
	class WidowControl:
		'''Запрет висячих строк'''
		standard: bool = True

	#@dataclass
	class ParagraphAlignments:
		header: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER
		standard: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.JUSTIFY

	#@dataclass
	class Levels:
		level_1_header: None = None
		level_2_header: None = None
		level_3_header: None = None
		standard: None = None
