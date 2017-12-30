# -*- coding: utf-8 -*-

from __future__ import division

import sys

import pandas as pd
import numpy as np

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_TABLE_DIRECTION
from docx.shared import RGBColor

from mspandas import style


class Handler():
	"""Handler with helpful methods to assist in creation of Microsoft Word documents

	Methods
	-------
	create_table(doc, df)
		Create a doc table using a pandas dataframe
	"""

	def create_table(self, doc, df,
					 style='Table Grid', section=None, overflow_margins=.5,
					 header=True, index=True, header_names=None, index_names=None,
					 column_totals=False, row_totals=False, column_totals_agg_map={}, row_totals_agg_map={}, column_totals_label='Total', row_totals_label='Total',
					 header_size=8, header_bold=True, header_italic=False, header_text_color=None, header_color=None, merge_header=None,
					 index_size=8, index_bold=False, index_italic=False, index_text_color=None,
					 totals_size=8, totals_bold=False, totals_italic=False, totals_text_color=None,
					 text_size=8, text_bold=False, text_italic=False, text_color=None, text_font_name=style.Font.name,
					 number_format='{:,.2f}', number_format_map=None,
					 numeric_cols_alignment=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER, char_cols_alignment=docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT, column_alignment_map={},
					 cell_margins='tight', banded_rows=False, row_height=None,
					 highlight_first_row=True, hightlight_first_col=False, highlight_last_row=False,
					 autofit=None, alignment=WD_TABLE_ALIGNMENT.CENTER, direction=WD_TABLE_DIRECTION.LTR,
					 encoding='utf-8', encoding_errors='strict'):
		"""Create a doc table using a pandas dataframe

		Parameters
		----------
		doc: docx.Document
			document object
		df: pd.DataFrame
			pandas dataframe object, can have pd.MultiIndex on either axis

		Returns
		-------
		table: docx.table
			docx table shape object, access table.Table for table object inside placeholder shape

		Keyword Arguements
		------------------
		style: str
			document style, default 'Table Grid', for help seee: http://python-docx.readthedocs.io/en/latest/user/styles-understanding.html
		section: docx.section.Section
			document section with custom page settings, default None (uses last section in document)
		overflow_margins:
			percent of margin width table width is allowed to overflow, default .8 (80%)
		header: bool
			whether or not to include header in table, default True
		index: bool
			whether or not to include index in table, default True
		column_totals: bool
			whether or not to sum columns and include a totals as last row, default False
		row_totals=False
			whether or not to sum rows and include a totals as last column, default False
		column_totals_agg_map: dict
			map of column names and aggregation method to be applied when totaling columns, default {}, example: {'a':'mean', 'b':'mean'}, will otherwise sum all data (WARNING: only works in pandas version >= 0.18)
		row_totals_agg_map: dict
			map of index names and aggregation method to be applied when totaling columns, default {}, example: {0:'mean', 1:'mean'}, will otherwise sum all data (WARNING: only works in pandas version >= 0.18)
		column_totals_label: str
			name of totals row in index, default 'Total'
		row_totals_label: str
			name of totals column in header, default 'Total'
		header_size: int
			header text size in ppt font size, default 8
		header_bold: bool
			whether or not to bold header text, default False
		header_italic: bool
			whether or not to italicize header text, default False
		header_text_color: list
			list of 3 RGB codes to color header text, default ...
		header_names: list
			custom df.columns.names, default None
		header_color: list
			list of 3 RGB codes to fill and color header row, default None
		merge_header: dict (or list of dicts)
			map of start and end columns (index or name) to be merged in header with optional level and alignment, default None, example: {'start':0, 'end':3, 'level':1, 'alignment':'center'} with required keys "start", "end", and optional keys "level","alignment"
		index_size: int
			index text size in ppt font size, default 8
		index_bold: bool
			whether or not to bold index text, default False
		index_italic: bool
			whether or not to italicize index text, default True
		index_text_color: list
			list of 3 RGB codes to color index text, default ...
		index_names: list
			custom df.index.names, default None
		totals_size: int
			totals text size in ppt font size, default 8
		totals_bold: bool
			whether or not to bold totals text, default False
		totals_italic: bool
			whether or not to italicize totals text, default False
		totals_color: list
			list of 3 RGB codes to color totals text, default
		text_size: int
			text size in ppt font size, default 8
		text_bold: bool
			whether or not to bold text, default False
		text_italic: bool
			whether or not to italicize text, default False
		text_color: list
			list of 3 RGB codes to color table text, default ...
		text_font_name: str
			string font name in Microsoft Word to overarchingly apply to all table text (e.g. 'Arial'), default pandas-msx.style.Font.name
		number_format: str
			python format string for formatting numerical data before converting to text per python-docx acceptance, default '{:0.2f}', for help see: https://www.python.org/dev/peps/pep-3101/
		number_format_map: dict
			sparse dictionary for mapping df column names (dict keys) to format strings (dict values) (e.g. '{:0.2f%}'), overrides number_format
		numeric_cols_alignment: docx.enum.text.WD_ALIGN_PARAGRAPH
			paragraph alignment for numeric data columns (values and header), default CENTER
		char_cols_alignment: docx.enum.text.WD_ALIGN_PARAGRAPH
			paragraph alignment for non-numeric data columns (values and header), default LEFT
		column_alignment_map: dict
			sparse dictionary for mapping df column names (dict keys) to paragraph alignments (dict values) (e.g. 'left', 'center', 'right'), overrides ~_cols_alignment
		cell_margins: str
			keyword for setting cell margin widths, keywords are adopted from ppt with a custom tight setting, default 'tight'
		row_height: float
			row height in inches, default None
		banded_rows: bool
			whether or not to alternate fill color of rows in light grey, default False
		highlight_first_row: bool
			whether or not to highlight first row, default True
		hightlight_first_col: bool
			whether or not to highlight first column, default True
		highlight_last_row: bool
			whether or not to highlight last row, default True
		autofit: bool
			whether or not to have word autofit table instead of our custom method, default False
		alignment: docx.enum.table.WD_TABLE_ALIGNMENT
			alignment of text in table, default CENTER
		direction: docx.enum.table.WD_TABLE_DIRECTION
			direction in which table columns are ordere (e.g. left to right, or right to left), default LTR

		Notes
		-----
		See http://python-docx.readthedocs.io/en/latest/api/table.html
		"""

		# ppt cell margins standards in inches
		margins_master = {
			'normal': {'top': 0.05, 'bottom': 0.05, 'left': 0.1, 'right': 0.1},
			'none': {'top': 0, 'bottom': 0, 'left': 0, 'right': 0},
			'narrow': {'top': 0.05, 'bottom': 0.05, 'left': 0.05, 'right': 0.05},
			'wide': {'top': 0.15, 'bottom': 0.15, 'left': 0.15, 'right': 0.15},
			# custom style, does not exist in ppt, is half of narrow
			'tight': {'top': 0.025, 'bottom': 0.025, 'left': 0.025, 'right': 0.025},
		}

		# convert column alignment map to pptx enum codes
		try:
			# python 2.7
			column_alignment_map = {k:docx.enum.text.WD_ALIGN_PARAGRAPH.__dict__[v.upper()] for k,v in column_alignment_map.iteritems()}
		except AttributeError:
			# python 3
			column_alignment_map = {k:docx.enum.text.WD_ALIGN_PARAGRAPH.__dict__[v.upper()] for k,v in column_alignment_map.items()}

		# total columns and concat with data as last row
		if column_totals:
			names = list(df.index.names)
			ordered_columns = list(df.columns)
			try:
				c_totals = df.fillna(0).agg({col:(column_totals_agg_map[col] if col in column_totals_agg_map else np.sum) for col in df.columns}).rename(column_totals_label)
			except:
				c_totals = []
				for col,method in column_totals_agg_map.items():
					t = df.fillna(0)[[col]].apply(method)
					c_totals = c_totals + [t]
				t = df.fillna(0)[[col for col in df.columns if col not in column_totals_agg_map.keys()]].sum()
				c_totals = c_totals + [t]
				c_totals = pd.concat(c_totals).rename(column_totals_label)
			c_totals = c_totals.to_frame().T
			# create multiindex if needed
			for i in range(df.index.nlevels-1):
			    c_totals['dummy'] = ' '
			    c_totals = c_totals.set_index('dummy',append=True)
			    c_totals.index.names = [None]*len(c_totals.index.names)
			try:
				df = pd.concat([df, c_totals], axis=0)
			except TypeError:
				# df index is categorical
				# add Total category and append
				df.index = df.index.add_categories(column_totals_label)
				df = pd.concat([df, c_totals], axis=0)
			df = df.reindex_axis(ordered_columns, axis=1)
			df.index.names = names
		# total rows and concat with data as last column
		if row_totals:
			ordered_index = list(df.index)
			try:
				r_totals = df.T.fillna(0).agg({col:(row_totals_agg_map[col] if col in row_totals_agg_map else np.sum) for col in df.T.columns}).rename(row_totals_label)
			except:
				totals = []
				for col,method in row_totals_agg_map.items():
					t = df.T.fillna(0)[[col]].apply(method)
					totals = totals + [t]
				t = df.T.fillna(0)[[col for col in df.T.columns if col not in row_totals_agg_map.keys()]].sum()
				totals = totals + [t]
				r_totals = pd.concat(totals).rename(row_totals_label)
			# TODO: Create multiindex if needed
			try:
				df = pd.concat([df, r_totals.to_frame()], axis=1)
			except TypeError:
				# df columns are categorical
				# add Total category and append
				df.columns = df.columns.add_categories(row_totals_label)
				df = pd.concat([df, r_totals.to_frame()], axis=1)
			df = df.reindex_axis(ordered_index, axis=0)
		# save list of column data types
		# accessed during dynamic formatting (e.g. paragraph alignment, column width calculations etc.)
		if isinstance(df.columns, pd.MultiIndex):
			numeric_cols = [(str(c1),str(c2)) for c1,c2 in df._get_numeric_data().columns]
			char_cols = [(str(c1),str(c2)) for c1,c2 in df.columns if not (str(c1),str(c2)) in numeric_cols]
		else:
			numeric_cols = [str(col) for col in df._get_numeric_data().columns]
			char_cols = [str(col) for col in df.columns if not str(col) in numeric_cols]
		# counts
		num_numeric_cols = len(numeric_cols) if len(numeric_cols) > 0 else 1
		num_char_cols = len(char_cols) if len(char_cols) > 0 else 1

		# convert numeric data to strings
		for col in df.columns:
			if col in df._get_numeric_data().columns:
				fmt = number_format
				if not number_format_map is None:
					try:
						fmt = number_format_map[col]
					except KeyError:
						# column was not specified in map
						pass
				df.loc[:,col] = df[col].fillna(0).apply(lambda x: fmt.format(x))
			else:
				# handle encoding for pptx intake
				# convert all to unicode for acceptance
				# values
				while True:
					try:
						if (sys.version_info < (3, 0)):
							df[col] = df[col].fillna('').apply(lambda s: unicode(s.encode(encoding), 'utf-8', errors=encoding_errors) if isinstance(s, unicode) else unicode(s, encoding))
						else:
							df[col] = df[col].fillna('').apply(lambda s: s.encode(encoding).decode('utf-8', errors=encoding_errors) if isinstance(s, str) else s.decode(encoding))
						break
					except (TypeError, AttributeError):
						df[col] = df[col].astype(str).fillna('')
						continue

		# handle encoding for docx intake
		# convert all to unicode for acceptance
		# columns
		names = df.columns.names
		multi = []
		for level in range(df.columns.nlevels):
			values = []
			for s in df.columns.get_level_values(level):
				try:
					if (sys.version_info < (3, 0)):
						values.append(unicode(s.encode(encoding), 'utf-8', errors=encoding_errors) if isinstance(s, unicode) else unicode(s, encoding))
					else:
						values.append(s.encode(encoding).decode('utf-8', errors=encoding_errors) if isinstance(s, str) else s.decode(encoding))
				except (TypeError, AttributeError):
					# value is numeric
					values.append(str(s))
					pass
			if isinstance(df.columns, pd.MultiIndex):
				multi += [values]
				if level == df.columns.nlevels - 1:
					df.columns = pd.MultiIndex.from_arrays(multi)
			else:
				df.columns = values
		df.columns.names = names
		# indices
		names = df.index.names
		multi = []
		for level in range(df.index.nlevels):
			values = []
			for s in df.index.get_level_values(level):
				try:
					if (sys.version_info < (3, 0)):
						values.append(unicode(s.encode(encoding), 'utf-8', errors=encoding_errors) if isinstance(s, unicode) else unicode(s, encoding))
					else:
						values.append(s.encode(encoding).decode('utf-8', errors=encoding_errors) if isinstance(s, str) else s.decode(encoding))
				except (TypeError, AttributeError):
					# value is numeric
					values.append(str(s))
					pass
			if isinstance(df.index, pd.MultiIndex):
				multi += [values]
				if level == df.index.nlevels - 1:
					df.index = pd.MultiIndex.from_arrays(multi)
			else:
				df.index = values
			df.index.names = names

		# add custom index names
		if not index_names is None:
			for i,name in enumerate(index_names):
				i = None if not isinstance(df.index, pd.MultiIndex) else i
				df.index = df.index.set_names(name, level=i)

		# add custom header names
		if not header_names is None:
			for i,name in enumerate(header_names):
				i = None if not isinstance(df.columns, pd.MultiIndex) else i
				df.columns = df.columns.set_names(name, level=i)

		# define table dimensions
		num_rows = len(df)
		num_cols = len(df.columns)

		if header:
			num_rows += df.columns.nlevels
		if index:
			num_cols += df.index.nlevels

		# insert table into shape
		table = doc.add_table(rows=num_rows, cols=num_cols)

		if section is None:
			section = doc.sections[-1]

		# save width of document section page from template
		table_width = section.page_width - (section.left_margin*(1-overflow_margins)) - (section.right_margin*(1-overflow_margins))

		# style
		table.style = style

		# alignment
		table.alignment = alignment

		# autofit
		if not autofit is None:
			table.allow_autofit = True
			table.autofit = True
			# TODO: FIGURE THIS OUW
			#how = 0 if autofit == 'window' else 1 if autofit == 'content' else 2

		#table direction
		table.table_direction = direction

		# convert colors to docx RGB
		if not header_text_color is None:
			header_text_color = RGBColor(*header_text_color)
		if not index_text_color is None:
			index_text_color = RGBColor(*index_text_color)
		if not totals_text_color is None:
			totals_text_color = RGBColor(*totals_text_color)
		if not text_color is None:
			text_color = RGBColor(*text_color)

		# add header to table
		if header:
			for level in range(df.columns.nlevels):
				for i in range(num_cols):
					c = table.cell(level,i)
					label = ' '
					col = None
					if index and i <= df.index.nlevels-1:
						if any(name is not None for name in df.columns.names):
							label = df.columns.names[level]
						else:
							label = ' '
					elif index and i < df.index.nlevels:
						continue
					elif index and i > df.index.nlevels-1:
						col = i - df.index.nlevels
						label = df.columns.get_level_values(level)[col]
					else:
						col = i
						label = df.columns.get_level_values(level)[col]
					c.text = label or ' '
					c.margin_top = docx.shared.Inches(margins_master[cell_margins]['top'])
					c.margin_bottom = docx.shared.Inches(margins_master[cell_margins]['bottom'])
					c.margin_left = docx.shared.Inches(margins_master[cell_margins]['left'])
					c.margin_right = docx.shared.Inches(margins_master[cell_margins]['right'])
					if header_color is not None:
						rgb = RGBColor(*header_color)
						xml_shd = docx.oxml.parse_xml(r'<w:shd {} w:fill="{}"/>'.format(docx.oxml.ns.nsdecls('w'), rgb))
						c._tc.get_or_add_tcPr().append(xml_shd)
					p = c.paragraphs[0]
					if col is not None:
						if df.columns[col] in column_alignment_map:
							p.alignment = column_alignment_map[df.columns[col]]
						else:
							if df.columns[col] in numeric_cols:
								p.alignment = numeric_cols_alignment
							elif df.columns[col] in char_cols:
								p.alignment = char_cols_alignment
					try:
						r = p.runs[0]
						r.font.bold = header_bold
						r.font.italic = header_italic
						r.font.size = docx.shared.Pt(header_size)
						if not header_text_color is None:
							r.font.color.rgb = header_text_color
						r.font.name = text_font_name
					except IndexError:
						# mysteriously no paragraph / run exists
						pass

		# merge header cells
		if header:
			if not merge_header is None:
				iterator = [merge_header] if not isinstance(merge_header,list) else merge_header
				for merge_header in iterator:
					level = merge_header['level'] if isinstance(df.columns,pd.MultiIndex) and 'level' in merge_header.keys() else 0
					offset = df.index.nlevels if index else 0
					start = merge_header['start'] if isinstance(merge_header['start'],int) else df.columns.get_loc(merge_header['start'])
					start_cell = table.cell(level,start+offset)
					end = merge_header['end'] if isinstance(merge_header['end'],int) else df.columns.get_loc(merge_header['end'])
					end_cell = table.cell(level,end+offset)
					c = start_cell.merge(end_cell)
					# cell texts were concatenated with \n, remove any cell text which was None or empty string
					c.text = c.text.replace('\nnan','').replace('\n','').strip()
					# apply formatting
					c.margin_top = docx.shared.Inches(margins_master[cell_margins]['top'])
					c.margin_bottom = docx.shared.Inches(margins_master[cell_margins]['bottom'])
					c.margin_left = docx.shared.Inches(margins_master[cell_margins]['left'])
					c.margin_right = docx.shared.Inches(margins_master[cell_margins]['right'])
					if header_color is not None:
						rgb = RGBColor(*header_color)
						xml_shd = docx.oxml.parse_xml(r'<w:shd {} w:fill="{}"/>'.format(docx.oxml.ns.nsdecls('w'), rgb))
						c._tc.get_or_add_tcPr().append(xml_shd)
					p = c.paragraphs[0]
					if 'alignment' in merge_header.keys():
						p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.__dict__[merge_header['alignment'].upper()]
					try:
						r = p.runs[0]
						r.font.bold = header_bold
						r.font.italic = header_italic
						r.font.size = docx.shared.Pt(header_size)
						if not header_text_color is None:
							r.font.color.rgb = header_text_color
						r.font.name = text_font_name
					except IndexError:
						# mysteriously no paragraph / run exists
						pass

		# add index to table
		if index:
			for level in range(df.index.nlevels):
				for i in range(num_rows):
					c = table.cell(i,level)
					row = None
					keep_header_formatting = False
					if header and i == df.columns.nlevels-1:
						if any(name is not None for name in df.index.names):
							label = df.index.names[level]
							keep_header_formatting = True
						else:
							if header:
								continue
							else:
								label = ' '
					elif header and i < df.columns.nlevels:
						continue
					elif header and i >= df.columns.nlevels:
						row = i - df.columns.nlevels
						label = df.index.get_level_values(level)[row]
					else:
						row = i
						label = df.index.get_level_values(level)[row]
					c.text = label or ' '
					c.margin_top = docx.shared.Inches(margins_master[cell_margins]['top'])
					c.margin_bottom = docx.shared.Inches(margins_master[cell_margins]['bottom'])
					c.margin_left = docx.shared.Inches(margins_master[cell_margins]['left'])
					c.margin_right = docx.shared.Inches(margins_master[cell_margins]['right'])
					if banded_rows:
						if not keep_header_formatting:
							rgb = RGBColor(*style.RGB.grey_light) if (i-df.columns.nlevels) % 2 == 0 else RGBColor(*style.RGB.grey_light2)
							xml_shd = docx.oxml.parse_xml(r'<w:shd {} w:fill="{}"/>'.format(docx.oxml.ns.nsdecls('w'), rgb))
							c._tc.get_or_add_tcPr().append(xml_shd)
					try:
						r = c.paragraphs[0].runs[0]
						if not keep_header_formatting:
							r.font.bold = index_bold
							r.font.italic = index_italic
							r.font.size = docx.shared.Pt(index_size)
							if not index_text_color is None:
								r.font.color.rgb = index_text_color
							r.font.name = text_font_name
						else:
							r.font.bold = header_bold
							r.font.italic = header_italic
							r.font.size = docx.shared.Pt(header_size)
							if not header_text_color is None:
								r.font.color.rgb = header_text_color
							r.font.name = text_font_name
					except IndexError:
						# mysteriously no paragraph / run exists
						pass

		# iterate thru dataframe matrix and add data table, cell by cell
		# impute any missing data as empty string
		mat = df.fillna(' ').as_matrix()
		for row in range(df.shape[0]):
			for col in range(df.shape[1]):
				if header and index:
					c = table.cell(row+df.columns.nlevels,col+df.index.nlevels)
				elif header and not index:
					c = table.cell(row+df.columns.nlevels,col)
				elif index and not header:
					c = table.cell(row,col+df.index.nlevels)
				else:
					c = table.cell(row,col)
				c.text = mat[row,col] or ' '
				# alternative accessor
				#c.text = df.loc[df.index[row], df.columns[col]]
				c.margin_top = docx.shared.Inches(margins_master[cell_margins]['top'])
				c.margin_bottom = docx.shared.Inches(margins_master[cell_margins]['bottom'])
				c.margin_left = docx.shared.Inches(margins_master[cell_margins]['left'])
				c.margin_right = docx.shared.Inches(margins_master[cell_margins]['right'])
				if banded_rows:
					rgb = RGBColor(*style.RGB.grey_light) if row % 2 == 0 else RGBColor(*style.RGB.grey_light2)
					xml_shd = docx.oxml.parse_xml(r'<w:shd {} w:fill="{}"/>'.format(docx.oxml.ns.nsdecls('w'), rgb))
					c._tc.get_or_add_tcPr().append(xml_shd)
				p = c.paragraphs[0]
				if df.columns[col] in column_alignment_map:
					p.alignment = column_alignment_map[df.columns[col]]
				else:
					if df.columns[col] in numeric_cols:
						p.alignment = numeric_cols_alignment
					elif df.columns[col] in char_cols:
						p.alignment = char_cols_alignment
				try:
					r = p.runs[0]
					r.font.bold = text_bold
					r.font.italic = text_italic
					r.font.size = docx.shared.Pt(text_size)
					if not text_color is None:
						r.font.color.rgb = text_color
					r.font.name = text_font_name
				except IndexError:
					# mysteriously no paragraph / run exists
					pass

		# format totals
		if column_totals:
			for i in range(num_cols-df.index.nlevels):
				if index:
					c = table.cell(num_rows-1,i+df.index.nlevels)
				else:
					c = table.cell(num_rows-1,i)
				r = c.paragraphs[0].runs[0]
				r.font.bold = totals_bold
				r.font.italic = totals_italic
				r.font.size = docx.shared.Pt(totals_size)
				if not totals_text_color is None:
					r.font.color.rgb = totals_text_color
				r.font.name = text_font_name
		if row_totals:
			for i in range(num_rows-df.columns.nlevels):
				if header:
					c = table.cell(i+df.columns.nlevels,num_cols-1)
				else:
					c = table.cell(i,num_cols-1)
				r = c.paragraphs[0].runs[0]
				r.font.bold = totals_bold
				r.font.italic = totals_italic
				r.font.size = docx.shared.Pt(totals_size)
				if not totals_text_color is None:
					r.font.color.rgb = totals_text_color
				r.font.name = text_font_name

		# customize table row hieghts
		if not row_height is None:
			emu = row_height * docx.shared.Length._EMUS_PER_INCH
			for r in table.rows:
				r.height = docx.shared.Emu(round(emu))

		if autofit is None:
			# customize table column widths
			min_col_w = 4 #cm
			max_col_w = table_width / docx.shared.Length._EMUS_PER_CM / 2 # num_char_cols (don't hog the table)
			w_columns = 0
			for ix,col in enumerate(table.columns):
				if index and ix < df.index.nlevels:
					# compute width dynamically based on max text size in index, proportional to text size
					cm = min( max( len(max([s for s in df.index.get_level_values(ix)], key=len)) * 2 * (1/index_size), min_col_w), max_col_w)
					emu = docx.shared.Cm(int(np.ceil(cm)))
					col.width = emu
					for c in col.cells:
						c.width = emu
					w_columns += col.width
					continue
				elif index:
					# adjust ppt table column index for dataframe indexing
					df_ix = ix - df.index.nlevels
				else:
					# no index in ppt table, columns line up
					df_ix = ix
				# compute width dynamically based on max text size in column, proportional to text size
				cm = min( max( len(max([s for s in df.loc[:,df.columns[df_ix]]], key=len)) * 2 * (1/text_size), min_col_w), max_col_w)
				emu = docx.shared.Cm(int(np.ceil(cm)))
				col.width = emu
				for c in col.cells:
					c.width = emu
				w_columns += col.width

			# check if column widths overflow template table width
			if w_columns > table_width:
				# compute overflow and adjust columns proportionally
				w_overflow = w_columns - table_width
				w_columns_over = w_columns
				w_columns = 0
				for ix,col in enumerate(table.columns):
					# compute col's percent of table size
					percent_of_table = col.width / w_columns_over
					# compute factor of overflow as reduction amount
					reduction_amount = percent_of_table * w_overflow
					# reduce
					emu = col.width - docx.shared.Emu(round(reduction_amount))
					col.width = emu
					for c in col.cells:
						c.width = emu
					w_columns += col.width

			# check if column widths do not fill template table width
			if w_columns < table_width:
				# compute deficit and adjust columns proportionally
				w_deficit = table_width - w_columns
				w_columns_under = w_columns
				w_columns = 0
				for ix,col in enumerate(table.columns):
					# compute col's percent of table size
					percent_of_table = col.width / w_columns_under
					# compute factor of deficit as extension amount
					extension_amount = percent_of_table * w_deficit
					# extend
					emu = col.width + docx.shared.Emu(round(extension_amount))
					col.width = emu
					for c in col.cells:
						c.width = emu
					w_columns += col.width

		# highlight rows, or columns
		table.first_row = highlight_first_row
		table.first_col = hightlight_first_col
		table.last_row = highlight_last_row

		return table
