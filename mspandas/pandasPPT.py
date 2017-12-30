# -*- coding: utf-8 -*-

from __future__ import division

import sys

import pandas as pd
import numpy as np

import pptx
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor

from mspandas import style

class Handler():
	"""Handler with helpful methods to assist in creation of Microsoft PowerPoint Documents.

	Methods
	-------
	map_layouts(ppt)
		Create dictionary object of template layouts in slide master from ppt object, where keys are layout names.
	map_shapes(layout)
		Create dictionary object of slide shapes in template layout from layout object, where keys are shape names.
	create_table(table, df)
		Create a ppt table using a pandas dataframe
	create_chart(chart, df)
		Create a ppt chart using a pandas dataframe
	"""

	def map_layouts(self, ppt, verbose=False):
		"""Create dictionary object of template layouts in slide master from ppt object, where keys are layout names.

		Parameters
		----------
		ppt: ppt.Presentation
			Powerpoint presentation object.

		Returns
		-------
		layout_map: dict
			Dictionary of ppt layout objects where keys are layout names from slide master.
		"""
		layout_map = {}
		for slide in ppt.slide_layouts:
			layout_map[slide.name] = slide
			if verbose:
				print(slide.name)
		return layout_map

	def map_shapes(self, layout, verbose=False):
		"""Create dictionary object of slide shapes in template layout from layout object, where keys are shape names.

		Parameters
		----------
		layout: ppt.slide.SlideLayout
			Slide layout object.

		Returns
		-------
		shape_map: dict
			Dictionary of slide shape objects where keys are shape names from template layout.
		"""
		shape_map = {}
		for shape in layout.shapes:
			if shape.is_placeholder:
				phf = shape.placeholder_format
				shape_map[shape.name] = phf.idx
				if verbose:
					print('{} index: {}, type: {}'.format(shape.name, phf.idx, phf.type))
		return shape_map

	def create_table(self, table, df,
					 header=True, index=True, header_names=None, index_names=None,
					 column_totals=False, row_totals=False, column_totals_agg_map={}, row_totals_agg_map={}, column_totals_label='Total', row_totals_label='Total',
					 header_size=9, header_bold=True, header_italic=False, header_text_color=None, header_color=None, merge_header=None,
					 index_size=9, index_bold=False, index_italic=False, index_text_color=None,
					 totals_size=9, totals_bold=False, totals_italic=False, totals_text_color=None,
					 text_size=9, text_bold=False, text_italic=False, text_color=None, text_font_name=style.Font.name,
					 #auto_size=pptx.enum.text.MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE, fit_text=False,
					 number_format='{:,.2f}', number_format_map=None,
					 numeric_cols_alignment=pptx.enum.text.PP_ALIGN.CENTER, char_cols_alignment=pptx.enum.text.PP_ALIGN.LEFT, column_alignment_map={},
					 cell_margins='tight', banded_rows=True, row_height=.15,
					 highlight_first_row=True, hightlight_first_col=False, highlight_last_row=False,
					 encoding='utf-8', encoding_errors='strict'):
		"""Create a ppt table using a pandas dataframe

		Parameters
		----------
		table: pptx.shapes.table
			pptx shape object of type table, graphicframe placeholder
		df: pd.DataFrame
			pandas dataframe object, can have pd.MultiIndex on either axis

		Returns
		-------
		table: pptx.shapes.table
			pptx table shape object, access table.Table for table object inside placeholder shape

		Keyword Arguements
		------------------
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
			header text size in ppt font size, default 9
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
			index text size in ppt font size, default 9
		index_bold: bool
			whether or not to bold index text, default False
		index_italic: bool
			whether or not to italicize index text, default True
		index_text_color: list
			list of 3 RGB codes to color index text, default ...
		index_names: list
			custom df.index.names, default None
		totals_size: int
			totals text size in ppt font size, default 9
		totals_bold: bool
			whether or not to bold totals text, default False
		totals_italic: bool
			whether or not to italicize totals text, default False
		totals_color: list
			list of 3 RGB codes to color totals text, default
		text_size: int
			text size in ppt font size, default 9
		text_bold: bool
			whether or not to bold text, default False
		text_italic: bool
			whether or not to italicize text, default False
		text_color: list
			list of 3 RGB codes to color table text, default ...
		text_font_name: str
			string font name in Microsoft Word to overarchingly apply to all table text (e.g. 'Arial'), default pandas-msx.style.Font.name
		number_format: str
			python format string for formatting numerical data before converting to text per python-pptx acceptance, default '{:0.2f}', for help see: https://www.python.org/dev/peps/pep-3101/
		number_format_map: dict
			sparse dictionary for mapping df column names (dict keys) to format strings (dict values) (e.g. '{:0.2f%}'), overrides number_format
		numeric_cols_alignment: pptx.enum.text.PP_ALIGN
			paragraph alignment for numeric data columns (values and header), default CENTER
		char_cols_alignment: pptx.enum.text.PP_ALIGN
			paragraph alignment for non-numeric data columns (values and header), default LEFT
		column_alignment_map: dict
			sparse dictionary for mapping df column names (dict keys) to paragraph alignments (dict values) (e.g. 'left', 'center', 'right'), overrides ~_cols_alignment
		cell_margins: str
			keyword for setting cell margin widths, keywords are adopted from ppt with a custom tight setting, default 'tight'
		row_height: float
			row height in inches, default .15
		highlight_first_row: bool
			whether or not to highlight first row, default True
		hightlight_first_col: bool
			whether or not to highlight first column, default True
		highlight_last_row: bool
			whether or not to highlight last row, default True

		Notes
		-----
		See http://python-pptx.readthedocs.io/en/latest/api/table.html
		"""

		# ppt cell margins standards in inches
		margins_master = {
			'normal': {'top': 0.05, 'bottom': 0.05, 'left': 0.1, 'right': 0.1},
			'none': {'top': 0, 'bottom': 0, 'left': 0, 'right': 0},
			'narrow': {'top': 0.05, 'bottom': 0.05, 'left': 0.05, 'right': 0.05},
			'wide': {'top': 0.15, 'bottom': 0.15, 'left': 0.15, 'right': 0.15},
			# custom style, does not exist in ppt
			'tight': {'top': 0.025, 'bottom': 0.025, 'left': 0.025, 'right': 0.025},
		}

		# convert column alignment map to pptx enum codes
		try:
			# python 2.7
			column_alignment_map = {k:pptx.enum.text.PP_ALIGN.__dict__[v.upper()] for k,v in column_alignment_map.iteritems()}
		except AttributeError:
			# python 3
			column_alignment_map = {k:pptx.enum.text.PP_ALIGN.__dict__[v.upper()] for k,v in column_alignment_map.items()}

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

		# handle encoding for pptx intake
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
		table_shape = table.insert_table(rows=num_rows, cols=num_cols)

		# get table object from graphic frame
		table = table_shape.table

		# save desired width of table shape from template
		table_width = table_shape.width

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
					c.margin_top = pptx.util.Inches(margins_master[cell_margins]['top'])
					c.margin_bottom = pptx.util.Inches(margins_master[cell_margins]['bottom'])
					c.margin_left = pptx.util.Inches(margins_master[cell_margins]['left'])
					c.margin_right = pptx.util.Inches(margins_master[cell_margins]['right'])
					if not header_color is None:
						c.fill.solid()
						c.fill.fore_color.rgb = RGBColor(*header_color)
					tf = c.text_frame
					p = tf.paragraphs[0]
					if col is not None:
						if df.columns[col] in column_alignment_map:
							p.alignment = column_alignment_map[df.columns[col]]
						else:
							if df.columns[col] in numeric_cols:
								p.alignment = numeric_cols_alignment
							if df.columns[col] in char_cols:
								p.alignment = char_cols_alignment
					try:
						r = p.runs[0]
						r.font.bold = header_bold
						r.font.italic = header_italic
						r.font.size = pptx.util.Pt(header_size)
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
					start = merge_header['start']+offset if isinstance(merge_header['start'],int) else df.columns.get_loc(merge_header['start'])+offset
					end = merge_header['end']+offset if isinstance(merge_header['end'],int) else df.columns.get_loc(merge_header['end'])+offset
					length = end - start + 1
					cells = [c for c in table.rows[level].cells][start:end]
					cells[0]._tc.set('gridSpan', str(length))
					for c in cells[1:]:
						c._tc.set('hMerge', '1')
					# apply formatting
					tf = cells[0].text_frame
					p = tf.paragraphs[0]
					if 'alignment' in merge_header.keys():
						p.alignment = pptx.enum.text.PP_ALIGN.__dict__[merge_header['alignment'].upper()]

		# add index to table
		if index:
			for level in range(df.index.nlevels):
				for i in range(num_rows):
					c = table.cell(i,level)
					label = ' '
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
					c.margin_top = pptx.util.Inches(margins_master[cell_margins]['top'])
					c.margin_bottom = pptx.util.Inches(margins_master[cell_margins]['bottom'])
					c.margin_left = pptx.util.Inches(margins_master[cell_margins]['left'])
					c.margin_right = pptx.util.Inches(margins_master[cell_margins]['right'])
					if banded_rows:
						if not keep_header_formatting:
							c.fill.solid()
							c.fill.fore_color.rgb = RGBColor(*style.RGB.grey_light) if (i-df.columns.nlevels) % 2 == 0 else RGBColor(*style.RGB.grey_light2)
					tf = c.text_frame
					p = tf.paragraphs[0]
					try:
						r = p.runs[0]
						if not keep_header_formatting:
							r.font.bold = index_bold
							r.font.italic = index_italic
							r.font.size = pptx.util.Pt(index_size)
							if not index_text_color is None:
								r.font.color.rgb = index_text_color
							r.font.name = text_font_name
						else:
							r.font.bold = header_bold
							r.font.italic = header_italic
							r.font.size = pptx.util.Pt(header_size)
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
				c.margin_top = pptx.util.Inches(margins_master[cell_margins]['top'])
				c.margin_bottom = pptx.util.Inches(margins_master[cell_margins]['bottom'])
				c.margin_left = pptx.util.Inches(margins_master[cell_margins]['left'])
				c.margin_right = pptx.util.Inches(margins_master[cell_margins]['right'])
				if banded_rows:
					c.fill.solid()
					c.fill.fore_color.rgb = RGBColor(*style.RGB.grey_light) if row % 2 == 0 else RGBColor(*style.RGB.grey_light2)
				tf = c.text_frame
				p = tf.paragraphs[0]
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
					r.font.size = pptx.util.Pt(text_size)
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
				tf = c.text_frame
				p = tf.paragraphs[0]
				r = p.runs[0]
				r.font.bold = totals_bold
				r.font.italic = totals_italic
				r.font.size = pptx.util.Pt(totals_size)
				if not totals_text_color is None:
					r.font.color.rgb = totals_text_color
				r.font.name = text_font_name
		if row_totals:
			for i in range(num_rows-df.columns.nlevels):
				if header:
					c = table.cell(i+df.columns.nlevels,num_cols-1)
				else:
					c = table.cell(i,num_cols-1)
				tf = c.text_frame
				p = tf.paragraphs[0]
				r = p.runs[0]
				r.font.bold = totals_bold
				r.font.italic = totals_italic
				r.font.size = pptx.util.Pt(totals_size)
				if not totals_text_color is None:
					r.font.color.rgb = totals_text_color
				r.font.name = text_font_name

		# customize table row hieghts
		if not row_height is None:
			emu = row_height * pptx.util.Length._EMUS_PER_INCH
			for r in table.rows:
				r.height = pptx.util.Emu(round(emu))

		# customize table column widths
		min_col_w = 4 #cm
		max_col_w = table_width / pptx.util.Length._EMUS_PER_CM / 2 # num_char_cols (don't hog the table)
		w_columns = 0
		for ix,col in enumerate(table.columns):
			if index and ix < df.index.nlevels:
				# compute width dynamically based on max text size in index, proportional to text size
				cm = min( max( len(max([s for s in df.index.get_level_values(ix)], key=len)) * 2 * (1/index_size), min_col_w), max_col_w)
				col.width = pptx.util.Cm(int(np.ceil(cm)))
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
			col.width = pptx.util.Cm(int(np.ceil(cm)))
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
				col.width = col.width - pptx.util.Emu(round(reduction_amount))
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
				col.width = col.width + pptx.util.Emu(round(extension_amount))
				w_columns += col.width

		# highlight rows, or columns
		table.first_row = highlight_first_row
		table.first_col = hightlight_first_col
		table.last_row = highlight_last_row

		return table_shape

	def create_chart(self, chart, df,
					 chart_type=pptx.enum.chart.XL_CHART_TYPE.LINE, #chart_style=1,
					 text_font_name=style.Font.name,
					 chart_title=None, chart_title_text_size=12, chart_title_text_bold=False, chart_title_text_color=None,
					 axis_label_text_size=10, axis_label_text_bold=False, axis_label_text_color=None,
					 axis_text_size=10, axis_text_bold=False, axis_text_color=None,
					 data_labels=True, data_label_position=None, data_label_rotate=False,
					 data_label_text_size=8, data_label_text_bold=False, data_label_text_color=None,
					 chart_legend=True, legend_in_layout=False, legend_position=pptx.enum.chart.XL_LEGEND_POSITION.BOTTOM,
					 legend_text_size=10, legend_text_bold=False, legend_text_color=None,
					 category_axis_label=None, value_axis_label=None,
					 highlight_line=None, number_format=None,
					 line_width=30000, bar_gap_width=None, bar_overlap=None,
					 encoding='utf-8', encoding_errors='strict'):
		"""Create a ppt chart using a pandas dataframe

		Parameters
		----------
		chart: pptx.shapes.chart
			pptx shape object of type chart, graphicframe placeholder
		df: pd.DataFrame
			pandas dataframe object

		Returns
		-------
		chart: pptx.shapes.chart
			pptx chart shape object, access chart.Chart for chart object inside placeholder shape

		Keyword Arguements
		------------------
		chart_type: pptx.enum.chart.XL_CHART_TYPE
			pptx enum chart type
		text_font_name: str
			string font name in Microsoft Word to overarchingly apply to all table text (e.g. 'Arial'), default pandas-msx.style.Font.name
		chart_title: str
			text to be written as chart title, default None
		chart_title_text_size: int
			title text size in ppt font size, default 12
		chart_title_text_bold: bool
			whether or not to bold title text, default False
		chart_title_text_color: list
			list of 3 RGB codes to color chart title text, default ...
		axis_label_text_size: int
			axis label text size in ppt font size, default 10
		axis_label_text_bold: bool
			whether or not to bold axis label text, default False
		axis_label_text_color: list
			list of 3 RGB codes to color axis label text, default ...
		axis_text_size: int
			axis text size in ppt font size, default 10
		axis_text_bold: bool
			whether or not to bold axis text, default False
		axis_text_color: list
			list of 3 RGB codes to color axis text text, default ...
		data_labels: bool
			whether or not to show data labels, default True
		data_label_position: pptx.enum.chart.XL_LEGEND_POSITION
			position of data labels relative to chart series, default None (i.e. ppt infers)
		data_label_rotate: bool
			whether or not to rotate data label text 180 degrees counterclockwise (i.e. vertical), default False
		data_label_text_size: int
			data label text size in ppt font size, default 8
		data_label_text_bold: bool
			whether or not to bold data label text, default False
		data_label_text_color: list
			list of 3 RGB codes to color data label text, default ...
		chart_legend: bool
			whether or not to include a chart legend, default True
		legend_in_layout: bool
			whether or not to include the chart legend inside the chart bounding box, default False
		legend_position: pptx.enum.chart.XL_LEGEND_POSITION.BOTTOM
			position of chart legend relative to chart bounding box, default BOTTOM
		legend_text_size: int
			legend text size in ppt font size, default 10
		legend_text_bold: bool
			whether or not to bold legend text, default False
		legend_text_color: list
			list of 3 RGB codes to color legend text, default ...
		category_axis_label: str
			text to be used as category axis label if chart has category axis, default None
		value_axis_label: str
			text to be used as category axis label if chart has category axis, default None
		highlight_line: str
			name of line chart series (dataframe column) to be increased in wheight by 2x for emphasis, default None
		number_format: str
			formatted string as per Microsoft's standard, default is None (i.e ppt infers), for help see: http://python-pptx.readthedocs.io/en/latest/api/enum/ExcelNumFormat.html
		line_width: int
			width of line in EMU, default 30000
		bar_gap_width: int
			percent of bar width (from 0 to 500) to be set as gap width between bars, default None (i.e. ppt infers)
		bar_overlap: int
			percent of bar width (from -100 to 100) to be set as overlap amount of adjacent bars, default None (i.e. ppt infers)

		Notes
		-----
		Your dataframe must be properly formatted! Use df.pivot_table() for typical transformation.
			- For chart type Line or Column, the dataframe index represents x axis, columns represents series, and values represents y axis.
			- For chart type Pie, the dataframe should have a single row with no index where columns represents series and values represent size.
		"""

		# impute any missing data as 0
		df = df.fillna(0)

		# create chart data
		chart_data = ChartData()

		# assign categories to chart data
		chart_data.categories = df.index

		# populate chart data
		if chart_type == pptx.enum.chart.XL_CHART_TYPE.PIE:
			# PIE charts are special, use only single column as series
			for col in df.columns:
				try:
					chart_data.add_series(str(col), (list(df[col])))
				except UnicodeEncodeError:
					chart_data.add_series(col.encode('ascii', errors='ignore'), (list(df[col])))
		else:
			# iterate rows, add each row as series
			for col,row in df.iteritems():
				try:
					chart_data.add_series(str(col), (list(row)))
				except UnicodeEncodeError:
					chart_data.add_series(col.encode('ascii', errors='ignore'), (list(row)))

		# insert chart into shape
		chart_shape = chart.insert_chart(chart_type, chart_data)

		# get chart object from graphic frame
		chart = chart_shape.chart

		# convert colors to pptx RGB
		if not chart_title_text_color is None:
			chart_title_text_color = RGBColor(*chart_title_text_color)
		if not axis_label_text_color is None:
			axis_label_text_color = RGBColor(*axis_label_text_color)
		if not axis_text_color is None:
			axis_text_color = RGBColor(*axis_text_color)
		if not data_label_text_color is None:
			data_label_text_color = RGBColor(*data_label_text_color)
		if not legend_text_color is None:
			legend_text_color = RGBColor(*legend_text_color)

		# convert font size to Pt
		chart_title_text_size = pptx.util.Pt(chart_title_text_size)
		axis_label_text_size = pptx.util.Pt(axis_label_text_size)
		axis_text_size = pptx.util.Pt(axis_text_size)
		data_label_text_size = pptx.util.Pt(data_label_text_size)
		legend_text_size = pptx.util.Pt(legend_text_size)

		# title
		if not chart_title is None:
			chart.has_title = True
			title = chart.chart_title
			title.has_text_frame = True
			title.text_frame.text = chart_title
			r = title.text_frame.paragraphs[0].runs[0]
			r.font.bold = chart_title_text_bold
			r.font.size = chart_title_text_size
			if not chart_title_text_color is None:
				r.font.color.rgb = chart_title_text_color
			r.font.name = text_font_name

		# legend
		chart.has_legend = chart_legend
		if chart_legend:
			chart.legend.include_in_layout = legend_in_layout
			chart.legend.position = legend_position
			chart.legend.font.size = legend_text_size
			chart.legend.font.bold = legend_text_bold
			if not legend_text_color is None:
				chart.legend.font.color.rgb = legend_text_color
			chart.legend.font.name = text_font_name

		# get plot
		plot = chart.plots[0]

		# labels
		plot.has_data_labels = data_labels
		if data_labels:
			plot.data_labels.font.size = data_label_text_size
			plot.data_labels.font.bold = data_label_text_bold
			if not data_label_text_color is None:
				plot.data_labels.font.color.rgb = data_label_text_color
			plot.data_labels.font.name = text_font_name
			if not data_label_position is None:
				plot.data_labels.position = data_label_position
			if data_label_rotate:
				# currently only supports rotation -270 degrees
				txPr = plot.data_labels._element.get_or_add_txPr()
				txPr.bodyPr.set('rot','-5400000')
			if number_format:
				plot.data_labels.number_format = number_format

		# set initial chart style (NOT WORKING)
		# chart.chart_style = chart_style

		# customize chart series
		color_ix = 0
		for i, series in enumerate(chart.series):
			try:
				if chart_type == pptx.enum.chart.XL_CHART_TYPE.PIE:
					for i, slice in enumerate(series.points):
						slice.format.fill.solid()
						slice.format.fill.fore_color.rgb = RGBColor(*style.RGB.colorbar_colorbrewer[i])
				if chart_type == pptx.enum.chart.XL_CHART_TYPE.LINE:
					series.format.line.width = line_width
					series.format.line.color.rgb = RGBColor(*style.RGB.colorbar_colorbrewer[i])
					if not highlight_line is None and i == df.columns.get_loc(highlight_line):
						series.format.line.width = line_width * 2
				else:
					series.format.fill.solid()
					series.format.fill.fore_color.rgb = RGBColor(*style.RGB.colorbar_colorbrewer[i])
					series.format.line.color.rgb = RGBColor(*style.RGB.white)
			except IndexError:
				# start recycling colors
				color_ix = 0
			color_ix += 1

		# axis
		try:
			# titles
			if not category_axis_label is None:
				if not chart.category_axis.has_title:
					chart.category_axis.has_title = True
				chart.category_axis.axis_title.text_frame.text = category_axis_label
				r = chart.category_axis.axis_title.text_frame.paragraphs[0].runs[0]
				r.font.bold = axis_label_text_bold
				r.font.size = axis_label_text_size
				if not axis_label_text_color is None:
					r.font.color.rgb = axis_label_text_color
				r.font.name = text_font_name
			if not value_axis_label is None:
				if not chart.value_axis.has_title:
					chart.value_axis.has_title = True
				chart.value_axis.axis_title.text_frame.text = value_axis_label
				r = chart.value_axis.axis_title.text_frame.paragraphs[0].runs[0]
				r.font.bold = axis_label_text_bold
				r.font.size = axis_label_text_size
				if not axis_label_text_color is None:
					r.font.color.rgb = axis_label_text_color
				r.font.name = text_font_name
			# turn off grid lines
			chart.category_axis.has_major_gridlines = False
			chart.category_axis.has_minor_gridlines = False
			chart.value_axis.has_major_gridlines = False
			chart.value_axis.has_minor_gridlines = False

			# format axis text
			chart.value_axis.tick_labels.font.bold = axis_text_bold
			chart.value_axis.tick_labels.font.size = axis_text_size
			if not axis_text_color is None:
				chart.value_axis.tick_labels.font.color.rgb = axis_text_color
			chart.value_axis.tick_labels.font.name = text_font_name
			chart.category_axis.tick_labels.font.bold = axis_text_bold
			chart.category_axis.tick_labels.font.size = axis_text_size
			if not axis_text_color is None:
				chart.category_axis.tick_labels.font.color.rgb = axis_text_color
			chart.category_axis.tick_labels.font.name = text_font_name
			if number_format:
				chart.value_axis.tick_labels.number_format = number_format
		except ValueError:
			# pie charts do not have axis
			pass

		# customize bars
		if not bar_gap_width is None or not bar_overlap is None:
			try:
				if not bar_gap_width is None:
					plot.gap_width = bar_gap_width
				if not bar_overlap is None:
					plot.overlap = bar_overlap
			except ValueError:
				# only bar and column charts have this setting
				pass

		return chart_shape
