"""Convenience tools for accessing PPT templates using python-pptx.
"""

import pptx

# Microsoft cell margins in inches
margins_master = {
    'normal': {'top': 0.05, 'bottom': 0.05, 'left': 0.1, 'right': 0.1},
	'none': {'top': 0, 'bottom': 0, 'left': 0, 'right': 0},
	'narrow': {'top': 0.05, 'bottom': 0.05, 'left': 0.05, 'right': 0.05},
	'wide': {'top': 0.15, 'bottom': 0.15, 'left': 0.15, 'right': 0.15},
	# custom style, does not exist in ppt
	'tight': {'top': 0.025, 'bottom': 0.025, 'left': 0.025, 'right': 0.025},
}

def map_layouts(ppt, verbose=False):
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

def map_shapes(layout, verbose=False):
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

def set_row_height(table, row_height=0.15):
    """Set row height for table rows.

    Parameters
    ----------
    row_height: float
        Row height in inches.

    Returns:
    --------
    table: pptx.table.Table
        pptx table graphic frame object, with new row height.
    """
    emu = row_height * pptx.util.Length._EMUS_PER_INCH
    for r in table.rows:
        r.height = pptx.util.Emu(round(emu))
    return table

def format_cell(cell,
                font_size=None,
                font_color=None,
                font_name=None,
                bold=False,
                fill=False,
                fill_color=None,
                cell_margins='tight'):
    """Format a given table cell.

    Parameters
    ----------
    font_size: int
        Cell text font size. For more, see pptx.util.Pt
    font_color: tuple
        Cell text font color. Must be RGB code as tuple of 3 integers, or HEX code as string. For more see pptx.dml.color.RGBColor
    font_name:
        Cell text font name, for example 'Arial'.
    bold: bool
        Whether or not to bold the text.
    fill: bool
        Whether or not to fill the cell backgound color.
    fill_color: tuple or pptx.enum.dml.MSO_THEME_COLOR
        Color to fill cell background. Must be RGB code as tuple of 3 integers, or instance of pptx.enum.dml.MSO_THEME_COLOR.
    cell_margins: str
        Keyword for setting cell margin widths. Use one of 'normal', 'none', 'narrow', 'tight', or 'wide'. Keywords are adopted from ppt with a custom tight setting.

    Returns:
    --------
    cell: pptx.table._Cell
        Table cell with applied formatting.
    """
    cell.margin_top = pptx.util.Inches(margins_master[cell_margins]['top'])
    cell.margin_bottom = pptx.util.Inches(margins_master[cell_margins]['bottom'])
    cell.margin_left = pptx.util.Inches(margins_master[cell_margins]['left'])
    cell.margin_right = pptx.util.Inches(margins_master[cell_margins]['right'])

    if fill:
        cell.fill.solid()
        if isinstance(fill_color,pptx.enum.base.EnumValue):
            cell.fill.fore_color.theme_color = fill_color
        elif isinstance(fill_color,tuple):
            cell.fill.fore_color.rgb = pptx.dml.color.RGBColor(*fill_color)
        elif isinstance(fill_color,str):
            cell.fill.fore_color.rgb = pptx.dml.color.RGBColor.from_string(fill_color) if not fill_color.startswith('#') else pptx.dml.color.RGBColor.from_string(fill_color[1:])
        else:
            raise ValueError('Incorrect value for fill_color. \
            Please provide one of RGB code as `tuple` of 3 integers, HEX code as string, or an instance of `pptx.enum.dml.MSO_THEME_COLOR`')

    tf = cell.text_frame
    try:
        p = tf.paragraphs[0]
    except IndexError:
        p = tf.add_paragraph()
    try:
        r = p.runs[0]
    except IndexError:
        r = p.add_run()
    if not font_size == None:
        r.font.size = pptx.util.Pt(font_size)
    if not font_color == None:
        if isinstance(font_color,tuple):
            r.font.color.rgb = pptx.dml.color.RGBColor(*font_color)
        elif isinstance(font_color,str):
            r.font.color.rgb = pptx.dml.color.RGBColor.from_string(font_color[1:]) if font_color.startswith('#') else pptx.dml.color.RGBColor.from_string(font_color)
        else:
            raise ValueError('Incorrect value for font_color. \
            Please provide one of RGB code as `tuple` of 3 integers, or a HEX code as string')
    if not font_name == None:
        r.font.name = font_name
    if bold:
        r.font.bold = True
    return cell
