import docx

def add_hyperlink(paragraph, url, text,
                  color=None,
                  underline=False,
                  font_size=None):
    """
    A function that places a hyperlink within a docucment paragraph object.

    This is an updated copy from what was originally suggested here: https://github.com/python-openxml/python-docx/issues/74#issuecomment-261169410

    Parameters:
    -----------
    paragraph: docx.text.paragraph.Paragraph[source]
        the paragraph where the hyperlink will be appended to
    url: str
        a string containing the required url
    text: str
        text to be displayed

    Returns:
    --------
    hyperlink: docx.oxml.shared.OxmlElement('w:hyperlink')
        the hyperlink object
    """

    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    # Add color if it is given
    if not color is None:
      c = docx.oxml.shared.OxmlElement('w:color')
      c.set(docx.oxml.shared.qn('w:val'), color)
      rPr.append(c)

    # Add underlining if it is requested
    if underline:
      u = docx.oxml.shared.OxmlElement('w:u')
      u.set(docx.oxml.shared.qn('w:val'), 'none')
      rPr.append(u)

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text

    # Customize font size if it is requested
    if not font_size is None:
      size = docx.oxml.shared.OxmlElement('w:sz')
      size.set(docx.oxml.shared.qn('w:val'), str(font_size*2))
      rPr.append(size)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink
