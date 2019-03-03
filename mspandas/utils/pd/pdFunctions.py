"""Convenience tools for accessing PPT templates using python-pptx.
"""

def get_column_types(df):
    """Get map of DataFrame columns by type

    Parameters
    ---------
    df = pandas.DataFrame
        DataFrame on which to map column types

    Returns
    -------
    types: dict
        Retreived map of column names by data type
    """
    types = {}
    for col,vals in df.iteritems():
        types[vals.dtype] = types.get(vals.dtype,[]) + [col]
    return types
