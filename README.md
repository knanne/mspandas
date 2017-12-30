# mspandas

`mspandas` is a convenience wrapper on top of [`python-pptx`](http://python-pptx.readthedocs.io/en/latest/) and [`python-docx`](https://python-docx.readthedocs.io/en/latest/) which accepts [`pandas`](https://pandas.pydata.org/) `DataFrames` for writing tables and charts in Microsoft PowerPoint and Word Documents. The project homepage is [github.com/knanne/mspandas](https://github.com/knanne/mspandas)  

# Background

This library was designed to help automate analytical reporting pipelines, by making a lot of the code reusable. This library started out as multiple reusable functions which kept growing and growing until it was decided to compile them into a library for import. Hopefully, by open sourcing this code, it can be seriously improved (along with my skills in Python development).

The main features of this library are the ability to quickly write Pandas DataFrames to a Microsoft Office tables and charts. Currently, the library includes the following notable methods.  

  - `mspandas.pandasPPT.create_table`
  - `mspandas.pandasDOC.create_table`
  - `mspandas.pandasPPT.create_chart`

In addition to the above functions, included are also some helpful methods like `pandasPPT.map_layouts` and `pandasPPT.map_shapes`, and possibly some helpful code snippets, for not-yet implemented features in the base libraries for, which can be found under `mspandas.monkey_patches`.  

# Usage

Please refer to the example Jupyter Notebooks in [/examples](/examples)  

Basic code for reusable PowerPoint reporting looks like:  

```python
import pptx
import pandas as pd
from mspandas import pandasPPT

Handler = pandasPPT.Handler()

df = pd.DataFrame(np.random.rand(4, 4), columns=['a', 'b', 'c', 'd'])

ppt = pptx.Presentation('template.pptx')

layout_map = Handler.map_layouts(ppt=ppt)
slide = ppt.slides.add_slide(layout_map['Slide Layout with Chart'])

shape_map = Handler.map_shapes(layout_map['Slide Layout with Chart'])
chart = slide.placeholders[shape_map['Chart Placeholder']]

chart = Handler.create_chart(chart, df)

ppt.save('report.pptx')
```

Basic code for reusable Word reporting looks like:  

```python
import docx
import pandas as pd
from mspandas import pandasDOC

Handler = pandasDOC.Handler()

df = pd.DataFrame(np.random.rand(4, 4), columns=['a', 'b', 'c', 'd'])

doc = docx.Document()

table = Handler.create_table(doc, df)

doc.save('report.docx')
```

# Installation

**This library is currently not on pip!**  

### Option 1

Place this library in your Python's site packages folder (e.g. `~\Continuum\Anaconda3\Lib\site-packages`)  

### Option 2

Alternatively, to temporarily add this library to path during runtime, run the following.  

```python
import sys
import os

# define path
mspandas = '~/path-to/mspandas'

# add mspandas to path
sys.path.append(os.path.abspath(mspandas))

import mspandas
```

# Dependencies

The library currently works on both Python 2.7 and 3+, and is also tested on Pandas 0.19+  

The modules in this library are designed to be used in addition to the following libraries. Therefore, please first educate yourself on documentation of those libraries.  

[Python PPTx](https://python-pptx.readthedocs.io/en/latest/). Install via `pip install python-pptx`  
[Python DOCx](https://python-pptx.readthedocs.io/en/latest/). Install via `pip install python-pptx`  

[Pandas](http://pandas.pydata.org/). Install via `pip install pandas`  

[Numpy](http://www.numpy.org/). Install via `pip install numpy`  
