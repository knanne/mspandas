"""Default style guide

You can, for example, change this according to your company's branding standards.
"""


class Font():
	name = 'Calibri'


class RGB():
	white = [255, 255, 255]
	grey_light = [245, 245, 245]
	grey_light2 = [235, 235, 235]
	grey = [150, 150, 150]
	grey_dark = [80, 80, 80]
	colorbar_colorbrewer = [
		[141,211,199], # colorbrewer 1
		[255,255,179], # colorbrewer 2
		[190,186,218], # colorbrewer 3
		[251,128,114], # colorbrewer 4
		[128,177,211], # colorbrewer 5
		[253,180,98], # colorbrewer 6
		[179,222,105], # colorbrewer 7
		[252,205,229], # colorbrewer 8
		[217,217,217], # colorbrewer 9
		[188,128,189], # colorbrewer 10
		[204,235,197], # colorbrewer 11
		[255,237,111], # colorbrewer 12
	]
	colorbar_microsoft = [
	    [255, 255, 0], # microsoft yellow
		[146, 208, 80], # microsoft light green
		[255, 192, 0], # microsoft orange
		[112, 48, 160], # microsoft purple
		[0, 176, 80], # microsoft green
		[0, 32, 96], # microsoft dark blue
		[0, 112, 192], # microsoft blue
		[192, 0, 0], # microsoft dark red
		[0, 176, 240], # microsoft light blue
		[255, 0, 0], # microsoft red
	]


class Hex():
	white = '#FFFFFF'
	grey_light = '#F5F5F5'
	grey_light2 = '#EBEBEB'
	grey = '#969696'
	grey_dark = '#505050'
	colorbar_colorbrewer = [
		'#8dd3c7', # colorbrewer 1
		'#ffffb3', # colorbrewer 2
		'#bebada', # colorbrewer 3
		'#fb8072', # colorbrewer 4
		'#80b1d3', # colorbrewer 5
		'#fdb462', # colorbrewer 6
		'#b3de69', # colorbrewer 7
		'#fccde5', # colorbrewer 8
		'#d9d9d9', # colorbrewer 9
		'#bc80bd', # colorbrewer 10
		'#ccebc5', # colorbrewer 11
		'#ffed6f', # colorbrewer 12
	]
	colorbar_microsoft = [
	    '#FFFF00', # microsoft yellow
		'#92D050', # microsoft light green
		'#FFC000', # microsoft orange
		'#7030A0', # microsoft purple
		'#00B050', # microsoft green
		'#002060', # microsoft dark blue
		'#0070C0', # microsoft blue
		'#C00000', # microsoft dark red
		'#00B0F0', # microsoft light blue
		'#FF0000', # microsoft red
	]
