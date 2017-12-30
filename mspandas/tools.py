import pandas as pd
import numpy as np

class Dummy():
	"""Dummy data used for populating sample reports.
	"""

	title = 'Lorem Ipsum '
	sentence_short = 'Ut pharetra enim id fermentum sodales. '
	sentence_long = 'Neque porro quisquam est qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit... '
	paragraph = 'Maecenas enim nulla, commodo vitae aliquam nec, semper eu lacus. Cras ut ligula porta, tempor ante nec, cursus ipsum. Cras venenatis enim a lectus dictum, a faucibus libero ultricies. Pellentesque vitae elit eu velit tincidunt ultricies ut ut eros. Praesent sit amet tristique arcu. Ut pharetra enim id fermentum sodales. Suspendisse ut ante tempus, tempus enim efficitur, maximus enim. '

	df = pd.DataFrame(np.random.rand(6, 4),
					  columns=['a', 'b', 'c', 'd'],
					  index=list(range(pd.datetime.today().year-5,pd.datetime.today().year+1)))

	date = pd.datetime.now().strftime('%Y-%m-%d')
