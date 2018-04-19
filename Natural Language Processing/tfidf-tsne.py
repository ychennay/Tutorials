import json
import numpy as np
import matplotlib.pyplot as plt
from sklearn.utils import shuffle
from sklearn.manifold import TSNE
from datetime import datetime
from utils import get_wikipedia_data

import os
import sys
sys.path.append(os.path.abspath('..'))

sentences, word2idx = get_wikipedia_data(n_files=10, n_vocab=1500, by_paragraph=True)
print(sentences)
