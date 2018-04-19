import numpy as np
import theano
import theano.tensor as T
import matplotlib.pyplot as plt

from utils import init_weight, all_parity_pairs
from sklearn.utils import shuffle

class HiddenLayer:

    def __init__(self, M1, M2, an_id):
        self.id = an_id # id to help debug
        self.M1 = M1 # length of incoming layer
        self.M2 = M2 # length of next layer
        W = init_weight(M1, M2) # initialize weights
        B = np.zeros(M2) # of shape next layer

        self.W = theano.shared(W, f"W_{self.id}")
        self.b = theano.shared(b, f"b_{self.id}")
        self.params = [self.W, self.b]

    def forward(self, X):
        return T.nnet.relu(X.dot(self.W) + self.b)

class ANN:

    def __init__(self, hidden_layer_sizes):
        self.hidden_layer_sizes = hidden_layer_sizes

    def fit(self, X, Y, learning_rate=10e-3, mu=0.99, reg=10e-12, eps=10e-10, epochs-400, batch_sz=20, print_period=1, show_fig=False):
        Y = Y.astype(np.int32)

        N, D = X.shape # N = # of rows, D = # of dimensions
        K = len(set(Y)) # number of classes
        self.hidden_layers = []
        M1 = D # the number of rows in the first weight matrix will be the number of dimensions in the feature space
        count = 0

        for M2 in self.hidden_layer_sizes: # create new hidden layers
            h = HiddenLayer(M1, M2, count)
            self.hidden_layers.append(h)
            M1 = M2
            count += 1
        W = init_weight(M1, K)
        b = np.zeros(K)

        self.W = theano.shared(W, 'W_logreg')
        self.b = theano.shared(b, 'b_logreg')

        self.params = [self.W, self.b]
        for h in self.hidden_layers: # iterate through adding the different parameters in the hidden layers ([weight matrix, bias vector])
            self.params += h.params

        dparams = [theano.shared(np.zeros(p.get_value().shape)) for p in self.params]

        thX = T.matrix('X') #thX stands for theanoX variable
        thY = T.ivector('Y') #thY stands for theanoY variable
        pY = self.forward(thX)

        rcost = reg * T.sum([(p*p).sum() for p in self.params])
        cost = -T.mean(T.log(pY[T.arange(thY.shape[0]), thY])) + rcost
        prediction = self.predict(thX)
