import string
import numpy as np
import matplotlib.pyplot
import theano
import theano.tensor as T
from sklearn.utils import shuffle

def remove_punctuation(text):
    return text.translate(str.maketrans('','',string.punctuation))

def fetch_frost_corpus():
    # START and END tokens for RNN will be the first two indices in our dictionary
    word_to_index = {'START': 0,
                     'END': 1}

    current_index = 2
    sentences = []
    for line in open('./robert_frost.txt'):
        line = line.strip() # strip whitespace characters from beginning and end
        if line:
            tokens = remove_punctuation(line.lower()).split() # lowercase the line, remove punctuation, and split on whitespace
            sentence = []
            for token in tokens: # iterate through each token in the sentence
                if token not in word_to_index: # check if the token is in the current dictionary
                    word_to_index[token] = current_index # assign it a spot in our dictionary
                    current_index += 1 # increment the current index

                index = word_to_index[token] # get the index of the current token
                sentence.append(index) # add it to our sentence list (sentence = [23, 212, 11, 42, 22, 10])
            sentences.append(sentence) # add it to our list of sentences
    return sentences, word_to_index

class SimpleRNN:

    def __init__(self, D, M, V):
        '''
        D: word embedding size
        M: hidden layer size
        V: vocabulary size
        '''
        self.D = D
        self.M = M
        self.V = V

    def fit(self, X, learning_rate=10e-1, mu=0.99, reg=1.0,
            activation=T.tanh, epochs=500, show_fig=False):

        N = len(X) # N = number of data points
        D = self.D # word embedding length
        M = self.M # hidden layer
        V = self.V # vocabulary
        self.f = activation

        We = self.init_weight(V, D)
        Wx = self.init_weight(D, M)
        Wh = self.init_weight(M, M)
        bh = np.zeros(M) # bias for the hidden to hidden recurrent layer
        h0 = np.zeros(M)
        Wo = init_weight(M,V) # weights for the output
        b0 = np.zeros(V) # bias for the output (a softmax of vocabulary)

        self.We = theano.shared(We) # word embeddings
        self.Wx = theano.shared(Wx)
        self.Wh = theano.shared(Wh)
        self.bh = theano.shared(bh)
        self.h0 = theano.shared(h0)
        self.Wo = theano.shared(Wo)
        self.bo = theano.shared(bo)

        self.params = [self.We,
                       self.Wx,
                       self.Wh,
                       self.bh,
                       self.h0,
                       self.Wo,
                       self.bo]

        thX = T.ivector('X')
        Ei = self.We(thX) # T x D matrix, T is the length of the sequence, word embeddings of the ith line
        thY = T.ivector('Y')

        def recurrence(x_t, h_t1):
            #returns h(t), y(t)
            h_t = self.f(x_t.dot(self.Wx) + h_t1.dot(self.Wh) + self.bh)
            y_t = T.nnet.softmax(h_t.dot(self.Wo) + self.bo)
            return h_t, y_t

        [h, y], _ = theano.scan(fn=recurrence,
                                outputs_info=[self.h0, None], # initial values for
                                sequences=Ei,
                                n_steps=Ei.shape[0])

        py_x = y[:,0,:]




    def init_weight(self, Mi, Mo):
        return np.random.randn(Mi, Mo) / np.sqrt(Mi + Mo)



if __name__ == "__main__":
    fetch_frost_corpus()
