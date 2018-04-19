# https://deeplearningcourses.com/c/deep-learning-recurrent-neural-networks-in-python
# https://udemy.com/deep-learning-recurrent-neural-networks-in-python
from __future__ import print_function, division
from builtins import range
# Note: you may need to update your version of future
# sudo pip install -U future


import numpy as np
import string
import os
import sys
import operator
from nltk import pos_tag, word_tokenize
from datetime import datetime

def init_weight(Mi, Mo):
    return np.random.randn(Mi, Mo) / np.sqrt(Mi + Mo)

def all_parity_pairs(nbit):
    # total number of samples (Ntotal) will be a multiple of 100
    # why did I make it this way? I don't remember.
    N = 2**nbit
    remainder = 100 - (N % 100)
    Ntotal = N + remainder
    X = np.zeros((Ntotal, nbit))
    Y = np.zeros(Ntotal)
    for ii in range(Ntotal):
        i = ii % N
        # now generate the ith sample
        for j in range(nbit):
            if i % (2**(j+1)) != 0:
                i -= 2**j
                X[ii,j] = 1
        Y[ii] = X[ii].sum() % 2
    return X, Y

def all_parity_pairs_with_sequence_labels(nbit):
    X, Y = all_parity_pairs(nbit)
    N, t = X.shape

    # we want every time step to have a label
    Y_t = np.zeros(X.shape, dtype=np.int32)
    for n in range(N):
        ones_count = 0
        for i in range(t):
            if X[n,i] == 1:
                ones_count += 1
            if ones_count % 2 == 1:
                Y_t[n,i] = 1

    X = X.reshape(N, t, 1).astype(np.float32)
    return X, Y_t

# unfortunately Python 2 and 3 translates work differently
def remove_punctuation_2(s):
    return s.translate(None, string.punctuation)

def remove_punctuation_3(s):
    return s.translate(str.maketrans('','',string.punctuation))

if sys.version.startswith('2'):
    remove_punctuation = remove_punctuation_2
else:
    remove_punctuation = remove_punctuation_3


def get_robert_frost():
    word2idx = {'START': 0, 'END': 1}
    current_idx = 2
    sentences = []
    for line in open('../hmm_class/robert_frost.txt'):
        line = line.strip()
        if line:
            tokens = remove_punctuation(line.lower()).split()
            sentence = []
            for t in tokens:
                if t not in word2idx:
                    word2idx[t] = current_idx
                    current_idx += 1
                idx = word2idx[t]
                sentence.append(idx)
            sentences.append(sentence)
    return sentences, word2idx

def my_tokenizer(s):
    s = remove_punctuation(s)
    s = s.lower() # downcase
    return s.split()

def get_wikipedia_data(n_files, n_vocab, by_paragraph=False):
    prefix = '../large_files/'

    if not os.path.exists(prefix):
        print("Are you sure you've downloaded, converted, and placed the Wikipedia data into the proper folder?")
        print("I'm looking for a folder called large_files, adjacent to the class folder, but it does not exist.")
        print("Please download the data from https://dumps.wikimedia.org/")
        print("Quitting...")
        exit()

    input_files = [f for f in os.listdir(prefix) if f.startswith('enwiki') and f.endswith('txt')]

    if len(input_files) == 0:
        print("Looks like you don't have any data files, or they're in the wrong location.")
        print("Please download the data from https://dumps.wikimedia.org/")
        print("Quitting...")
        exit()

    # return variables
    sentences = []
    word2idx = {'START': 0, 'END': 1}
    idx2word = ['START', 'END']
    current_idx = 2
    word_idx_count = {0: float('inf'), 1: float('inf')}

    if n_files is not None:
        input_files = input_files[:n_files]

    for f in input_files:
        print("reading:", f)
        for line in open(prefix + f):
            line = line.strip()
            # don't count headers, structured data, lists, etc...
            if line and line[0] not in ('[', '*', '-', '|', '=', '{', '}'):
                if by_paragraph:
                    sentence_lines = [line]
                else:
                    sentence_lines = line.split('. ')
                for sentence in sentence_lines:
                    tokens = my_tokenizer(sentence)
                    for t in tokens:
                        if t not in word2idx:
                            word2idx[t] = current_idx
                            idx2word.append(t)
                            current_idx += 1
                        idx = word2idx[t]
                        word_idx_count[idx] = word_idx_count.get(idx, 0) + 1
                    sentence_by_idx = [word2idx[t] for t in tokens]
                    sentences.append(sentence_by_idx)

    # restrict vocab size
    sorted_word_idx_count = sorted(word_idx_count.items(), key=operator.itemgetter(1), reverse=True)
    word2idx_small = {}
    new_idx = 0
    idx_new_idx_map = {}
    for idx, count in sorted_word_idx_count[:n_vocab]:
        word = idx2word[idx]
        print(word, count)
        word2idx_small[word] = new_idx
        idx_new_idx_map[idx] = new_idx
        new_idx += 1
    # let 'unknown' be the last token
    word2idx_small['UNKNOWN'] = new_idx
    unknown = new_idx

    assert('START' in word2idx_small)
    assert('END' in word2idx_small)
    assert('king' in word2idx_small)
    assert('queen' in word2idx_small)
    assert('man' in word2idx_small)
    assert('woman' in word2idx_small)

    # map old idx to new idx
    sentences_small = []
    for sentence in sentences:
        if len(sentence) > 1:
            new_sentence = [idx_new_idx_map[idx] if idx in idx_new_idx_map else unknown for idx in sentence]
            sentences_small.append(new_sentence)

    return sentences_small, word2idx_small

def get_tags(s):
    tuples = pos_tag(word_tokenize(s))
    return [y for x, y in tuples]

def get_poetry_classifier_data(samples_per_class, load_cached=True, save_cached=True):
    datafile = 'poetry_classifier_data.npz'
    if load_cached and os.path.exists(datafile):
        npz = np.load(datafile)
        X = npz['arr_0']
        Y = npz['arr_1']
        V = int(npz['arr_2'])
        return X, Y, V

    word2idx = {}
    current_idx = 0
    X = []
    Y = []
    for fn, label in zip(('../hmm_class/edgar_allan_poe.txt', '../hmm_class/robert_frost.txt'), (0, 1)):
        count = 0
        for line in open(fn):
            line = line.rstrip()
            if line:
                print(line)
                # tokens = remove_punctuation(line.lower()).split()
                tokens = get_tags(line)
                if len(tokens) > 1:
                    # scan doesn't work nice here, technically could fix...
                    for token in tokens:
                        if token not in word2idx:
                            word2idx[token] = current_idx
                            current_idx += 1
                    sequence = np.array([word2idx[w] for w in tokens])
                    X.append(sequence)
                    Y.append(label)
                    count += 1
                    print(count)
                    # quit early because the tokenizer is very slow
                    if count >= samples_per_class:
                        break
    if save_cached:
        np.savez(datafile, X, Y, current_idx)
    return X, Y, current_idx


def get_stock_data():
    input_files = os.listdir('stock_data')
    min_length = 2000

    # first find the latest start date
    # so that each time series can start at the same time
    max_min_date = datetime(2000, 1, 1)
    line_counts = {}
    for f in input_files:
        n = 0
        for line in open('stock_data/%s' % f):
            # pass
            n += 1
        line_counts[f] = n
        if n > min_length:
            # else we'll ignore this symbol, too little data
            # print 'stock_data/%s' % f, 'num lines:', n
            last_line = line
            date = line.split(',')[0]
            date = datetime.strptime(date, '%Y-%m-%d')
            if date > max_min_date:
                max_min_date = date

    print("max min date:", max_min_date)

    # now collect the data up to min date
    all_binary_targets = []
    all_prices = []
    for f in input_files:
        if line_counts[f] > min_length:
            prices = []
            binary_targets = []
            first = True
            last_price = 0
            for line in open('stock_data/%s' % f):
                if first:
                    first = False
                    continue
                date, price = line.split(',')[:2]
                date = datetime.strptime(date, '%Y-%m-%d')
                if date < max_min_date:
                    break
                prices.append(float(price))
                target = 1 if last_price < price else 0
                binary_targets.append(target)
                last_price = price
            all_prices.append(prices)
            all_binary_targets.append(binary_targets)

    # D = number of symbols
    # T = length of series
    return np.array(all_prices).T, np.array(all_binary_targets).T # make it T x D



import os
import numpy as np

def init_weight(Mi, Mo):
    return np.random.randn(Mi, Mo) / np.sqrt(Mi + Mo)


def find_analogies(w1, w2, w3, We, word2idx):
    king = We[word2idx[w1]]
    man = We[word2idx[w2]]
    woman = We[word2idx[w3]]
    v0 = king - man + woman

    def dist1(a, b):
        return np.linalg.norm(a - b)
    def dist2(a, b):
        return 1 - a.dot(b) / (np.linalg.norm(a) * np.linalg.norm(b))

    for dist, name in [(dist1, 'Euclidean'), (dist2, 'cosine')]:
        min_dist = float('inf')
        best_word = ''
        for word, idx in iteritems(word2idx):
            if word not in (w1, w2, w3):
                v1 = We[idx]
                d = dist(v0, v1)
                if d < min_dist:
                    min_dist = d
                    best_word = word
        print("closest match by", name, "distance:", best_word)
        print(w1, "-", w2, "=", best_word, "-", w3)


class Tree:
    def __init__(self, word, label):
        self.left = None
        self.right = None
        self.word = word
        self.label = label


def display_tree(t, lvl=0):
    prefix = ''.join(['>']*lvl)
    if t.word is not None:
        print("%s%s %s" % (prefix, t.label, t.word))
    else:
        print("%s%s -" % (prefix, t.label))
        # if t.left is None or t.right is None:
        #     raise Exception("Tree node has no word but left and right child are None")
    if t.left:
        display_tree(t.left, lvl + 1)
    if t.right:
        display_tree(t.right, lvl + 1)


current_idx = 0
def str2tree(s, word2idx):
    # take a string that starts with ( and MAYBE ends with )
    # return the tree that it represents
    # EXAMPLE: "(3 (2 It) (4 (4 (2 's) (4 (3 (2 a) (4 (3 lovely) (2 film))) (3 (2 with) (4 (3 (3 lovely) (2 performances)) (2 (2 by) (2 (2 (2 Buy) (2 and)) (2 Accorsi))))))) (2 .)))"
    # NOTE: not every node has 2 children (possibly not correct ??)
    # NOTE: not every node has a word
    # NOTE: every node has a label
    # NOTE: labels are 0,1,2,3,4
    # NOTE: only leaf nodes have words
    # s[0] = (, s[1] = label, s[2] = space, s[3] = character or (

    # print "Input string:", s, "len:", len(s)

    global current_idx

    label = int(s[1])
    if s[3] == '(':
        t = Tree(None, label)
        # try:

        # find the string that represents left child
        # it can include trailing characters we don't need, because we'll only look up to )
        child_s = s[3:]
        t.left = str2tree(child_s, word2idx)

        # find the string that represents right child
        # can contain multiple ((( )))
        # left child is completely represented when we've closed as many as we've opened
        # we stop at 1 because the first opening paren represents the current node, not children nodes
        i = 0
        depth = 0
        for c in s:
            i += 1
            if c == '(':
                depth += 1
            elif c == ')':
                depth -= 1
                if depth == 1:
                    break
        # print "index of right child", i

        t.right = str2tree(s[i+1:], word2idx)

        # except Exception as e:
        #     print "Exception:", e
        #     print "Input string:", s
        #     raise e

        # if t.left is None or t.right is None:
        #     raise Exception("Tree node has no word but left and right child are None")
        return t
    else:
        # this has a word, so it's a leaf
        r = s.split(')', 1)[0]
        word = r[3:].lower()
        # print "word found:", word

        if word not in word2idx:
            word2idx[word] = current_idx
            current_idx += 1

        t = Tree(word2idx[word], label)
        return t


def get_ptb_data():
    # like the wikipedia dataset, I want to return 2 things:
    # word2idx mapping, sentences
    # here the sentences should be Tree objects

    if not os.path.exists('../large_files/trees'):
        print("Please create ../large_files/trees relative to this file.")
        print("train.txt and test.txt should be stored in there.")
        print("Please download the data from http://nlp.stanford.edu/sentiment/")
        exit()
    elif not os.path.exists('../large_files/trees/train.txt'):
        print("train.txt is not in ../large_files/trees/train.txt")
        print("Please download the data from http://nlp.stanford.edu/sentiment/")
        exit()
    elif not os.path.exists('../large_files/trees/test.txt'):
        print("test.txt is not in ../large_files/trees/test.txt")
        print("Please download the data from http://nlp.stanford.edu/sentiment/")
        exit()

    word2idx = {}
    train = []
    test = []

    # train set first
    for line in open('../large_files/trees/train.txt'):
        line = line.rstrip()
        if line:
            t = str2tree(line, word2idx)
            # if t.word is None and t.left is None and t.right is None:
            #     print "sentence:", line
            # display_tree(t)
            # print ""
            train.append(t)
            # break

    # test set
    for line in open('../large_files/trees/test.txt'):
        line = line.rstrip()
        if line:
            t = str2tree(line, word2idx)
            test.append(t)
    return train, test, word2idx
