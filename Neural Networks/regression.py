import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

N = 500
X = np.random.random((N, 2)) * 4 - 2
Y = X[:,0] * X[:,1]

def plot_data():
    fig = plt.figure()
    ax = fig.add_subplot(111, projection="3d")
    ax.scatter(X[:,0], X[:,1], Y)
    plt.show()

D = 2 # number of dimensions
M = 100 # number of hidden units

# layer 1
W = np.random.randn(D, M) / np.sqrt(D)
B = np.zeros(M)

# layer 2
V = np.random.randn(M) /  np.sqrt(M)
c = 0

def relu(Z):
    return Z * (Z > 0)

def forward(X):
    Z = X.dot(X) + b
    Z = relu(Z)
    Y_hat = Z.dot(V) + c
    return Z, Y_hat