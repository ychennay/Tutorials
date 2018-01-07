import numpy as np
from datetime import datetime

a = np.random.randn(100)
b = np.random.randn(100)

T = 100000

def slow_dot_product(a, b):
	
	result = 0
	for e, f in zip(a, b):
		result += e * f
	return result

start = datetime.now()

for t in range(T):
	result = slow_dot_product(a, b)

elapsed_time = datetime.now() - start

print("Total time for slow method: {}, final result: {}".format(elapsed_time, result))

start = datetime.now()

for t in range(T):
	result = np.dot(a, b)

elapsed_time = datetime.now() - start

print("Total time for fast method: {}, final result: {}".format(elapsed_time, result))




