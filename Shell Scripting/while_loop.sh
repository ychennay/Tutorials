#! /bin/bash

counter = 0
while [ $counter -lt 50 ]; do
	let counter=counter + 1
	echo $counter
