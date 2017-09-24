var yu = {
	name: 'Yu',
	lastName: 'Chen',
	yearOfBirth: 1988,
	job: 'consultant',
	isMarried: false,
	testScores: [98,99,78],
	calculateAge: function(){
		return (new Date()).getFullYear() - this.yearOfBirth;

	}

};

var ali = new Object()
ali.name = 'Ali'
ali.lastName = 'Wallace'
ali['yearOfBirth'] = 1989

console.log(yu.isMarried)
console.log(yu.testScores)
console.log(yu.calculateAge(1988))
console.log(ali)