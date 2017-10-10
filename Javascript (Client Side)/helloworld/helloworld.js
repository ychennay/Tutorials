var name = "Yu Chen"
var yearOfBirth;
var yearOfRetirement = 65;
var currentYear = (new Date()).getFullYear()
console.log("The current year is " + currentYear)


// console.log(calculateYearsUntilRetirement(32)) -> you can't use this until after declaration
console.log(calculateAge(1992)) // -> but you can use this due to function hoisting

// a function expression
function calculateAge(yearOfBirth) {
	var currentAge = currentYear - yearOfBirth
	return currentAge;
} 

// a function expression
var calculateYearsUntilRetirement = (currentAge) => {

	var yearsUntilRetirement = yearOfRetirement - currentAge

	if (yearsUntilRetirement > 0) return yearsUntilRetirement; 
	else {
		console.log("This person should have retired already!")
		return 0;
	}

}

var ageYuChen = calculateAge(yearOfBirth = 1988)
console.log(ageYuChen)

var yearsLeft = calculateYearsUntilRetirement(calculateAge(1932))
console.log(yearsLeft)
console.log(calculateYearsUntilRetirement)
console.log(calculateAge)

// example of arrays
var names = ['John', 'James', 'Yu'];
var yearsOfBirth = new Array(1990, 1991, 1998);

// add some elements to arrays
names.push('Ali')
yearsOfBirth.push(1989)

console.log(names)
console.log(yearsOfBirth)

var lastName = names.pop()
var firstYear = yearsOfBirth.shift()
console.log(names)
console.log(yearsOfBirth)
console.log(lastName)
console.log(firstYear)

searchForName = (searchName) =>{
	if (names.indexOf(searchName) === -1){
		alert("This person is not in our system!")
	} else {
		alert("Yes, we've found this person!")
	}
}

//searchForName('Yu')


