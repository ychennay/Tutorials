# An Intro Into Object-Oriented Programming

One of the most important concepts to learn in computer science and programming
is **object-oriented progamming (OOP)** . The development of **object-oriented programming** is one of the greatest breakthroughs in 20th century software engineering, because it allowed programmers to work with more concrete concepts, like a `Business` or `Employee` object. This allowed software engineers to build incredibly complex applications, like the operating systems or iPhone applications we use today.

What does object-oriented programming actually look like? Think about if you wanted to save data about an employee. What type of information would you need?

## Primitives

You probably want to store the employee's name (in VBA, a `String` data type), age in years (an `Integer`), his/her salary (a `Double` or `Currency`), and many, many more. All of these are **primitives**. They're called primitives because they only have one attribute- their value. 

For example, I can declare a few variables with the operator `Dim`, and then assign values to them:

```
Dim employeeName As String
Dim employeeSalary As Double

'Assign some values to the variables now:
employeeName = "Yu Chen"
employeeSalary = 9000
```

Each variable (`employeeName`, `employeeSalary`, `primitiveVar1`, `primitiveVar2`) has only one property- its value.

## Objects

On the other hand, an **`Object`** in programming is a **collection of primitive data types** that represents a real-world entity. It is an `advanced data type` because it combines smaller data types together, and because **it can perform actions** (essentially, it can run `Sub`s itself!).

You've been working with Objects already in VBA. For example, take this familiar Object:

`ActiveCell`

When you type `ActiveCell.Value = 5`, you are setting the  