<!-- TOC -->

- [An Intro Into Object-Oriented Programming](#an-intro-into-object-oriented-programming)
    - [Primitives](#primitives)
    - [Objects](#objects)
        - [Objects Have (Multiple) Properties](#objects-have-multiple-properties)
        - [Objects Can Perform Actions](#objects-can-perform-actions)
        - [Using `Dim` vs `Set`](#using-dim-vs-set)

<!-- /TOC -->

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

### Objects Have (Multiple) Properties

You've been working with Objects already in VBA. For example, take this familiar Object:

`ActiveCell`

When you type `ActiveCell.Font`, you're actually accessing the `Font` object of `ActiveCell`. And this `Font` object has a primitive property called `Bold`, which is a `Boolean` (either a `True` or `False`).

So when you type `ActiveCell.Font.Bold = True`, what you're actually saying is 

*Take the `ActiveCell` object, and give me its `Font` object. Inside the `Font` object, set its primitive property `Bold` to `True`.*

What about this?

`ActiveCell.Interior.ColorIndex = 43`

This sets the `ColorIndex` integer primitive property of `Interior`, which is itself a property of `ActiveCell` to the value of 43.

### Objects Can Perform Actions

Objects not only have multiple primitive properties, but they can also ** do stuff **. It's as if they have their own `Sub`s written just for them.

For example, `ActiveCell.Delete`. This `ActiveCell` object can perform an action called `Delete`. It deletes its current value. That's just one of the **many** actions that it can perform:
```
ActiveCell.Cut
ActiveCell.Paste
ActiveCell.Activate
ActiveCell.Find
```

Here's a table of the `ActiveCell` object, and its `properties` and `methods`:

![Methods](/VBA/Images/activecell.png)

An action that a object can perform is called a **`method`**. In VBA, you can tell which attributes of an object are methods and which ones are properties by looking at the icon:

![Methods](/VBA/Images/methods.png)

### Using `Dim` vs `Set`

`Dim` is used to **declare** a primitive variable or object. When you write 

`Dim MyVar As Double`

this is essentially telling the computer, *hey, I want you to reserve a block of disk space and memory to store a numerical value, and call this `MyVar`*.

However, it does not actually provide `MyVar` with a value. You have to next **assign** the variable a value:

```
Dim strVariable As String
strVariable = "Yu Chen"
```

For Objects, it's very similar, except for one key difference - you need to use the `Set` operator:

```
Dim myRange as Range
Set myRange = Sheet1.Range("A1")
```

What type of Object is being declared here? A `Range` object - you know, the a collection of cells in an Excel worksheet. When you use `Set`, you need to make sure that both the object on the left and right side of the `=` are the `same type of object`.

You can check what type of object by using `TypeName()`:

```
Dim Ctrl As Control = New TextBox
MsgBox(TypeName(Ctrl))
```

One final note: do not use `Set` to assign primitive variables values. It won't work! `Set` is only used to create new Objects!