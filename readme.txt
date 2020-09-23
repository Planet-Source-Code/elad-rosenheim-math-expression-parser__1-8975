-------------------------------
Math Expression Parser
-------------------------------
Class & Demo by Elad Rosenheim
e-mail: eladro@barak-online.net
-------------------------------

The purpose of this code is to show you how to write simple, yet very
powerful and flexible parsers for all kinds of data. The source can be
easily extended to support many more features - My aim was just to show you the basics.

To make a (really) long story short:
The first step in creating a parser is defining exactly what the
syntax of legal expressions is. The common way to describe complex, recursive syntaxes is by using a hierarchical structure.

In a mathematical expression, there are precedence levels the different operators, so we'll define a different level in the hierarchy for each operator "class". 
We'll also give each level a name.

In the top-level there's the "MathExpression". At this level we take care of +|- operations between "Term" values. Term is the level where
we handle * | / operations, which must be carried out before +|-
operations. Each value we handle in the "Term" level can be a 
math expression by itself, enclosed in parentheses, or it may be just a  simple number...

The language I implemented a parser for looks like this:

MathExpression: Term [+|- Term]*
Term: Value ['*'|/ Value]*
Value: SubExpression|Atom
SubExpression: '(' MathExpression ')'
Atom: Number|Constant|Function
Number: 0-9[0-9]*[.][0-9]*
Constant:SymbolName
Function:SymbolName '(' MathExpression ')'
SymbolName: BeginLetter[Letter]*
BeginLetter: [A-Z|a-z]
Letter:BeginLetter|0-9|'_'

Legend:
-------
A|B - Alternation - Can be A OR (instead of) B
[xxx] - Optional - may or may not appear
* - The last element can appear any number of times

Yes, I know it all may seem very strange and unclear. Look at the demo, follow the code and even insert some syntax errors (like "1+1+"
or "sin(50") to see how they are handled.

I hope you get it - It's a very interesting field.
