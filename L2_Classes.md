#Classes

In this lecture we covered classes and learn how to implement them in vba.

Example of VBA class:

First edit name to something like cEmployee in the`(Name)` box

<img src="http://i.imgur.com/c7wO4bz.png" alt="Drawing" height="500" />

```VB
Option Explicit

Private pName as String
Private pRate as double

'''''''''''''''
'pName property
'''''''''''''''

'Property get allows reading of the variable pName
Property get Name() as String
    Name = pName
End Property

'Property let allows writing to variable pName
Property let Name(value as string)
    pName = left(value, 30) 'limits string length to 30 characters
end property

'''''''''''''''
'pRate property
'''''''''''''''

'Property get allows reading of the variable pRate
Property get Rate() as double
    Rate = pRate
End Property

'Property let allows writing to variable pRate
Property let Rate(value as double)
    pRate = max(0, value) 'pRate must be greater than zero
end property
    
```
