#Classes

In this lecture we covered classes and learn how to implement them in vba.

Example of VBA class:

First edit name to something like cEmployee in the`(Name)` box

<img src="http://i.imgur.com/c7wO4bz.png" alt="Drawing" height="500" />

```VB
Option Explicit

Private pName As String
Private pRate As Double
Private NormalHrs As Integer
Private OvertimeHrs As Integer

'''''''''''''''
'Name property
'''''''''''''''

'Property get allows reading of the variable pName
Property Get Name() As String
    Name = pName
End Property

'Property let allows writing to variable pName
Property Let Name(value As String)
    pName = Left(value, 30) 'limits string length to 30 characters
End Property

'''''''''''''''
'Rate property
'''''''''''''''

'Property get allows reading of the variable pRate
Property Get Rate() As Double
    Rate = pRate
End Property

'Property let allows writing to variable pRate
Property Let Rate(value As Double)
    pRate = WorksheetFunction.Max(0, value) 'pRate must be greater than zero
End Property

'''''''''''''''
'FirstName property
'''''''''''''''

Property Get FirstName() As String
    'finds the location of the first space and then splits the string at that point to return the first name
    FirstName = Left(pName, WorksheetFunction.Search(" ", pName) - 1)
End Property

'''''''''''''''
'HoursPerWeek property
'''''''''''''''

'writes in hours per week as one simple number
Property Let HoursPerWeek(value As Integer)
    'if the hours per week are greater than 35 it is considered overtime
    NormalHrs = WorksheetFunction.Min(35, value)
    OvertimeHrs = WorksheetFunction.Max(0, value - 35)
End Property

'''''''''''''''
'WeeklyPay function
'''''''''''''''

Public Function WeeklyPay() As Double
    WeeklyPay = NormalHrs * pRate + OvertimeHrs * pRate * 1.5
End Function
    
```

Then to use this class we run a subroutine:

```VB
Option Explicit

Sub testemp()

    'instantiate cEmployee as emp
    Dim emp As New cEmployee
    
    'the class properties must be set
    emp.Name = "Ralph McWiggins"
    emp.Rate = 15
    emp.HoursPerWeek = 42
    
    'the employees weekly pay is desribed in a sentence
    MsgBox emp.FirstName & "'s weekly pay is £" & emp.WeeklyPay & "/wk"

End Sub
```
