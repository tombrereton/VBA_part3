```VB
Option Explicit

Function removeDegreeCelsius(temperature As String) As Double

    Dim manipTemperature As String
    
    manipTemperature = WorksheetFunction.Substitute(temperature, "°C", "")
    removeDegreeCelsius = CDbl(manipTemperature)

End Function
------------------------------------------------------------------------------------------------
Function temperatureRangeTitle(rangeText As String) As String

    Dim tempTemperatureArr As Variant
    
    tempTemperatureArr = Split(rangeText, "(")
    temperatureRangeTitle = Trim(tempTemperatureArr(0))
     
End Function
------------------------------------------------------------------------------------------------

Function getTemperatureRange(rangeText As String) As String

    Dim tempTemperatureArr As Variant
    Dim tempTextString As String
    
    tempTemperatureArr = Split(rangeText, "(")
    tempTextString = Trim(tempTemperatureArr(1))
    tempTextString = WorksheetFunction.Substitute(tempTextString, " to ", "|")
    tempTextString = WorksheetFunction.Substitute(tempTextString, "°C)", "")
    
    If InStr(tempTextString, ">") Then
        tempTextString = WorksheetFunction.Substitute(tempTextString, ">", "")
        tempTextString = tempTextString & "|100"
    End If
        
    getTemperatureRange = tempTextString
    
End Function
------------------------------------------------------------------------------------------------

Function getLowerTemperature(tempRange As String) As Double

    Dim tempTemperatureArr As Variant
    Dim tempTextString As String
    
    tempTemperatureArr = Split(tempRange, "|")
    tempTextString = tempTemperatureArr(0)
    getLowerTemperature = CDbl(tempTextString)
    
End Function
------------------------------------------------------------------------------------------------

Function getHigherTemperature(tempRange As String) As Double

    Dim tempTemperatureArr As Variant
    Dim tempTextString As String
    
    tempTemperatureArr = Split(tempRange, "|")
    tempTextString = tempTemperatureArr(1)
    getHigherTemperature = CDbl(tempTextString)
    
End Function
------------------------------------------------------------------------------------------------

Function isFrozenFood(frozenFood As String) As String

    If InStr(frozenFood, "Frozen") Then
        isFrozenFood = "Frozen"
    Else
        isFrozenFood = "Refrigerated"
    End If
    
End Function
------------------------------------------------------------------------------------------------

Function currentFoodStatus(foodTemp As Double, foodTempValues As Variant) As String
    
    Dim foodTempArr As Variant
    Dim i As Integer
    
    foodTempArr = foodTempValues
    
    If foodTemp > foodTempArr(4, 2) Then
        currentFoodStatus = foodTempArr(4, 1)
    ElseIf foodTemp > foodTempArr(3, 2) Then
        currentFoodStatus = foodTempArr(3, 1)
    ElseIf foodTemp > foodTempArr(2, 2) Then
        currentFoodStatus = foodTempArr(2, 1)
    ElseIf foodTemp > foodTempArr(1, 2) Then
        currentFoodStatus = foodTempArr(1, 1)
    End If
    
       
'    For i = 4 To 1
'        If foodTemp > foodTempArr(i, 2) Then
'            currentFoodStatus = foodTempArr(i, 1)
'            Exit For
'        End If
'    Next i

End Function
------------------------------------------------------------------------------------------------

Function throwFoodAway(desStatus As String, curStatus As String) As String

    If desStatus = curStatus Then
        throwFoodAway = "Keep"
    ElseIf desStatus = "Frozen" And curStatus <> "Frozen" Then
        throwFoodAway = "Throw"
    ElseIf desStatus = "Refrigerated" And curStatus = "Excessive heat" Then
        throwFoodAway = "Throw"
    ElseIf desStatus = "Refrigerated" And curStatus = "Room temperature" Then
        throwFoodAway = "Sell today"
    ElseIf desStatus = "Refrigerated" And curStatus = "Frozen" Then
        throwFoodAway = "Throw"
    Else
        throwFoodAway = "ERROR"
    End If
    
End Function
------------------------------------------------------------------------------------------------

Sub keepTheFood()

    Application.ScreenUpdating = False

    Dim desiredStatusArr(1 To 18) As Variant, currentStatusArr(1 To 18) As Variant
    Dim TemperatureRangeArray(1 To 4, 1 To 3) As Variant
    Dim i As Integer
   
    For i = 1 To 4
        TemperatureRangeArray(i, 1) = temperatureRangeTitle(Cells(i, 9))
        TemperatureRangeArray(i, 2) = getLowerTemperature(getTemperatureRange(Cells(i, 9)))
        TemperatureRangeArray(i, 3) = getHigherTemperature(getTemperatureRange(Cells(i, 9)))
    Next i
    
    For i = 1 To 18
        desiredStatusArr(i) = isFrozenFood(Cells(i, 1))
        currentStatusArr(i) = currentFoodStatus(removeDegreeCelsius(Cells(i, 2)), TemperatureRangeArray)
        Cells(i, 3) = throwFoodAway(CStr(desiredStatusArr(i)), CStr(currentStatusArr(i)))
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

```
