'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Public disableScreenRequests As Long
Public disableCalcRequests As Long

Option Explicit
'simple function to convert a 2D variant array with 1 row into a CSV record
Public Function varrToCSV(inarr() As Variant) As String
Dim j As Long

j = LBound(inarr, 2)
varrToCSV = inarr(1, j)
For j = 2 To UBound(inarr, 2)
    varrToCSV = varrToCSV & "," & inarr(1, j)
Next j

End Function
Public Function maxInt(ParamArray invals()) As Long
Dim i As Long
i = LBound(invals, 1)
maxInt = CLng(invals(i))

For i = (i + 1) To UBound(invals, 1)
    If CLng(invals(i)) > maxInt Then maxInt = CLng(invals(i))
Next i


End Function
Public Function minInt(ParamArray invals()) As Long
Dim i As Long
i = LBound(invals, 1)
minInt = CLng(invals(i))

For i = (i + 1) To UBound(invals, 1)
    If CLng(invals(i)) < minInt Then minInt = CLng(invals(i))
Next i

End Function
Public Function maxFloat(ParamArray invals()) As Single
Dim i As Long
i = LBound(invals, 1)
maxFloat = invals(i)

For i = (i + 1) To UBound(invals, 1)
    If invals(i) > maxFloat Then maxFloat = invals(i)
Next i

End Function

Public Function minFloat(ParamArray invals()) As Single
Dim i As Long
i = LBound(invals, 1)
minFloat = invals(i)

For i = (i + 1) To UBound(invals, 1)
    If invals(i) < minFloat Then minFloat = invals(i)
Next i

End Function
'return the index of the item with the minimum value.
'maybe we do this with a collection....
Function MinList(ParamArray nums() As Variant) As Collection
Dim i As Long
Dim mins As Collection
Set mins = New Collection
Dim minimum As Single
minimum = 9999999
For i = LBound(nums, 1) To UBound(nums, 1)
    If nums(i) < minimum Then
        minimum = nums(i)
        mins.add nums(i)
    End If
Next i
Set MinList = New Collection
For i = mins.count To 1 Step -1
    MinList.add mins(i)
Next i
End Function

Public Sub noUnders()
Dim cl As Range
Dim inrng As Range
Set inrng = Selection
For Each cl In inrng
    cl.value = Trim(replace(cl.value, "_", vbNullString))
Next cl
    
End Sub
Public Sub CharCodes()
Dim i As Long
Dim total As Long
total = 132
Dim str

str = vbNullString
i = 7
While total > 0 And i >= 0
    If 2 ^ i < total Then
        total = total - 2 ^ i
        str = str & "1"
    Else
        str = str & "0"
    End If
    i = i - 1
Wend
Debug.Print str

End Sub
'Cut down on the verbage for all the set blah = nothing
Public Sub dispose(ParamArray objs())
Dim itm
For Each itm In objs
    Set itm = Nothing
Next itm
End Sub
Public Function minLong(ParamArray nums()) As Long
Dim i As Long
minLong = CLng(nums(0))
For i = LBound(nums, 1) + 1 To UBound(nums, 1)
    If CLng(nums(i)) < minLong Then minLong = CLng(nums(i))
Next i
End Function

Public Function timeStamp() As String
timeStamp = CStr(Now())
End Function
'Shut down all rendering and calculating, for performance reasons.
Public Sub MuzzleExcel()
DisableScreenUpdates
DisableCalculations
End Sub
'Request to enable all rendering and calculating, returning excel to normal.
Public Sub UnMuzzleExcel()
EnableScreenUpdates
EnableCalculations
End Sub
Public Sub EnableCalculations()
If disableCalcRequests > 0 Then
    disableCalcRequests = disableCalcRequests - 1
End If

If disableCalcRequests = 0 Then
    Application.Calculation = xlCalculationAutomatic
End If

End Sub
Public Sub DisableCalculations()
If disableCalcRequests = 0 Then
    Application.Calculation = xlCalculationManual
End If
disableCalcRequests = disableCalcRequests + 1
End Sub
Public Sub EnableScreenUpdates()
If disableScreenRequests > 0 Then
    disableScreenRequests = disableScreenRequests - 1
End If

If disableScreenRequests = 0 Then
    Application.ScreenUpdating = True
End If

End Sub
Public Sub DisableScreenUpdates()
If disableScreenRequests = 0 Then
    Application.ScreenUpdating = False
End If
disableScreenRequests = disableScreenRequests + 1
End Sub

Public Function isScreenDisabled() As Boolean
isScreenDisabled = disableScreenRequests = 0
End Function

Sub niltest()
Dim blah As Collection
pprint nil(blah)
pprint nil(list(1, 2, 3))
pprint nil(newdict(1, 2))
pprint nil("A")

End Sub
Public Function exists(ByRef inval) As Boolean
exists = Not (inval Is Nothing)
End Function
Sub existstest()
Dim blah As Collection
pprint exists(blah)
pprint exists(list(1, 2, 3))
End Sub
Public Function nil(Optional ByRef inval) As Boolean

If IsMissing(inval) Then
    nil = True
Else
    Select Case vartype(inval)
        Case VbVarType.vbEmpty, VbVarType.vbNull
            nil = True
        Case VbVarType.vbObject
            nil = inval Is Nothing
        Case Else
            nil = False
    End Select
End If

End Function
'Moved from TimeStep_Engine, these are general utilities for using the status bar.
'TOM Change
Public Sub beginlog(msg As String)
Application.DisplayStatusBar = True
Application.StatusBar = msg
End Sub
'TOM Change
Public Sub logStatus(msg As String)
Application.StatusBar = msg
End Sub
'TOM Change
Public Sub endlog(msg As String)

Application.StatusBar = msg
Application.StatusBar = "Idle"

End Sub


Public Function rangeKeyVals(inrng As Range) As String
Dim nm As String
Dim rw As Range
Dim np As Dictionary
nm = inrng.Cells(1, 1).value

Set np = New Dictionary
For Each rw In inrng.rows
    np.add CStr(rw.Cells(1, 1).value), CStr(rw.Cells(1, 2).value)
Next rw

rangeKeyVals = printstr(np)
Set np = Nothing

End Function

Public Function rangeDictionary(inrng As Range) As Dictionary
Dim nm As String
Dim rw As Range
Dim np As Dictionary
nm = inrng.Cells(1, 1).value

Set np = New Dictionary
For Each rw In inrng.rows
    np.add CStr(rw.Cells(1, 1).value), CStr(rw.Cells(1, 2).value)
Next rw

Set rangeDictionary = np
Set np = Nothing

End Function

Public Function delimitedColumn(invals As Collection) As String
Dim itm
For Each itm In invals
    delimitedColumn = delimitedColumn & CStr(itm) & vbTab & vbCrLf
Next itm
End Function
