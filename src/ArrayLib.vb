'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit

Public Function dimensioncount(inarr)
Dim dimension As Long
Dim i As Long
On Error GoTo catch
i = 1
While i < 2 ^ 30
    dimension = UBound(inarr, i)
    If UBound(inarr, i + 1) Then i = i + 1
Wend

catch:
On Error Resume Next
dimensioncount = i

End Function

Public Function getSize(inarr) As Collection
Dim dimension As Long
Dim i As Long

Set getSize = New Collection

On Error GoTo catch

i = 1
While i < 2 ^ 30
    dimension = UBound(inarr, i)
    getSize.add LBound(inarr, i)
    getSize.add UBound(inarr, i)
    If UBound(inarr, i + 1) Then i = i + 1
Wend

catch:
On Error Resume Next

End Function
'squeezes a 2D record, with one row, into a 1D record
Public Function flatRecord(inarr() As Variant) As Variant()
Dim tmp()
Dim i
ReDim tmp(LBound(inarr, 2) To UBound(inarr, 2))
For i = LBound(tmp, 1) To UBound(tmp, 1)
    tmp(i) = inarr(1, i)
Next i

flatRecord = tmp
End Function

Public Function strArray(ParamArray args()) As String()
Dim tmp() As String
Dim i As Long
ReDim tmp(1 To UBound(args, 1) + 1)
For i = 1 To UBound(tmp, 1)
    tmp(i) = CStr(args(i - 1))
Next i
strArray = tmp
End Function

Sub tst()
Dim sarr() As String
Dim sngArr() As Single
Dim lngarr() As Long
Dim dblarr() As Double

sarr = strArray("A", "B", "C", "D")
lngarr = lngArray(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
sngArr = sngArray(1, 0.2, 0.03, 0.004)
dblarr = dblArray(2000000, -1, -0.00000004, 0.554)

End Sub
Private Function unpack(invarr() As Variant) As Variant

If IsArray(invarr(LBound(invarr, 1), 1)) Then
    unpack = invarr(LBound(invarr, 1), 1)
Else
    unpack = invarr
End If

End Function

Public Function lngArray(ParamArray args()) As Long()
Dim tmp() As Long
Dim i As Long
ReDim tmp(1 To UBound(args, 1) + 1)
For i = 1 To UBound(tmp, 1)
    tmp(i) = CLng(args(i - 1))
Next i
lngArray = tmp
End Function

Public Function sngArray(ParamArray args()) As Single()
Dim tmp() As Single
Dim i As Long
ReDim tmp(1 To UBound(args, 1) + 1)
For i = 1 To UBound(tmp, 1)
    tmp(i) = CSng(args(i - 1))
Next i
sngArray = tmp
End Function
Public Function dblArray(ParamArray args()) As Double()
Dim tmp() As Double
Dim i As Long
ReDim tmp(1 To UBound(args, 1) + 1)
For i = 1 To UBound(tmp, 1)
    tmp(i) = CDbl(args(i - 1))
Next i
dblArray = tmp
End Function

Public Function tostrArray(invarr() As Variant) As String()
tostrArray = strArray(invarr)

End Function

Public Function lngOfList(inlist As Collection) As Long()
Dim tmp() As Long
Dim i As Long
ReDim tmp(1 To inlist.count)
For i = 1 To inlist.count
    tmp(i) = CLng(inlist(i))
Next i
lngOfList = tmp
End Function
Public Function sngOfList(inlist As Collection) As Single()
Dim tmp() As Single
Dim i As Long
ReDim tmp(1 To inlist.count)
For i = 1 To inlist.count
    tmp(i) = CSng(inlist(i))
Next i
sngOfList = tmp
End Function

Public Function dblOfList(inlist As Collection) As Double()
Dim tmp() As Double
Dim i As Long
ReDim tmp(1 To inlist.count)
For i = 1 To inlist.count
    tmp(i) = CDbl(inlist(i))
Next i
dblOfList = tmp
End Function
Public Function strOfList(inlist As Collection) As String()
Dim tmp() As String
Dim i As Long
ReDim tmp(1 To inlist.count)
For i = 1 To inlist.count
    tmp(i) = CStr(inlist(i))
Next i
strOfList = tmp
End Function
