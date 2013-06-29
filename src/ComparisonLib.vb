'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit
Public Enum ComparisonType
    generic = 0
    text = 1
    int32 = 2
    float32 = 3
    float64 = 4
End Enum
Private Const epsilon As Double = 10 ^ -6

Public Function float32comp(lhs As Single, RHS As Single) As Comparison
If lhs < RHS Then
    float32comp = Comparison.lessthan
ElseIf lhs > RHS Then
    float32comp = Comparison.greaterthan
Else
    float32comp = Comparison.equal
End If
End Function
Public Function float64comp(lhs As Double, RHS As Double) As Comparison
Dim delta As Double
delta = lhs - RHS
If Abs(delta) < epsilon Then
    float64comp = Comparison.equal
ElseIf delta < 0 Then
    float64comp = lessthan
Else
    float64comp = greaterthan
End If

End Function
Public Function textcomp(lhs As String, RHS As String) As Comparison
If lhs < RHS Then
    textcomp = Comparison.lessthan
ElseIf lhs > RHS Then
    textcomp = Comparison.greaterthan
Else
    textcomp = Comparison.equal
End If
End Function
Public Function floatVstext(lhs As Single, RHS As String) As Comparison
If lhs < RHS Then
    floatVstext = Comparison.lessthan
ElseIf lhs > RHS Then
    floatVstext = Comparison.greaterthan
Else
    floatVstext = Comparison.equal
End If
End Function

Public Function textVsfloat(lhs As String, RHS As Single) As Comparison
If lhs < RHS Then
    textVsfloat = Comparison.lessthan
ElseIf lhs > RHS Then
    textVsfloat = Comparison.greaterthan
Else
    textVsfloat = Comparison.equal
End If
End Function

Public Function longVslong(lhs As Long, RHS As Long) As Comparison
If lhs < RHS Then
    longVslong = Comparison.lessthan
ElseIf lhs > RHS Then
    longVslong = Comparison.greaterthan
Else
    longVslong = Comparison.equal
End If
End Function
Public Function genericComp(ByRef lhs As Variant, RHS As Variant) As Comparison
Dim ltype As VbVarType
Dim rtype As VbVarType

ltype = vartype(lhs)
rtype = vartype(RHS)
If ltype = rtype Then
    Select Case ltype
        Case VbVarType.vbDouble
            genericComp = float64comp(CDbl(lhs), CDbl(RHS))
        Case VbVarType.vbObject
            Err.Raise 101, , "Cannot compare objects, yet!"
        Case Else
            If lhs < RHS Then
                genericComp = lessthan
            ElseIf lhs > RHS Then
                 genericComp = greaterthan
            Else
                genericComp = equal
            End If
    End Select
Else
    If lhs < RHS Then
        genericComp = lessthan
    ElseIf lhs > RHS Then
         genericComp = greaterthan
    Else
        genericComp = equal
    End If
End If

End Function
Public Function compareAny(ByRef lhs As Variant, RHS As Variant, Optional comptype As ComparisonType) As Comparison
Select Case comptype
    Case ComparisonType.generic
        compareAny = genericComp(lhs, RHS)
    Case ComparisonType.text
        compareAny = textcomp(CStr(lhs), CStr(RHS))
    Case ComparisonType.int32
        compareAny = longVslong(CLng(lhs), CLng(RHS))
    Case ComparisonType.float32
        compareAny = float32comp(CSng(lhs), CSng(RHS))
    Case ComparisonType.float64
        compareAny = float64comp(CSng(lhs), CSng(RHS))
End Select
End Function

Public Function invertComparer(comp As IComparator) As IComparator
Dim nc As GenericComparerNot
Set nc = New GenericComparerNot
nc.compose comp
Set invertComparer = nc

End Function
Public Function newComparer(Optional comptype As ComparisonType, Optional descending As Boolean) As IComparator
Dim c As GenericComparer

Set c = New GenericComparer
c.comptype = comptype
If descending Then
    Set newComparer = invertComparer(c)
Else
    Set newComparer = c
End If
End Function
'make a list comparer out of all the comparers in incoll
Public Function listComparer(incoll As Collection) As IComparator
Dim comp As IComparator
Dim lc As GenericComparerList
Set lc = New GenericComparerList
For Each comp In incoll
    lc.addComparer comp
Next comp

End Function

'compare using a list of comparers
Public Function compareBy(ByRef lhs As Variant, ByRef RHS As Variant, Optional complist As Collection)
Dim res As Comparison
Dim comp As IComparator
If complist Is Nothing Then
    compareBy = compareAny(lhs, RHS)
Else
    compareBy = Comparison.equal
    For Each comp In complist
        res = comp.compare(lhs, RHS)
        If res <> Comparison.equal Then Exit For
    Next comp
    compareBy = res
End If
End Function

Public Function dictComparer(indict As Dictionary) As IComparator
Dim kvpc As GenericComparerDict
Set kvpc = New GenericComparerDict
kvpc.setKvps indict
Set dictComparer = kvpc
Set kvpc = Nothing
End Function
'builds comparers from IFn interfaces.
Public Function fnComparer(f As IFn) As IComparator
Dim c As GenericComparerFn
Set c = New GenericComparerFn
Set c.sortingf = f
Set fnComparer = c
Set c = Nothing
End Function