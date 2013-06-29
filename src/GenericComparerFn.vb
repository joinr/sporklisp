'sorts according to an IFn of two arguments!
Option Explicit

Public sortingf As IFn
Implements IComparator

Public Function compare(ByRef lhs As Variant, ByRef RHS As Variant) As Comparison
Dim res As Long
res = CLng(sortingf.apply(list(lhs, RHS)))
If res < 0 Then
    compare = lessthan
ElseIf res > 0 Then
    compare = greaterthan
Else
    compare = equal
End If
End Function
Private Function IComparator_compare(lhs As Variant, RHS As Variant) As Comparison
Dim res As Long
res = CLng(sortingf.apply(list(lhs, RHS)))
If res < 0 Then
    IComparator_compare = lessthan
ElseIf res > 0 Then
    IComparator_compare = greaterthan
Else
    IComparator_compare = equal
End If

End Function
