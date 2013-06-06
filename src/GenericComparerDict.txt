'A comparer that consults a dictionary for its comparison.
'Specifically, we pass the values, mapped from keys, as the
'comparison values, rather than sorting on the keys explicitly.
Option Explicit

Private kvps As Dictionary
Private composed As IComparator
Implements IComparator

Public Sub setKvps(keyvals As Dictionary)
Set kvps = Nothing
Set kvps = keyvals
End Sub
Public Sub composeWith(newcomp As IComparator)
Set composed = Nothing
Set composed = newcomp
End Sub
Public Function compare(ByRef lkey As Variant, ByRef rkey As Variant) As Comparison

If kvps.exists(lkey) And kvps.exists(rkey) Then
    If composed Is Nothing Then
        compare = ComparisonLib.compareAny(kvps(lkey), kvps(rkey))
    Else
        compare = composed.compare(kvps(lkey), kvps(rkey))
    End If
Else
    Err.Raise 101, , "Keys do not exist in the Dictionary you are using for comparison!"
End If

End Function

Private Sub Class_Terminate()
Set kvps = Nothing
End Sub

'use the list of comparers to try to reach a conclusion about these guys
Private Function IComparator_compare(lhs As Variant, RHS As Variant) As Comparison
IComparator_compare = compare(lhs, RHS)
End Function