'sorts in ascending order by default!
'descending behavior is accomplished by composing with a not
Option Explicit

Public comptype As ComparisonType
Implements IComparator

Private Sub Class_Initialize()
comptype = ComparisonType.generic
End Sub
Public Function compare(ByRef lhs As Variant, ByRef RHS As Variant) As Comparison
compare = compareAny(lhs, RHS, comptype)
End Function
Private Function IComparator_compare(lhs As Variant, RHS As Variant) As Comparison
IComparator_compare = compare(lhs, RHS)
End Function