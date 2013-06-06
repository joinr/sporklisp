Option Explicit

Public source As IComparator
Implements IComparator

Public Function compose(comp As IComparator) As IComparator
Set source = comp
Set compose = Me
End Function
Public Function compare(ByRef lhs As Variant, ByRef RHS As Variant) As Comparison
compare = source.compare(RHS, lhs)
End Function

Private Function IComparator_compare(lhs As Variant, RHS As Variant) As Comparison
IComparator_compare = compare(lhs, RHS)
End Function