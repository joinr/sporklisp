'interface for comparing 2 items.
'this is fairly typical in most languages
Public Enum Comparison
    lessthan = -1
    equal = 0
    greaterthan = 1
End Enum
Option Explicit
Public Function compare(ByRef lhs As Variant, ByRef RHS As Variant) As Comparison

End Function