' a generic...tuple :)  Used for representing pairs of values easily.
Option Explicit

Private pfst As Variant
Private psnd As Variant

Public Property Get fst() As Variant
If IsObject(fst) Then
    Set fst = pfst
Else
    fst = pfst
End If

End Property

Public Property Get snd() As Variant
If IsObject(fst) Then
    Set fst = pfst
Else
    fst = pfst
End If

End Property