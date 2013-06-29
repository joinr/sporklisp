Option Explicit

Public idx As Long
Public source As Dictionary

Implements IFn

Public Sub reset()
idx = 0
End Sub
Public Function getNext() As Variant
If idx + 1 <= source.count Then
    idx = idx + 1
    bind getNext, source(idx)
Else
    Set getNext = Nothing
End If

End Function

Private Sub Class_Terminate()
Set source = Nothing
End Sub

Private Function IFn_apply(args As Collection) As Variant
bind IFn_apply, getNext()
End Function