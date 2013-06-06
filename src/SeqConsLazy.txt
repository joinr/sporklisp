'Steps through a sequence, gradually realizing more and more values of
'the sequence.

'Stores realized values as a linked list.
'As more values are required, uses getnext to generate new values,
'binding them as cons cells to the realized results.

Option Explicit

Public hd As Variant
Public tl As IFn

Public realized As ISeq
Implements ISeq

Public Sub setCons(h As Variant, gen As IFn)
bind hd, h
Set tl = gen
End Sub


Private Function ISeq_cons(o As Variant) As ISeq
End Function

Private Function ISeq_count() As Long
Dim pmore As ISeq
Set pmore = ISeq_more

If pmore Is Nothing Then
    ISeq_count = 1
Else
    ISeq_count = 1 + pmore.count
End If

End Function

Private Property Get ISeq_fst() As Variant
bind ISeq_fst, hd
End Property

Private Function ISeq_more() As ISeq
Dim nextval As Variant
If realized Is Nothing Then
    If exists(tl) Then
        bind nextval, (tl.apply(list())) 'produces a cons cell
        If Not nil(nextval) Then
            Set realized = lazyCons(nextval, tl) 'this guy makes cons cells.
        Else
            Set realized = Nothing
        End If
        Set tl = Nothing 'no longer need the function.
    End If
End If

bind ISeq_more, realized

End Function

Private Property Get ISeq_nxt() As ISeq
'bind ISeq_nxt, ISeq_more
End Property

Private Function ISeq_seq() As ISeq
bind ISeq_seq, Me
End Function