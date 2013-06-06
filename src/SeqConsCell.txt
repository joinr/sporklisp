'Cons cells are the basis for lists in lisp, and for the generic Seq implementation.
Option Explicit

'final public class Cons extends ASeq implements Serializable {

Private pFirst As Variant
Private pmore As ISeq
Implements ISeq

Public Sub setCons(ByRef first As Variant, Optional ByRef more As ISeq)
bind pFirst, first
If exists(more) Then bind pmore, more
End Sub
Public Sub setRest(ByRef r As ISeq)
Set pmore = r
End Sub

Public Function first() As Variant
bind first, pFirst
End Function

Public Function getNext() As ISeq
If more Is Nothing Then
    Set getNext = Nothing
Else
    Set getNext = more().seq
End If
End Function

Public Function more() As ISeq

Set more = pmore
End Function
'makes a new cons cell
Public Function cons(itm As Variant) As SeqConsCell
Dim sc As SeqConsCell
Set sc = New SeqConsCell
sc.setCons itm

Set pmore = sc
Set sc = Nothing
Set cons = Me

End Function

Public Function count() As Long

If pmore Is Nothing Then
    count = 1
Else
    count = 1 + pmore.count
End If

End Function

Private Function ISeq_cons(ByRef o As Variant) As ISeq
Dim c As SeqConsCell
Set c = New SeqConsCell
c.setCons o, Me
Set ISeq_cons = c
Set c = Nothing
End Function

Private Function ISeq_count() As Long
ISeq_count = count()
End Function

Private Property Get ISeq_fst() As Variant
bind ISeq_fst, pFirst
End Property

Private Function ISeq_more() As ISeq
Set ISeq_more = getNext()
End Function

Private Property Get ISeq_nxt() As ISeq
Set ISeq_nxt = pmore
End Property

Private Function ISeq_seq() As ISeq
Set ISeq_seq = Me
End Function