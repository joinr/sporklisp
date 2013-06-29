'if we have a collection, let's just turn the collection into a series o cons cells.
'we might be stitching multiple colls together.
'we want sequences to be immutable (i.e. they read from the underlying values in the collection, but
'don't modify the collection).

'As we traverse the seq, i.e as seq.first, seq.next are called, we build a chain of cons cells.
'each new cons cell will have the current idx of the collection as its value, and have the more of
'the cons point back to the collseq.
'As as the collseq gets new requests for nxt, it'll add to the idx val, adjusting the value returned
'by first.  When we return a new first val, we really return a new conscell.

'I think that'll do it.

Option Explicit
Implements ISeq
'Private d As Dictionary
'Private keys As Collection
'Private ct As Long
'Private idx As Long
Public Function fromDict(dict As Dictionary)
'Set c = incoll
'idx = 1
'ct = incoll.count
End Function
Private Function ISeq_cons(ByRef o As Variant) As ISeq
'If c Is Nothing Then
'    Set c = New Dictionary
'    idx = 1
'End If
'
'c.add o
End Function
Private Function ISeq_count() As Long
'ISeq_count = ct
End Function
Property Get ISeq_fst() As Variant
'If c Is Nothing Then
'    Set ISeq_fst = Nothing
'Else
'    bind ISeq_fst, (c.item(idx))
'End If
    
End Property
Private Function ISeq_more() As ISeq

End Function
Private Property Get ISeq_nxt() As ISeq
'If c Is Nothing Then
'    Set ISeq_nxt = Nothing
'Else
'    If idx = c.count Then
'        'sequence has been consumed.
'        Set c = Nothing 'release the sequence
'        Set ISeq_nxt = Nothing
'    Else
'        idx = idx + 1
'        Set ISeq_nxt = Me
'    End If
'End If
End Property
Private Function ISeq_seq() As ISeq
'Set ISeq_seq = Me
End Function