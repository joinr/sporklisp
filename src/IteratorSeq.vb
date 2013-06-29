'An iterator for seqs that lazily generates cons cells, that is, consumers of the sequence
'receive values on-demand.
Option Explicit

Private myseq As ISeq
Private remaining As Long
Private drop As Long
Private filter As IFn
Private takewhile As Boolean
Private mapFn As IFn

Implements IFn

Public Sub create(s As ISeq, Optional todrop As Long, Optional totake As Long, _
                        Optional fltr As IFn, Optional tw As Boolean, Optional mapf As IFn)
Set myseq = s
drop = todrop

Set filter = fltr
If totake > 0 Then
    remaining = totake
Else
    remaining = -99 'exhaust list
End If

takewhile = tw

Set mapFn = mapf

End Sub
'returns cons cells
'This guy iterates over any seq.
Public Function getNext() As Variant 'return a lazy list.
Dim link As ISeq

If drop > 0 Then
    Do
        If exists(myseq) And drop > 0 Then
            Set myseq = myseq.more
            drop = drop - 1
        Else
            drop = 0
            Exit Do
        End If
    Loop
End If

If remaining > 0 Then
    If exists(myseq) Then
        If exists(filter) Then
            Do
                If exists(myseq) Then
                    If filter.apply(list(myseq.fst)) = False Then
                        If takewhile Then
                            Set myseq = Nothing
                            remaining = 0
                            Exit Do
                        Else
                            Set myseq = myseq.more
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
        End If
        If exists(myseq) Then
            bind getNext, myseq.fst
            Set myseq = myseq.more
            remaining = remaining - 1
        End If
    Else
        remaining = 0
    End If
ElseIf remaining = -99 Then
    If exists(myseq) Then
        If exists(filter) Then
            Do
                If exists(myseq) Then
                    If filter.apply(list(myseq.fst)) = False Then
                        If takewhile Then
                            Set myseq = Nothing
                            remaining = 0
                            Exit Do
                        Else
                            Set myseq = myseq.more
                        End If
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            If exists(myseq) Then
                bind getNext, myseq.fst
                Set myseq = myseq.more
            End If
        Else
            bind getNext, myseq.fst 'realizes a val, passes the generator
            Set myseq = myseq.more
        End If
    Else
        remaining = 0
    End If
Else
    Set myseq = Nothing 'let the seq get garbage collected.
End If

End Function

Private Function IFn_apply(args As Collection) As Variant
Dim res As Variant
bind res, getNext()
If Not (mapFn Is Nothing) Then
    If Not nil(res) Then
        bind IFn_apply, mapFn.apply(list(res))
    Else
        bind IFn_apply, res
    End If
Else
    bind IFn_apply, res
End If

End Function
