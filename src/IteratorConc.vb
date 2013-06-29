'An iterator for seqs that builds concatenations.  It concatenates two sequences, producing
'a stream of cons cells that exhausts the first sequence, then the second.
Option Explicit

Private myseq As ISeq
Private moreseqs As ISeq

Implements IFn

Public Sub create(sequences As ISeq)
Set myseq = SeqLib.lazyseq(seq(sequences.fst))
Set moreseqs = SeqLib.lazyseq(sequences.more)

End Sub
'returns cons cells
'This guy iterates over any seq.
Public Function getNext() As Variant 'return a lazy list.

If exists(myseq) Then
    bind getNext, myseq.fst
    Set myseq = myseq.more
    If Not exists(myseq) Then
        Set myseq = moreseqs.fst  'advance to the next sequence.
        Set moreseqs = moreseqs.more 'pop the next sequence off...
    End If
End If

End Function


Private Function IFn_apply(args As Collection) As Variant
bind IFn_apply, getNext()
End Function

