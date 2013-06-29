'An interface for operating on sequences.  Unfortunately, we don't have laziness or other nice things in
'vba, but we can build some things that help when manipulating sequences of things.
'This is a hackish implementation of Clojure's ISeq, or .Net's IEnumerable.
Option Explicit

Public Property Get fst() As Variant
End Property

Public Property Get nxt() As ISeq
End Property

Public Function more() As ISeq
End Function

Public Function cons(ByRef o As Variant) As ISeq
End Function

Public Function count() As Long
End Function

Public Function seq() As ISeq
End Function