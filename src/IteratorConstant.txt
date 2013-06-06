'An iterator that produces the same value, or repeatedly evaluates an expression.
'Never ends...
Option Explicit
Private myval As Variant
Private myfunc As IFn
Private isfunc As Boolean
Implements IFn
Public Sub repeatVal(ByRef inval As Variant)
bind myval, inval
isfunc = False
End Sub
Public Sub repeatedlyFunc(ByRef infunc As IFn)
Set myfunc = infunc
isfunc = True
End Sub
'returns cons cells
'This guy iterates over any seq.
Public Function getNext() As Variant 'return a lazy list.

If isfunc = False Then
    bind getNext, myval
Else
    bind getNext, myfunc.apply(CollectionLib.getEmptyList())
End If

End Function

Private Function IFn_apply(args As Collection) As Variant
If isfunc = False Then
    bind IFn_apply, myval
Else
    bind IFn_apply, myfunc.apply(CollectionLib.getEmptyList())
End If
End Function

