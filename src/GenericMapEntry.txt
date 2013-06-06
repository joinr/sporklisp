Option Explicit

Public key As Variant
Public val As Variant

'public Object nth(int i)
'{   if(i == 0)      return key();   else if(i == 1)     return val();
'    else        throw new IndexOutOfBoundsException();}
Public Function nth(i As Long) As Variant

If i = 0 Then
    bind nth, key
ElseIf i = 1 Then
    bind nth, val
Else
    Err.Raise 101, , "Index out of Bounds"
End If

End Function

'private IPersistentVector asVector()
'{   return LazilyPersistentVector.createOwning(key(), val());}


'public IPersistentVector assocN(int i, Object val){ return asVector().assocN(i, val);}

'public int count()
'{   return 2;}
Public Function count() As Long
count = 2
End Function
Public Function seq() As ISeq

End Function
'public ISeq seq()
'{   return asVector().seq();}
'public IPersistentVector cons(Object o){
'    return asVector().cons(o);}
'public IPersistentCollection empty()
'{   return null;}
'public IPersistentStack pop()
'{   return LazilyPersistentVector.createOwning(key());}
'public Object setValue(Object value)
'{   throw new UnsupportedOperationException();}