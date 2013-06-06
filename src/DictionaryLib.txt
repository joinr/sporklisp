'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'This is a stock library for working with dictionaries.
'I will probably start using it a lot to eliminate boilerplate code in
'in a lot of the dictionary stuff.
Option Explicit
'New Dictionary Functions added 16 July 2012
'Return the result of merging multiple dictionaries.
'Merging is a left->right process, in that dictionaries earlier
'in the process have their keys superceded by dictionaries later in
'the sequence.  This is like the merge function in clojure.
'We return a new dictionary the represents the result of the
'merge.  Primitive values are copied, references to objects still
'point to the same objects though, and are subject to mutation....
'The resulting dictionary is an amalgamation of the inputs.
Public Function mergeDicts(ParamArray indicts()) As Dictionary
Dim keytable As Dictionary
Dim d As Dictionary
Dim i As Long
Dim ds As Collection
Dim k

Set keytable = New Dictionary
Set ds = arglist(indicts)
'collects all the keys, from left to right, with right most overriding
'the leftmost keys.  Values are indexed accoring to the dictionary they
'came from.
For i = 1 To ds.count
    Set d = ds(i)
    For Each k In listKeys(d)
        Set keytable = assoc(keytable, k, i)
    Next k
Next i
Set mergeDicts = New Dictionary
For Each k In keytable
    Set d = ds(keytable(k))
    mergeDicts.add k, d(k)
Next k
Set keytable = Nothing
        
End Function
Private Sub mergetest()
pprint mergeDicts(newdict("A", 1, "B", 2), newdict("Tom", 30, "Spoon", 3), newdict("Tom", 31))

End Sub

'New Dictionary library adde 16 July 2012
'Fetch a collection of values, in order, associated with keys.  Keys must exist in the dictionary.
Public Function getKeys(keys As Collection, indict As Dictionary) As Collection
Dim k
Set getKeys = New Collection
For Each k In keys
    If indict.exists(k) Then
        getKeys.add indict(k)
    Else
        Err.Raise 101, , "Key " & k & " does not exist in the dictionary!"
    End If
Next k

End Function

'New function added 14 july 2012
'Returns a list of the values associated to all keys in the dictionary.
'the default .items() method of the dictionary returns a variant array, which we plug into a list
'for more accesibility.
Public Function listVals(indict As Dictionary) As Collection
Dim itm

Set listVals = New Collection
For Each itm In indict.items
    listVals.add itm
Next itm

End Function

'new function added 14 july 2012
'view a dictionary as a sequence of key-value pairs
'This is just a thin wrapper around list.fromDict
Public Function listKeyVals(indict As Dictionary) As Collection
Set listKeyVals = CollectionLib.fromDict(indict)
End Function


'Like Lisp/Clojure's assoc function.
Public Function assoc(indict As Dictionary, ByRef key As Variant, ByRef value As Variant)
If Not indict.exists(key) Then
    indict.add key, value
Else
    If IsObject(value) Then
        Set indict.item(key) = value
    Else
        indict.item(key) = value
    End If
End If

Set assoc = indict
End Function
Public Function assocIn(indict As Dictionary, keys As Collection, ByRef value As Variant) As Dictionary
Dim k
Dim i As Long
Dim ptr As Dictionary
Dim newptr As Dictionary


Set ptr = indict
For i = 1 To keys.count - 1
    k = keys(i)
    If Not ptr.exists(k) Then
        Set newptr = New Dictionary
        ptr.add k, newptr
        Set ptr = newptr
        Set newptr = Nothing
    End If
Next i
   
ptr.add keys(keys.count), value
       
Set assocIn = indict

End Function


'Allows in-line creation of dictionaries.  Truncates the last argument.
Public Function newdict(ParamArray kvps()) As Dictionary
Dim remaining As Long
Dim key
Dim val
Dim i As Long
Set newdict = New Dictionary
remaining = UBound(kvps, 1) + 1
i = 0
While remaining >= 2
    newdict.add kvps(i), kvps(i + 1)
    i = i + 2
    remaining = remaining - 2
Wend

End Function
'reads a collection, assigning k v to alternate elements.
Public Function dictOfAList(Optional assoclist As Collection) As Dictionary
Dim remaining As Long
Dim key
Dim val
Dim i As Long

Set dictOfAList = New Dictionary
If exists(assoclist) Then
    remaining = assoclist.count
    i = 1
    While remaining >= 2
        dictOfAList.add assoclist(i), assoclist(i + 1)
        i = i + 2
        remaining = remaining - 2
    Wend
End If

End Function
'reads a collection, assigning k v to alternate elements.
Public Function dictOfList(Optional list As Collection) As Dictionary
Dim val
Dim i As Long

Set dictOfList = New Dictionary
If exists(list) Then
    i = 0
    For Each val In list
        dictOfList.add i, val
        i = i + 1
    Next val
End If

End Function


Public Function newSet(ParamArray keys()) As Dictionary
Dim key
Set newSet = New Dictionary
For Each key In keys
    If Not newSet.exists(key) Then newSet.add key, key
Next key
End Function
Public Function onlyKeys(keys As Collection, indict As Dictionary) As Dictionary
Dim k
Dim filter As Dictionary
If keys.count = 0 Then
    Set onlyKeys = indict
Else
    Set onlyKeys = New Dictionary
    Set filter = New Dictionary
    For Each k In keys
        If Not filter.exists(k) Then
            filter.add k, 0
        End If
    Next k
    For Each k In filter
        If indict.exists(k) Then
            onlyKeys.add k, indict(k)
        End If
    Next k
End If
End Function
Public Function excludeKeys(keys As Collection, indict As Dictionary) As Dictionary
Dim k
Dim filter As Dictionary
If keys.count = 0 Then
    Set excludeKeys = indict
Else
    Set excludeKeys = New Dictionary
    Set filter = New Dictionary
    For Each k In keys
        If Not filter.exists(k) Then
            filter.add k, 0
        End If
    Next k
    For Each k In indict
        If Not filter.exists(k) Then
            excludeKeys.add k, indict(k)
        End If
    Next k
End If

End Function
Public Function getKey(key As Variant, indict As Dictionary) As Variant
If indict.exists(key) Then
    If IsObject(indict(key)) Then
        Set getKey = indict(key)
    Else
        getKey = indict(key)
    End If
Else
    Err.Raise 101, , "Dictionary does not contain key " & CStr(key)
End If
End Function

Public Function addDict(indict As Dictionary, ParamArray kvps()) As Dictionary
Dim remaining As Long
Dim key
Dim val
Dim i As Long


remaining = UBound(kvps, 1) + 1
i = 0
While remaining >= 2
    indict.add kvps(i), kvps(i + 1)
    i = i + 2
    remaining = remaining - 2
Wend

Set addDict = indict

End Function
Public Function copyDict(indict As Dictionary) As Dictionary

'Note -> using memcpy is probably much much faster, this is a naive way to do it.
Dim key
Set copyDict = New Dictionary
For Each key In indict
    copyDict.add key, indict(key)
Next key

End Function
'usage...
Public Sub tst()
Dim d1 As Dictionary
Dim d2 As Dictionary
 
Set d1 = addDict(newdict("A", 2), "B", 3)
printDict d1
Set d2 = copyDict(d1)
printDict d2

End Sub
'zipMap is a useful function, common in functional programming, that acts like the functional zip,
'drawing from two sequences bound to keys and values.  In this context, the sequences are combined to
'create key/value entries in a new dictionary, which is returned.  Draws incrementally from both
'collections, terminating when one runs out first
Public Function zipMap(ks As Collection, vs As Collection) As Dictionary

Dim k
Dim count As Long
Set zipMap = New Dictionary
count = count + 1

If exists(ks) And exists(vs) Then
    While count <= ks.count And count <= vs.count
        zipMap.add ks(count), vs(count)
        count = count + 1
    Wend
End If

End Function

Public Function asDict(ByRef itm As Variant) As Dictionary
Set asDict = itm
End Function


Public Function listKeys(indict As Dictionary) As Collection
Dim itm
Set listKeys = New Collection
For Each itm In indict.keys
    listKeys.add itm
Next itm

End Function