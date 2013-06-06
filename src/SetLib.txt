'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'A library for working with unordered sets, based on dictionaries.
'The assumption is that set members are derived from the keys in a dictionary.
'We thus extend the union, intersection, disjoint union, and other operations to
'dictionaries, the data structure that makes VBA bearable!

'Except SetLib.addSet, these are not destructive operations.  They are pure functions of
' :: dictionary->dictionary->dictionary

Option Explicit
'addSet is mutable!
'Add is special for sets, in that we ignore items already present in the set, instead of throwing an
'error (typical dictionary behavior)
Public Function addSet(s As Dictionary, k As Variant, Optional v As Variant) As Dictionary
If Not s.exists(k) Then
    If IsMissing(v) Then v = 0
    s.add k, v
End If
Set addSet = s
End Function
Public Function emptySet() As Dictionary
Set emptySet = New Dictionary
End Function
Public Function setOfList(Optional args As Collection) As Dictionary
Dim s As Dictionary
Dim itm
Set s = emptySet()
For Each itm In args
    Set s = addSet(s, itm)
Next itm
Set setOfList = s
Set s = Nothing

End Function
Public Function union(s1 As Dictionary, s2 As Dictionary) As Dictionary
Dim smaller As Dictionary
Dim larger As Dictionary
Dim key 'variant

Set union = New Dictionary

'note, I chose to inline this.  Could generalize with sortSize, but I am wary of the possible overhead.
If s1.count < s2.count Then
    Set smaller = s1
    Set larger = s2
Else
    Set smaller = s2
    Set larger = s1
End If
   
'capture unique keys from smaller
For Each key In smaller
    If larger.exists(key) = False Then
        addSet union, key, smaller(key)
    End If
Next key

'pull in all keys from larger
For Each key In larger
    addSet union, key, larger(key)
Next

End Function
'sorts by member count, in ascending order.  returns [smaller, larger]
Public Function sortSize(s1 As Dictionary, s2 As Dictionary) As Collection
Set sortSize = New Collection

If s1.count > s2.count Then
    sortSize.add s2
    sortSize.add s1
Else
    sortSize.add s1
    sortSize.add s2
End If

End Function
Public Function intersection(s1 As Dictionary, s2 As Dictionary) As Dictionary

Dim smaller As Dictionary
Dim larger As Dictionary
Dim key 'variant

Set intersection = New Dictionary

If s1.count < s2.count Then
    Set smaller = s1
    Set larger = s2
Else
    Set smaller = s2
    Set larger = s1
End If
   
'capture keys that exist in both
For Each key In smaller
    If larger.exists(key) Then
        addSet intersection, key
    End If
Next key

End Function
Public Function difference(s1 As Dictionary, s2 As Dictionary) As Dictionary
Dim smaller As Dictionary
Dim larger As Dictionary
Dim key 'variant

Set difference = New Dictionary
   
'capture keys that exist in s1 but not s2
For Each key In s1
    If Not (s2.exists(key)) Then
        addSet difference, key
    End If
Next key
End Function
Public Function disjointUnion(s1 As Dictionary, s2 As Dictionary) As Dictionary
Dim smaller As Dictionary
Dim larger As Dictionary
Dim eliminated As Dictionary
Dim key 'variant

Set disjointUnion = New Dictionary

'note, I chose to inline this.  Could generalize with sortSize, but I am wary of the possible overhead.
If s1.count < s2.count Then
    Set smaller = s1
    Set larger = s2
Else
    Set smaller = s2
    Set larger = s1
End If
   
Set eliminated = New Dictionary
'capture unique keys from smaller
For Each key In smaller
    If larger.exists(key) = False Then
        addSet disjointUnion, key, smaller(key)
    Else 'key exists in both
        addSet eliminated, key, smaller(key)
    End If
Next key

'pull in all keys from larger, iff the key does not exist in smaller
'not sure about the effeciency here, probably could be improved.
For Each key In difference(larger, eliminated)
    addSet disjointUnion, key, larger(key)
Next

End Function

Public Sub printset(s As Dictionary)
printDict s
End Sub

Sub tst()
End Sub