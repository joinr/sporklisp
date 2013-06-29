'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.


'Integrated 29 August 2012
'A library for operating on VBA collections, which are actually fancy doubly-linked lists.
'I borrow a lot from Lisp, and try to keep list operations (operations on collections) consistent between
'languages where possible.  FYI Collections are NOT built on cons cells (as far as I know), I think they are
'actually chunked arrays underneath.
Option Explicit
Public emptyList As Collection
Public Function getEmptyList() As Collection

If Not emptyList Is Nothing Then
    Set getEmptyList = CollectionLib.emptyList
Else
    Set CollectionLib.emptyList = New Collection
    Set getEmptyList = CollectionLib.emptyList
End If

End Function
'New Function added 23 Sep 2012
Public Function copyList(incoll As Collection) As Collection
Dim itm
Set copyList = New Collection
For Each itm In incoll
    copyList.add itm
Next itm

End Function
'New Function added 23 Sep 2012
Public Function concatList(baselist As Collection, appendee As Collection) As Collection
Dim itm
Set concatList = copyList(baselist)
For Each itm In appendee
    concatList.add itm
Next itm

End Function
Public Function sortColl(incoll As Collection, Optional descending As Boolean, Optional comparef As IComparator) As Collection
Set sortColl = sortAny(incoll, descending, comparef)
End Function
Public Function sortList(inlist As Collection, Optional descending As Boolean, Optional comparef As IComparator) As Collection
Set sortList = sortAny(inlist, descending, comparef)
End Function
'Aux function added 2 Sep 2012
Public Function restList(incoll As Collection) As Collection
Dim itm As Variant
Dim head As Boolean
Set restList = New Collection

head = True
For Each itm In incoll
    If head Then
        head = False
    Else
        restList.add itm
    End If
Next itm


End Function
Public Function dropList(ByVal n As Long, incoll As Collection) As Collection
Dim itm As Variant
Set dropList = New Collection
For Each itm In incoll
    If n <= 0 Then
        dropList.add itm
    Else
        n = n - 1
    End If
Next itm

End Function

'A generic list function for creating n-tuples of values, using a list
'representation. args are zero-based.
Public Function ntuple(n As Long, ByVal args As Variant) As Collection
Dim i As Long
If UBound(args, 1) >= n - 1 Then
    Set ntuple = New Collection
    For i = LBound(args, 1) To UBound(args, 1)
        ntuple.add args(i)
    Next i
Else
    Err.Raise 101, , "Need at least " & n & _
        " args for a " & n & " tuple."
End If

End Function

'New List functions added 18 July 2012
'Creates as list of 2 values, drawn from args.
Public Function pair(ParamArray args()) As Collection
Set pair = ntuple(2, args)
End Function

'New List functions added 18 July 2012
'Creates a list of 3 values, drawn from args.
Public Function triple(ParamArray args()) As Collection
Set triple = ntuple(3, args)
End Function

'New List functions added 14 July 2012
'Partition inlist into a list of n-lists.  distance can be used to set the inter-partition
'distance.  A distance of 1 produces partitions that share n-1 elements.
'A distance of n produces partitions that share no elements.
'A distance of n-1 produces partitions that share one element, useful for deriving
'adjacency lists.
Public Function partition(inlist As Collection, n As Long, Optional distance As Long) As Collection

Dim entry As Collection
Dim i As Long
Dim k As Long
Dim j As Long
If n <= 0 Then n = 1
Set partition = New Collection
k = n - 1
If distance = 0 Then distance = n
For i = 1 To inlist.count Step distance
    If (i + k) <= inlist.count Then
        Set entry = New Collection
        For j = i To i + k
            entry.add inlist(j)
        Next j
        partition.add entry
    End If
Next i

End Function

'New List Functions Added 16 July 2012
Public Function pairs(inlist As Collection) As Collection
Set pairs = partition(inlist, 2)
End Function
'New List Functions Added 16 july 2012
Public Function triples(inlist As Collection) As Collection
Set triples = partition(inlist, 3)
End Function
Public Function adjacentPairs(inlist As Collection) As Collection
Set adjacentPairs = partition(inlist, 2, 1)
End Function

'New List functions added 14 July 2012
'Convert a paramarray into a list of arguments
Public Function arglist(ByVal inval As Variant, Optional n As Long) As Collection
Dim itm
Dim entry As Collection
Dim i As Long
Dim k As Long
Dim j As Long
If n = 0 Then n = 1
Set arglist = New Collection
k = n - 1
For i = LBound(inval, 1) To UBound(inval, 1) Step n
    If n > 1 Then
        If (i + k) <= UBound(inval, 1) Then
            Set entry = New Collection
            For j = i To i + k
                entry.add inval(j)
            Next j
            arglist.add entry
        End If
    Else
        arglist.add inval(i)
    End If
Next i
                       
Set entry = Nothing

End Function
'New List functions added 14 July 2012
'Convert a paramarray into a list of arguments
Public Function getArgs(ByVal inval As Variant) As Collection
Set getArgs = arglist(inval)
End Function
'New List functions added 14 July 2012
'Convert a paramarray into a list of arguments
Public Function getArgPairs(ByVal inval As Variant) As Collection
Set getArgPairs = arglist(inval, 2)
End Function

Sub argstest(ParamArray alist())
pprint getArgPairs(alist)

End Sub
'New list functions added 14 july 2012
Public Function fst(lst As Collection) As Variant
If lst.count > 0 Then
    If IsObject(lst(1)) Then
        Set fst = lst(1)
    Else
        fst = lst(1)
    End If
Else
    Err.Raise 101, , "List is empty!"
End If

End Function
'New list functions added 14 july 2012
Public Function snd(lst As Collection) As Variant
If lst.count > 1 Then
    If IsObject(lst(2)) Then
        Set snd = lst(2)
    Else
        snd = lst(2)
    End If
Else
    Err.Raise 101, , "List has no second entry!"
End If
End Function
'New List functions added 14 july 2012
Public Function thrd(lst As Collection) As Variant
If lst.count > 2 Then
    If IsObject(lst(3)) Then
        Set thrd = lst(3)
    Else
        thrd = lst(3)
    End If
Else
    Err.Raise 101, , "List has no third entry!"
End If

End Function
'New List functions added 30 Aug 2012
Public Function frth(lst As Collection) As Variant
If lst.count > 3 Then
    If IsObject(lst(4)) Then
        Set frth = lst(4)
    Else
        frth = lst(4)
    End If
Else
    Err.Raise 101, , "List has no fourth entry!"
End If

End Function


'new list functions added 14 july 2012
Public Function intList(upper As Long, Optional stepwidth As Long, Optional lower As Long) As Collection
Dim i As Long
If stepwidth = 0 Then stepwidth = 1
Set intList = New Collection
For i = lower To upper Step stepwidth
    intList.add i
Next i

End Function
Public Function intList1(upper As Long, Optional stepwidth As Long) As Collection
Set intList1 = intList(upper, stepwidth, 1)
End Function
'new list functions added 14 july 2012
Public Function floatList(upper As Single, Optional stepwidth As Single, Optional lower As Single) As Collection
Dim i As Single
If stepwidth = 0 Then stepwidth = 1
Set floatList = New Collection
For i = lower To upper Step stepwidth
    floatList.add i
Next i
End Function
'new list functions added 14 july 2012
Public Function floatList1(upper As Single, Optional stepwidth As Single) As Collection
Dim i As Single
If stepwidth = 0 Then stepwidth = 1
Set floatList1 = New Collection
For i = 1 To upper Step stepwidth
    floatList1.add i
Next i
End Function
'new list functions added 29 Aug 2012
'Generic numerical list
Public Function numList(upper, Optional stepwidth, Optional lower) As Collection
Dim foundfloat As Boolean

If IsMissing(stepwidth) Then stepwidth = 0
If IsMissing(lower) Then lower = 0

If isFloat(upper) Or isFloat(stepwidth) Or isFloat(lower) Then
    Set numList = floatList(CSng(upper), CSng(stepwidth), CSng(lower))
Else
    Set numList = intList(CLng(upper), CLng(stepwidth), CLng(lower))
End If

End Function
Public Function isFloat(v) As Boolean

Select Case vartype(v)
    Case VbVarType.vbCurrency, VbVarType.vbDecimal, VbVarType.vbDouble, VbVarType.vbSingle
    isFloat = True
End Select
    

End Function
'new list functions added 14 July 2012
'zips n lists, where k is the length of the shortest list, into a single list of
'k entries, where each entry is a list of n entries, drawn from the kth element
'of each input list.
Public Function zip(ParamArray lists()) As Collection
Dim itm
Dim sublist As Collection
Dim i As Long
Dim min As Long
Dim found As Boolean
Dim entry As Collection

Set zip = New Collection
min = 999999
For Each itm In lists
    If itm.count < min Then
        min = itm.count
        found = True
    End If
Next itm

If found Then
    For i = 1 To min
        Set entry = New Collection
        For Each itm In lists
            entry.add itm(i)
        Next itm
        zip.add entry
    Next i
End If

End Function
'new list functions added 14 July 2012
'decomposes a list of k entries, where each entry is is composed of n sublists,
'into n sublists where each element is drawn from the nth entry
Public Function unzip(inlist As Collection) As Collection
Dim itm
Dim sublists As Collection
Dim i As Long, j As Long
Dim min As Long
Dim found As Boolean
Dim entry As Collection

Set sublists = New Collection
Set entry = inlist(1)
For j = 1 To entry.count
    sublists.add New Collection
Next j

For i = 1 To inlist.count
    Set entry = inlist(i)
    If entry.count <> sublists.count Then Err.Raise 101, , _
        "Detected an odd list at " & i & ", the number of entries " & entry.count & " is not " & sublists.count
    For j = 1 To sublists.count
        sublists(j).add entry(j)
    Next j
Next i
    
Set unzip = sublists
Set sublists = Nothing

End Function

Public Function newColl(ParamArray invals()) As Collection
Dim itm
Set newColl = New Collection
For Each itm In invals
    list.add itm
Next itm

End Function
'alias for list, closer to lispy syntax....since we're using collections as lists.
Public Function list(ParamArray invals()) As Collection
Dim itm
Set list = New Collection
For Each itm In invals
    list.add itm
Next itm

End Function
'effectively the same as the famous cons function in lisp, prepends itm to the head of a list,
'in this case a collection.
Public Function prepend(incoll As Collection, itm As Variant) As Collection
Set prepend = incoll
If incoll.count = 0 Then
    incoll.add itm
Else
    incoll.add itm, , 1
End If

End Function
'reverse the order of the collection, returns a new collection
Public Function reverse(incoll As Collection) As Collection
Dim itm
Set reverse = New Collection

For Each itm In incoll
    Set reverse = prepend(reverse, itm)
Next itm
End Function
'return a collection of the keys in the dictionary
Public Function fromSet(inset As Dictionary) As Collection
Dim k
Set fromSet = New Collection

For Each k In inset
    fromSet.add k
Next k

End Function
'matches a string against a collection of criteria
Public Function matchString(txt As String, matchcriteria As Collection) As Boolean
Dim itm
matchString = False
For Each itm In matchcriteria
    If InStr(1, txt, CStr(itm)) > 0 Then
        matchString = True
        Exit For
    End If
Next itm
End Function
'return a collection of pairs in the dictionary
Public Function fromDict(indict As Dictionary) As Collection
Dim k
Set fromDict = New Collection

For Each k In indict
    fromDict.add list(k, indict(k))
Next k

End Function

Public Function varrayToColl(inarr() As Variant) As Collection
Dim i
Dim dcount As Long
Set varrayToColl = New Collection

dcount = dimensioncount(inarr)
With varrayToColl
    If dcount = 1 Then
        For i = LBound(inarr, 1) To UBound(inarr, 1)
            .add inarr(i)
        Next i
    ElseIf dcount = 2 Then 'likely a 2D record representation
        For i = LBound(inarr, 1) To UBound(inarr, 1)
            .add inarr(i, 1)
        Next i
    Else
        Err.Raise 101, , "Don't know how to convert 3 D arrays into collections"
    End If
End With

End Function

Sub tst()
Dim inarr2d(1 To 1, 1 To 2)
Dim inarr1d(1 To 2)
Dim tbl(1 To 2, 1 To 2)

tbl(1, 1) = "A"
tbl(1, 2) = "B"
tbl(2, 1) = "C"
tbl(2, 2) = "D"

inarr2d(1, 1) = "A"
inarr2d(1, 2) = "B"
inarr1d(1) = "C"
inarr1d(2) = "D"


End Sub

Public Function asCollection(ByRef itm As Variant) As Collection
Select Case TypeName(itm)
    Case "Collection"
        Set asCollection = itm
    Case "Variant()"
        Dim tmp()
        tmp = itm
        Set asCollection = CollectionLib.varrayToColl(tmp)
    Case Else
        Err.Raise 101, , "Cannot coerce to collection"
End Select

End Function

Public Function listToVector(Optional elements As Collection) As Variant
Dim tmp() As Variant
Dim idx As Long
Dim itm As Variant

If elements Is Nothing Then
    ReDim tmp(0 To 0)
ElseIf elements.count = 0 Then
    ReDim tmp(0 To 0)
ElseIf elements.count > 0 Then
    ReDim tmp(0 To elements.count - 1)
    
    For Each itm In elements
        bind tmp(idx), itm
        idx = idx + 1
    Next itm
End If

listToVector = tmp
End Function
