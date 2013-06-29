'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'Integrated 30 Aug 2012.

' A place to dump generic sorting routines, mostly just quicksort.  I could throw in heapsort,
' but probably don't need it.
'The functions in this module will return sorted values for just about anything that composess a sequence
'of values.  For flat collections, it'll just sort generically by value.  For dictionaries, we
'will sort the keys, returning a list of keys in sorted order.  In cases where we want to
'use dictionary values as the sorting criteria, we call sortByVal, which composes a dictionary
'comparer (basically a pass through function), to evaluate the comparison based on the
'value in the dictionary associated with they key.

'Note: Objects, arrays, collections, and other compound types cannot be used for comparison.  We must
'have primitive, or atomic values that can be compared.  In the case of nested lists, we sort on the
'first value, which must be something we can compare.

'We're using variants a bit here, so uber performance is doubtful.  However, the routines are really
'generic, and can be used on tons of structures.  If you can dump your stuff to a collection or dictionary,
'you can sort it.  If you provide a comparison function, via a class that implements IComparator, you can
'also have custom comparisons on objects.
Option Explicit
Public Function sortAny(ByRef invals As Variant, Optional descending As Boolean, _
                            Optional comparef As IComparator) As Variant
Select Case vartype(invals)
    Case VbVarType.vbObject
        Select Case TypeName(invals)
            Case "Collection"
                Set sortAny = sortList(asCollection(invals), descending, comparef)
            Case "Dictionary"
                Set sortAny = sortKeys(asDict(invals), descending, comparef)
            Case Else
                Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
        End Select
    Case Else
        Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
End Select

End Function
Public Function sortByKey(ByRef invals As Variant, Optional descending As Boolean, _
                            Optional comparef As IComparator) As Variant
Select Case vartype(invals)
    Case VbVarType.vbObject
        Select Case TypeName(invals)
            Case "Collection"
                Set sortByKey = sortList(asCollection(invals), descending, comparef)
            Case "Dictionary"
                Set sortByKey = sortKeys(asDict(invals), descending, comparef)
            Case Else
                Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
        End Select
    Case Else
        Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
End Select
End Function
Public Function sortByVal(ByRef invals As Variant, Optional descending As Boolean, _
                            Optional comparef As IComparator) As Variant
Dim dcomp As GenericComparerDict

Select Case vartype(invals)
    Case VbVarType.vbObject
        Select Case TypeName(invals)
            Case "Collection"
                Set sortByVal = sortList(asCollection(invals), descending, comparef)
            Case "Dictionary"
                If comparef Is Nothing Then
                    Set dcomp = dictComparer(asDict(invals))
                Else
                    Set dcomp = dictComparer(asDict(invals))
                    dcomp.composeWith comparef
                End If
                
                Set sortByVal = sortKeys(asDict(invals), descending, dcomp)
                Set dcomp = Nothing

            Case Else
                Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
        End Select
    Case Else
        Err.Raise 101, , "Don't know how to sort " & TypeName(invals)
End Select
End Function
'TOM Change 24 Oct 2012 -> wrap functionality for inverting custom comparers.
Public Function acquireComparer(Optional comparef As IComparator, Optional descending As Boolean) As IComparator
If comparef Is Nothing Then
    Set acquireComparer = ComparisonLib.newComparer(, descending)
Else
    If descending Then
        Set acquireComparer = ComparisonLib.invertComparer(comparef)
    Else
        Set acquireComparer = comparef
    End If
End If

End Function

Public Function sortList(invals As Collection, Optional descending As Boolean, Optional comparef As IComparator) As Collection
Dim tmp() As Variant
Dim i As Long
Dim itm
ReDim tmp(1 To invals.count)
If TypeName(invals(1)) = "Collection" Then
    Set sortList = sort2DList(invals, descending, comparef)
Else
    If invals.count = 1 Then
        Set sortList = list(invals(1))
    Else
        i = 1
        For Each itm In invals
            If Not IsObject(invals(i)) Then
                tmp(i) = itm
            Else
                Set tmp(i) = itm
            End If
            i = i + 1
        Next itm
        
        'tom change Oct 24 2012
        'If comparef Is Nothing Then Set comparef = ComparisonLib.newComparer(, descending)
        tmp = sort(tmp, acquireComparer(comparef, descending), descending)
        Set sortList = New Collection
        For i = 1 To invals.count
            sortList.add tmp(i)
        Next i
    End If
End If

End Function
Public Function sortKeys(indict As Dictionary, Optional descending As Boolean, Optional comparef As IComparator) As Collection
Dim tmp() As Variant
Dim i As Long

tmp = indict.keys

'If comparef Is Nothing Then Set comparef = ComparisonLib.newComparer(, descending)
tmp = sort(tmp, acquireComparer(comparef, descending), descending)
Set sortKeys = New Collection
For i = LBound(tmp, 1) To UBound(tmp, 1)
    sortKeys.add tmp(i)
Next i

End Function
Public Function sort2DList(invals As Collection, Optional descending As Boolean, Optional comparef As IComparator) As Collection
Dim tmp() As Variant
Dim i As Long
Dim row As Collection
ReDim tmp(1 To invals.count, 1 To 2)
If invals.count = 1 Then
    Set sort2DList = list(invals(1))
Else
    For i = 1 To invals.count
        Set row = invals(i)
        If IsObject(row(1)) Then Err.Raise 101, , "Don't know how to sort on objects! Need primitive keys!"
        tmp(i, 1) = row(1)
        Set tmp(i, 2) = row
    Next i
    
    'If comparef Is Nothing Then Set comparef = ComparisonLib.newComparer(, descending)
    tmp = sort(tmp, acquireComparer(comparef, descending), descending)
    Set sort2DList = New Collection
    For i = 1 To invals.count
        sort2DList.add tmp(i, 2)
    Next i
End If

End Function
Private Function partitionAscending(ByRef inarr() As Variant, left As Long, right As Long, pivot As Long, Optional twod As Boolean, Optional comparef As IComparator) As Long
Dim pivotval As Variant
Dim i As Long, j As Long

If twod = False Then
    If Not IsObject(inarr(pivot)) Then
        pivotval = inarr(pivot)
    Else
        Set pivotval = inarr(pivot)
    End If
Else
    If Not IsObject(inarr(pivot, 1)) Then
        pivotval = inarr(pivot, 1)
    Else
        Set pivotval = inarr(pivot, 1)
    End If
End If
i = left
j = right
If twod = False Then
    While i <= j
        While comparef.compare(inarr(i), pivotval) = lessthan
            i = i + 1
        Wend
        While comparef.compare(inarr(j), pivotval) = greaterthan
            j = j - 1
        Wend
        If i <= j Then
            swap inarr, i, j, twod
            i = i + 1
            j = j - 1
        End If
    Wend
Else
    While i <= j
        While comparef.compare(inarr(i, 1), pivotval) = lessthan
            i = i + 1
        Wend
        While comparef.compare(inarr(j, 1), pivotval) = greaterthan
            j = j - 1
        Wend
        If i <= j Then
            swap inarr, i, j, twod
            i = i + 1
            j = j - 1
        End If
    Wend
End If

partitionAscending = i
End Function

'Private Function partitionDescending(ByRef inarr() As Variant, left As Long, right As Long, pivot As Long, Optional twod As Boolean, Optional comparef As IComparator) As Long
'Dim pivotval As Variant
'Dim i As Long, j As Long
'If twod = False Then
'    pivotval = inarr(pivot)
'Else
'    pivotval = inarr(pivot, 1)
'End If
'i = left
'j = right
'If twod = False Then
'    While i <= j
'        While comparef.compare(inarr(i), pivotval) = greaterthan
'            i = i + 1
'        Wend
'        While comparef.compare(inarr(j), pivotval) = lessthan
'            j = j - 1
'        Wend
'        If i <= j Then
'            swap inarr, i, j, twod
'            i = i + 1
'            j = j - 1
'        End If
'    Wend
'Else
'    While i <= j
'        While comparef.compare(inarr(i, 1), pivotval) = greaterthan
'            i = i + 1
'        Wend
'        While comparef.compare(inarr(j, 1), pivotval) = lessthan
'            j = j - 1
'        Wend
'        If i <= j Then
'            swap inarr, i, j, twod
'            i = i + 1
'            j = j - 1
'        End If
'    Wend
'End If
'
'partitionDescending = i
'
'End Function

Private Sub swap(ByRef inarr() As Variant, idx1 As Long, idx2 As Long, Optional twod As Boolean)
Dim tmp As Variant
Dim j As Long

If idx1 <> idx2 Then
    If twod = False Then
        If Not (IsObject(inarr(idx1))) Then
            tmp = inarr(idx1)
            inarr(idx1) = inarr(idx2)
            inarr(idx2) = tmp
        Else
            Set tmp = inarr(idx1)
            Set inarr(idx1) = inarr(idx2)
            Set inarr(idx2) = tmp
        End If
    Else
        For j = 1 To UBound(inarr, 2)
            If Not IsObject(inarr(idx1, j)) Then
                tmp = inarr(idx1, j)
                inarr(idx1, j) = inarr(idx2, j)
                inarr(idx2, j) = tmp
            Else
                Set tmp = inarr(idx1, j)
                Set inarr(idx1, j) = inarr(idx2, j)
                Set inarr(idx2, j) = tmp
            End If
        Next j
    End If
End If

End Sub

Private Sub qsort(inarr() As Variant, left As Long, right As Long, Optional twod As Boolean, Optional descending As Boolean, Optional comparef As IComparator)
Dim pivot As Long
Dim newidx As Long

If right > left Then
    pivot = (right + left) \ 2
    'If descending Then
    '    newidx = partitionDescending(inarr, left, right, pivot, twod, comparef)
    'Else
        newidx = partitionAscending(inarr, left, right, pivot, twod, comparef)
    'End If
    If left < newidx - 1 Then qsort inarr, left, newidx - 1, twod, descending, comparef
    If newidx < right Then qsort inarr, newidx, right, twod, descending, comparef
End If

End Sub
Private Function sort(ByRef inarr() As Variant, comparef As IComparator, Optional descending As Boolean) As Variant()
qsort inarr, LBound(inarr, 1), UBound(inarr, 1), dimensioncount(inarr) = 2, descending, comparef
sort = inarr
End Function

Private Sub sortTest()
Dim testlist As Collection
Dim testdict As Dictionary
Set testlist = list(list(500, "This Was First!"), _
                    list(-4, "This was Second!"), _
                    list(1, "This was Third!"))
                    
pprint "unsorted list:"
pprint testlist

pprint "sorted list:"
pprint sortAny(testlist)

pprint "sorted list descending:"
pprint sortAny(testlist, True)


Set testdict = newdict("Tom", 30, "Bill", 200, "Ed", -100)

pprint "testdict : "
pprint testdict

pprint "unsorted dictkeys: "
pprint listKeys(testdict)

pprint "sorted dictkeys: "
pprint sortAny(testdict)

pprint "sorted dictkeys, by dictionary values: " & vbCrLf
pprint sortByVal(testdict)


End Sub
