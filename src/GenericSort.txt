'This WAS a generic class for sort operations, with an implementation of quicksort.  I repurposed it
'to just be an Excel wrapper (ugh) to hide the nasty excel calling language for sorting tables of
'cells.

'Option Explicit
'
'Private Function partitionAscending(inarr(), left As Long, right As Long, pivot As Long, Optional twod As Boolean) As Long
'Dim pivotval
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
'        While inarr(i) < pivotval
'            i = i + 1
'        Wend
'        While inarr(j) > pivotval
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
'        While inarr(i, 1) < pivotval
'            i = i + 1
'        Wend
'        While inarr(j, 1) > pivotval
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
'partitionAscending = i
'
'End Function
'Private Function partitionDescending(inarr(), left As Long, right As Long, pivot As Long, Optional twod As Boolean) As Long
'Dim pivotval
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
'        While inarr(i) > pivotval
'            i = i + 1
'        Wend
'        While inarr(j) < pivotval
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
'        While inarr(i, 1) > pivotval
'            i = i + 1
'        Wend
'        While inarr(j, 1) < pivotval
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
'
'Private Sub swap(inarr(), idx1 As Long, idx2 As Long, Optional twod As Boolean)
'Dim tmp
'Dim j As Long
'
'If idx1 <> idx2 Then
'    If twod = False Then
'        tmp = inarr(idx1)
'        inarr(idx1) = inarr(idx2)
'        inarr(idx2) = tmp
'    Else
'        For j = 1 To UBound(inarr, 2)
'            tmp = inarr(idx1, j)
'            inarr(idx1, j) = inarr(idx2, j)
'            inarr(idx2, j) = tmp
'        Next j
'    End If
'End If
'
'End Sub
'
'Private Sub qsort(inarr(), left As Long, right As Long, Optional twod As Boolean, Optional descending As Boolean)
'Dim pivot As Long
'Dim newidx As Long
'
'If right > left Then
'    pivot = (right + left) \ 2
'    If descending Then
'        newidx = partitionDescending(inarr, left, right, pivot, twod)
'    Else
'        newidx = partitionAscending(inarr, left, right, pivot, twod)
'    End If
'    If left < newidx - 1 Then qsort inarr, left, newidx - 1, twod, descending
'    If index < right Then qsort inarr, newidx, right, twod, descending
'End If
'
'End Sub
''we only sort 1d or 2d arrays...
'Public Function sort(source(), Optional column As Long, Optional descending As Boolean) As Double()
'
'Dim dcount As Long
'dcount = ArrayLib.dimensioncount(source)
'If dcount > 2 Then Err.Raise 101, , "Cannot sort in greater than 2 dimensions currently, check your input array"
'
'If dcount = 2 Then
'    If column = 0 Then column = 1
'    qsort source, LBound(rendered, 1), UBound(rendered, 1), True, descending
'sort = rendered
'End Function
'

'call the built-in sorting functions to sort a range....only sorts by a single column.
Public Sub SortRange(inrng As Range, Optional bycolumn As Long, Optional ascending As Boolean)
Dim sorttype As Long
Dim found As Boolean
Dim col As Long
Dim clmn As Range

If ascending Then sorttype = xlAscending Else sorttype = xlDescending

If bycolumn > 0 Then
    col = bycolumn
Else
    col = 1 'sort using the first column
End If

inrng.sort inrng.Cells(1, col), sorttype

End Sub
'assumes a table with fields, ignores the headers.  Assumes first row is header row.
Public Sub SortTableRange(inrng As Range, byfield As String, Optional ascending As Boolean)
SortRange chopHeaders(inrng), findField(inrng, byfield), ascending
End Sub

Private Function findField(inrng As Range, field As String) As Long
Dim clmn As Range

findField = 1
For Each clmn In inrng.Columns
    If clmn.Cells(1, 1) = field Then
        found = True
        Exit For
    Else
        findField = findField + 1
    End If
Next clmn

If Not found Then Err.Raise 101, , "field " & field & " does not exist"
End Function

Private Function chopHeaders(inrng As Range) As Range
Set chopHeaders = inrng.offset(1, 0)
Set chopHeaders = chopHeaders.resize(chopHeaders.rows.count - 1, chopHeaders.Columns.count)
End Function

Private Function getHeaders(inrng As Range) As Range
Set getHeaders = inrng.rows(1)
End Function