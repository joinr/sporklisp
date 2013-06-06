'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'A library for operating on generic sequences via the ISeq interface.
'Note: sequences are boxed by nature, as a result, the heavy use of variants will probably slow things
'down or create garbage collection overhead.  VBA does some wierd stuff with its memory management, and
'I have seen space leaks in the past.  We'll have to profile and see.
Public Enum SeqTypes
    unknown = 0
    primitiveSeq = 1
    dictionarySeq = 2
    collectionSeq = 3
    vArraySeq = 4
    multiArraySeq = 5
    listArraySeq = 6
    ISeq = 100
End Enum
Public Enum SeqVal
    EmptySeq = -90929387#
End Enum
  
Option Explicit
'eats a reference and turns it into a sequence.
'for single values, this more or less returns a list (a SeqConsCell).
Public Function seq(ByRef coll As Variant, Optional t As SeqTypes) As ISeq
Dim st As SeqTypes

If t = unknown Then
    t = seqtype(coll, st)
End If
st = t

Dim c As Collection
Dim d As Dictionary

If st = SeqTypes.unknown Then
    Err.Raise 101, , "Don't know what type of seq this is."
Else
    Select Case st
        Case SeqTypes.ISeq
            Set seq = coll
        Case SeqTypes.primitiveSeq
            Set seq = primitiveToSeq(coll)
        Case SeqTypes.collectionSeq
            Set c = coll
            Set seq = collToSeq(c)
            Set c = Nothing
        Case SeqTypes.dictionarySeq
            Set d = coll
            Set seq = dictToSeq(d)
            Set d = Nothing
    End Select
End If

End Function
Public Function seqtype(ByRef s As Variant, Optional t As SeqTypes) As SeqTypes
Dim tn As String
seqtype = t 'allow short circuiting.
If seqtype = unknown Then
    Select Case vartype(s)
        Case vbVariant, vbArray, vbArray + vbVariant, vbNull, vbDate, vbString, vbBoolean, vbDouble, vbDecimal, _
                vbInteger, vbLong, vbSingle
            seqtype = primitiveSeq
        Case vbObject
            tn = TypeName(s)
            Select Case tn
                Case "ISeq", "SeqColl", "SeqConsCell", "SeqConsLazy", "SeqDict", "SeqArr"
                    seqtype = SeqTypes.ISeq
                Case "Dictionary"
                    seqtype = dictionarySeq
                Case "Collection"
                    seqtype = collectionSeq
                Case "Variant()"
                    seqtype = multiArraySeq
                Case "ISerializable"
                    seqtype = SeqTypes.ISeq
            End Select
    End Select
End If
End Function
Public Function isSeq(ByRef s As Variant) As Boolean
isSeq = seqtype(s) <> unknown
End Function
Public Function makeCons(ByRef v As Variant, Optional more As Variant) As SeqConsCell
Set makeCons = New SeqConsCell
If IsMissing(more) Then
    makeCons.setCons v
Else
    makeCons.setCons v, seq(more)
End If

End Function
Public Function primitiveToSeq(ByRef p As Variant) As ISeq
Set primitiveToSeq = makeCons(p)
End Function
Public Function collToSeq(incoll As Collection) As ISeq
Set collToSeq = collToCons(incoll)
'collToSeq.fromColl incoll
End Function
Public Function dictToSeq(indict As Dictionary) As SeqDict
Set dictToSeq = New SeqDict
'dictToSeq.fromdict indict
End Function
Public Function dictSeq(indict As Dictionary) As ISeq
End Function
Public Function collSeq(incoll As Collection) As ISeq
'Dim sc As SeqColl
'Set sc = New SeqColl
'sc.fromColl incoll
'Set collSeq = sc
'Set sc = Nothing

Set collSeq = collToCons(incoll)

End Function

Public Function collToCons(incoll As Collection) As SeqConsCell
Dim itm
Dim sc As SeqConsCell
Dim previous As SeqConsCell
Dim hd As SeqConsCell

For Each itm In incoll
    If previous Is Nothing Then
        Set sc = makeCons(itm)
        Set previous = sc
        Set hd = sc
    Else
        Set sc = previous.cons(itm)
        Set previous = sc.getNext
    End If
Next itm

Set collToCons = hd
Set hd = Nothing
Set previous = Nothing

End Function
'Converts an ISeq to a VBA collection.
Public Function seqToColl(s As ISeq) As Collection
Dim elmt As ISeq
Set seqToColl = New Collection
Set elmt = s
Do
    seqToColl.add elmt.fst
    If exists(elmt.more) Then
        Set elmt = elmt.more
    Else
        Exit Do
    End If
Loop
    
End Function
'Fetch the first element of the sequence, logically equivalent to the head of a list.
Public Function first(s As ISeq) As Variant
bind first, s.fst
End Function
'Return a sequence that logically represents the rest of the sequence. Logically
'equivalent to the tail of a list.
Public Function rest(s As ISeq) As ISeq
bind rest, s.more
End Function
'Return the first element of the rest of the sequence.  Logically equivalent to
'getting the head of the tail of the sequence.
Public Function seqNext(s As ISeq) As Variant
bind seqNext, first(rest(s))
End Function
'Return the nth element of the sequence, if it exists.
Public Function nth(n As Long, s As ISeq) As Variant
Dim link As ISeq
Dim ct As Long

assert n - n, 0 'make sure n is positive or 0

If n = 0 Then
    bind nth, s.fst
Else
    Set link = s
    ct = 0
    While exists(link) And ct < n
        Set link = link.more
        ct = ct + 1
    Wend
    If ct = n Then
        bind nth, link.fst
    End If
End If
    
End Function
'Returns the length of the sequence.
Public Function count(s As ISeq) As Long
count = s.count
End Function
Public Function arrSeq(ByRef invarr As Variant) As ISeq
End Function
'Returns a new sequence that concatenates N sequences, where the
'concatenation is represented by a new cons cell mapping the tail of
'the first sequence to the head of the next.  Non destructive.
Public Function concat(ParamArray seqs()) As ISeq
Dim c As IteratorConc
Dim itm
Dim cs As Collection

Set c = New IteratorConc
Set cs = New Collection
For Each itm In seqs
    cs.add itm
Next itm
c.create seq(cs)
Set cs = Nothing

End Function
Function collIterator(incoll As Collection) As IFn
'(let ((count 0) (lst (incoll)) (lambda () (if
End Function
Function lazyColl(incoll As Collection) As ISeq
Dim hd As SeqConsLazy
End Function
Function lazyCons(inval As Variant, generator As IFn) As SeqConsLazy
Set lazyCons = New SeqConsLazy
lazyCons.setCons inval, generator
End Function
'generates a lazy sequence of values from the underlying ISeq.
Function lazyseq(s As ISeq, Optional todrop As Long, Optional totake As Long, _
                    Optional filter As IFn, Optional whileseq As Boolean, _
                        Optional mapf As IFn) As ISeq
Dim gen As IteratorSeq
Set gen = New IteratorSeq
gen.create s, todrop, totake, filter, whileseq, mapf
If mapf Is Nothing Then
    Set lazyseq = lazyCons(gen.getNext, gen)
Else
    Set lazyseq = lazyCons(mapf.apply(list(gen.getNext)), gen) 'transform the first value.
End If

End Function
'Produces a sequence that repeats the same value.
Function repeat(inval As Variant) As ISeq
Dim iterator As IteratorConstant
Set iterator = New IteratorConstant
iterator.repeatVal inval
Set repeat = lazyCons(inval, iterator)
Set iterator = Nothing
End Function
'Repeatedly evalutes the thunk, which is a function of no arguments.
'Assuming side-effects.
Function repeatedly(thunk As IFn) As ISeq
Dim iterator As IteratorConstant
Set iterator = New IteratorConstant
iterator.repeatedlyFunc thunk
Set repeatedly = lazyCons(thunk.apply(CollectionLib.getEmptyList), iterator)
Set iterator = Nothing
End Function
'return the next-to-last element of the sequence.
Function butLast(s As ISeq) As ISeq
Dim link As SeqConsCell
Dim prev As SeqConsCell

Set prev = s

Do
    If exists(prev.more) Then
        Set link = prev.more
        If link.more Is Nothing Then
            Exit Do 'prev is our guy
        Else
            Set prev = link 'advance prev
        End If
    Else
        Exit Do
    End If
Loop
    
Set butLast = prev
Set link = Nothing
Set prev = Nothing
End Function
'return the last element of the sequence.
Function last(s As ISeq) As ISeq
Dim link As SeqConsCell

Set link = s
While exists(link.more)
    Set link = link.more
Wend
    
Set last = link
Set link = Nothing

End Function
'take the first n elements of the sequence. If n exceeds the count of the
'sequence, returns only those taken as a sequence.
Function take(n As Long, s As ISeq) As ISeq
Set take = lazyseq(s, , n)
End Function
Function drop(n As Long, s As ISeq) As ISeq
Set drop = lazyseq(s, n)
End Function
'Returns each element in s, where (fn element) is true.
Function filter(fn As IFn, s As ISeq) As ISeq
Set filter = lazyseq(s, , , fn)
End Function
'Returns a lazy sequence of elements in s, where (fn element) is
'true.  Sequence terminates on the first false value.
Function takewhile(fn As IFn, s As ISeq) As ISeq
Set takewhile = lazyseq(s, , , fn, True)
End Function
'builds a function that repeatedly evaluates the last value, based on
'an initial value, initval.
Function makeIterator(fn As IFn, initval As Variant) As IFn
Dim iterator As LispClosure
Set iterator = _
    makefn("(lambda (func initval) " & _
              "(let ((f func) " & _
                    "(lastval initval)) " & _
                        "(lambda () " & _
                            "(do (set! lastval (f lastval)) " & _
                                 "lastval))))")
                            
Set makeIterator = iterator.apply(list(fn, initval))
Set iterator = Nothing
End Function
'maps function fn to the sequence of values s, lazily returning a new sequence of values.
Function seqMap(fn As IFn, s As ISeq) As ISeq
Set seqMap = lazyseq(s, , , , , fn)
End Function
'Reverses seqeuence s.  If any elements have not been forced, will eval them....
Function seqReverse(s As ISeq) As ISeq
Dim c As Collection
Set c = CollectionLib.reverse(seqToColl(s))
Set seqReverse = collToSeq(c)
Set c = Nothing
End Function
'TOM added 9 Nov 2012
Function seqSortBy(fn As IFn, s As ISeq) As ISeq
Dim c As Collection
Set c = seqToColl(s)
Set seqSortBy = collToSeq(CollectionLib.sortColl(c, , ComparisonLib.fnComparer(fn)))
Set c = Nothing
End Function
'Attemps to sort the sequence using generic sorting functions.
Function seqSort(s As ISeq, Optional descending As Boolean) As ISeq
Dim c As Collection
Set c = seqToColl(s)
Set seqSort = collToSeq(CollectionLib.sortColl(c, descending))
Set c = Nothing
End Function
'TOM added 9 Nov 2012
'Conjoins itm onto collection...destructive...
Function conj(coll As ISeq, ByRef itm As Variant) As ISeq
Set conj = coll.cons(itm)
End Function

Function seqRange(n As Long) As ISeq
Set seqRange = reval("(take-while (fn (x) (< x " & n & ")) (iterate inc 0)))")
End Function

'returns a lazy sequence of elements, where the first element is
'the initial value, the second element is the result of applying
'f to the initial value, then f(f(initial)
Function iterate(fn As IFn, initval As Variant) As ISeq
Set iterate = lazyCons(initval, makeIterator(fn, initval))
End Function

'''map over a lazy sequence.
'''generate a lazy sequence of (lazy-cons (f (first s)) (map (f rest s)))
''Function lazymap(fn As IFn, s As ISeq) As ISeq
''Dim mapf As IFn
''Set mapf = makefn("(lambda (f hd) " & _
''                     "(let ((s hd)) " & _
''                       "(lambda () " & _
''                          "(if (nil? hd) " & _
''                            "nil " & _
''                            "(do " & _
''                                "(set! (f x)))").apply(list(fn))
''
''End Function

Sub seqtest()

Dim coll As Collection
Dim s As ISeq
Dim cons As SeqConsCell
Dim seqL As ISeq
Set coll = list(1, 2, 3)

Set s = seq(coll)
Set seqL = lazyseq(s)

End Sub
Sub gentest()
Dim nums As ISeq
Dim ten As ISeq
Dim evens As ISeq
Dim odds As ISeq
Dim lessthanfive As ISeq
Dim firstthree As ISeq
Dim butfirstthree As ISeq

Set nums = lazyCons(0, makefn("(let ((x 0)) (lambda () (do (set! x (+ x 1)) x)))")) 'this is an infinite list
Set ten = take(10, nums)
Set evens = filter(makefn("(lambda (x) (even? x))"), ten)
Set odds = filter(makefn("(lambda (x) (odd?  x))"), ten)

Set lessthanfive = takewhile(makefn("(lambda (x) (< x 5))"), nums)
Set firstthree = take(3, nums)
Set butfirstthree = drop(3, ten)


End Sub

Sub itertest()
Dim fn As IFn
Set fn = makeIterator(makefn("(lambda (x) (+ x 1))"), 0)

End Sub

Sub iteratetest()
Dim s As ISeq
Set s = iterate(makefn("(lambda (x) (+ x 1))"), 0)
End Sub

Sub maptest()
Dim s As ISeq
Set s = seqMap(getFunc("inc"), seqRange(10))

End Sub

Sub revsorttest()
Dim s As ISeq
Set s = take(10, repeatedly(lambda(list(), "(rand-between 10 20)")))
pprint seqSort(s, True)

End Sub
