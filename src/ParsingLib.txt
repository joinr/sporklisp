'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'A module for parsing, and printing various data structures.
'Primarily concerned with composite structures, like dictionaries, collections, etc.
'This lib is pretty useful for interactively viewing your data, when other views (like
'the locals window) won't let you peek inside of some structures (like dictionaries).

Option Explicit
Public Function parsePrimitive()

End Function
Public Sub printColl(incoll As Collection)

Debug.Print parseColl(incoll)

End Sub
Public Sub printDict(indict As Dictionary)
Debug.Print parseDict(indict)
End Sub

Public Function parseDictP(indict As Dictionary, Optional indent As Long, Optional depth As Long, _
                                Optional col As Long, Optional line As Long) As String
Dim key
Dim i As Long
Dim str As String
Dim o As Object
Dim inner As String
Dim sb As StringBuilder
Dim nested As Boolean
Dim cnt As Long
Dim cdepth As Long
cdepth = depth

If indent = 0 Then indent = 3
If indict.count > 0 Then
    Set sb = New StringBuilder
    cnt = 0
    For Each key In indict
        cnt = cnt + 1
        If Not IsObject(indict(key)) Then
            If nested Then
                'inner = inner & " "
                sb.append " "
            End If
            If vartype(key) = vbString Then
                'inner = inner & wrapstring(CStr(key), """")
                sb.append wrapstring(CStr(key), """")
            Else
                'inner = inner & CStr(key)
                sb.append CStr(key)
            End If
            
            If vartype(indict(key)) = vbString Then
                'inner = inner & " " & wrapstring(CStr(indict(key)), """")
                sb.append " " & wrapstring(CStr(indict(key)), """")
            Else
                'inner = inner & " " & CStr(indict(key))
                sb.append " " & CStr(indict(key))
            End If
            'If cnt < indict.count Then inner = inner & "," & vbCrLf & Space(indent * cdepth)
            If cnt < indict.count Then sb.append "," & vbCrLf & Space(indent * cdepth)
            nested = True
        Else
            nested = True
            Set o = indict(key)
            'inner = inner & " " '& vbCrLf '& Space(indent * (depth + 1))
            sb.append " "
            'inner = inner & " " & Space(indent * (depth + 1))
            
            If vartype(key) = vbString Then
                'inner = inner & wrapstring(CStr(key), """")
                sb.append wrapstring(CStr(key), """")
            Else
                'inner = inner & CStr(key)
                sb.append CStr(key)
            End If
            'inner = inner & " "
            sb.append " "
            If TypeName(o) = "Collection" Then
                'inner = inner & vbCrLf & parseCollP(o, indent, cdepth + 1)
                sb.append vbCrLf & parseCollP(o, indent, cdepth + 1)
                'inner = inner & "," & vbCrLf
                sb.append "," & vbCrLf
            ElseIf TypeName(o) = "Dictionary" Then
                'inner = inner & vbCrLf & parseDictP(o, indent, cdepth + 1)
                sb.append vbCrLf & parseDictP(o, indent, cdepth + 1)
                'inner = inner & "," & vbCrLf
                sb.append "," & vbCrLf
            Else
                'inner = inner & vbCrLf & parseDictP(JSONtoDictionary(asSerial(o).asString), indent, depth + 1)
                sb.append vbCrLf & parseDictP(JSONtoDictionary(asSerial(o).asString), indent, depth + 1)
                'inner = inner & "," & vbCrLf
                sb.append "," & vbCrLf
            End If
            If cnt < indict.count Then
                'inner = inner & Space(indent * cdepth)
                sb.append Space(indent * cdepth)
            'Else
            '    inner = inner & "," & vbCrLf
            End If
        End If
    Next key
    'If Mid(inner, 1, 1) = "," Then inner = Mid(inner, 2, Len(inner) - 1)
    If Mid(sb.toString, 1, 1) = "," Then sb.Remove 1, 1
    'If Mid(Mid(inner, Len(inner) - 2, 2), 1, 1) = "," Then inner = Mid(inner, 1, Len(inner) - 3)
    inner = sb.toString
    sb.Clear
    Set sb = Nothing
    If Mid(Mid(inner, Len(inner) - 2, 2), 1, 1) = "," Then inner = Mid(inner, 1, Len(inner) - 3)
End If
str = Space(indent * cdepth) & "{" & inner & "}"
parseDictP = str
End Function
Function parseDict(indict As Dictionary, Optional indent As Long, Optional depth As Long, _
                    Optional col As Long, Optional line As Long) As String
Dim key
Dim str As String
Dim o As Object
Dim sb As StringBuilder

If indent = 0 Then indent = 3
If indict.count > 0 Then
    Set sb = New StringBuilder
    For Each key In indict
        'str = str & key & " "
        If vartype(key) = vbString Then
            'inner = inner & wrapstring(CStr(key), """")
            sb.append wrapstring(CStr(key), """")
        Else
            'inner = inner & CStr(key)
            sb.append CStr(key)
        End If
        'str = str & key & " "
        sb.append " "
        If Not IsObject(indict(key)) Then
            'str = str & indict(key)
            sb.append indict(key)
        Else
            Set o = indict(key)
            If TypeName(o) = "Dictionary" Then
                'str = str & parseDict(o, indent + indent)
                sb.append parseDict(o, indent + indent)
            ElseIf TypeName(o) = "Collection" Then
                'str = str & parseColl(o, indent + indent)
                sb.append parseColl(o, indent + indent)
            End If
        End If
        'str = str & " , " '& vbCrLf
        sb.append " , "
    Next key
    str = sb.toString
    sb.Clear
    Set sb = Nothing
    If Mid(str, Len(str) - 2) = " , " Then str = Mid(str, 1, Len(str) - 3)
End If
str = "{" & str & "}"

parseDict = str
End Function

Public Function parseColl(incoll As Collection, Optional indent As Long) As String
Dim itm
Dim i As Long
Dim str As String
Dim o As Object
Dim sb As StringBuilder

If incoll.count > 0 Then
    Set sb = New StringBuilder
    For Each itm In incoll
        If Not IsObject(itm) Then
            'str = str & " , " & CStr(itm) & " "
            If vartype(itm) = vbString Then
                'inner = inner & "," & wrapstring(CStr(itm), """")
                sb.append "," & wrapstring(CStr(itm), """")
            Else
                'inner = inner & "," & CStr(itm)
                sb.append "," & CStr(itm)
            End If
        Else
            Set o = itm
            If TypeName(o) = "Collection" Then
                'str = str & " , " & parseColl(o)
                sb.append "," & parseColl(o)
            ElseIf TypeName(o) = "Dictionary" Then
                'str = str & " , " & parseDict(o)
                sb.append " ," & parseDict(o)
            End If
        End If
    '    str = str & vbCrLf
    Next itm
    str = sb.toString
    sb.Clear
    Set sb = Nothing
    If Mid(str, 1, 1) = "," Then str = Mid(str, 2, Len(str) - 1)
End If
str = "(" & str & ")"
parseColl = str
End Function
Public Function parseCollP(incoll As Collection, Optional indent As Long, Optional depth As Long, _
                                Optional col As Long, Optional line As Long) As String
Dim itm
Dim i As Long
Dim str As String
Dim o As Object
Dim inner As String
Dim sb As StringBuilder
Dim nested As Boolean
Dim cdepth As Long

cdepth = depth

If indent = 0 Then indent = 3
If incoll.count > 0 Then
    Set sb = New StringBuilder
    For Each itm In incoll
        If Not IsObject(itm) Then
            If nested Then
                'inner = inner & vbCrLf & " "
                sb.append vbCrLf & " "
            End If
            If vartype(itm) = vbString Then
                'inner = inner & "," & wrapstring(CStr(itm), """")
                sb.append "," & wrapstring(CStr(itm), """")
            Else
                'inner = inner & "," & CStr(itm)
                sb.append "," & CStr(itm)
            End If
            nested = False
        Else
            nested = True
            Set o = itm
            If TypeName(o) = "Collection" Then
                'inner = inner & ","
                sb.append ","
                'inner = inner & vbCrLf & parseCollP(o, indent, cdepth + 1)
                sb.append vbCrLf & parseCollP(o, indent, cdepth + 1)
            ElseIf TypeName(o) = "Dictionary" Then
                'inner = inner & ","
                sb.append ","
                'inner = inner & vbCrLf & parseDictP(o, indent, cdepth + 1)
                sb.append vbCrLf & parseDictP(o, indent, cdepth + 1)
            Else 'TOM Change Nov 6 2012...
                'inner = inner & vbCrLf & parseDictP(JSONtoDictionary(asSerial(o).asString), indent, depth + 1)
                sb.append vbCrLf & parseDictP(JSONtoDictionary(asSerial(o).asString), indent, depth + 1)
                'inner = inner & "," & vbCrLf
                sb.append "," & vbCrLf
            End If
        End If
    Next itm
    inner = sb.toString
    sb.Clear
    Set sb = Nothing
    If Mid(inner, 1, 1) = "," Then inner = Mid(inner, 2, Len(inner) - 1)
End If
str = Space(indent * cdepth) & "(" & inner & ")"
parseCollP = str

End Function

'Implemented
Public Function parseSeqP(inseq As ISeq, Optional indent As Long, Optional depth As Long, _
                                Optional col As Long, Optional line As Long) As String
Dim itm
Dim i As Long
Dim str As String
Dim o As Object
Dim inner As String
Dim sb As StringBuilder
Dim nested As Boolean
Dim cdepth As Long

cdepth = depth

If indent = 0 Then indent = 3

Set sb = New StringBuilder
While exists(inseq)
    bind itm, inseq.fst
    If Not IsObject(itm) Then
        If nested Then
            'inner = inner & vbCrLf & " "
            sb.append vbCrLf & " "
        End If
        If vartype(itm) = vbString Then
            'inner = inner & "," & wrapstring(CStr(itm), """")
            sb.append "," & wrapstring(CStr(itm), """")
        Else
            'inner = inner & "," & CStr(itm)
            sb.append "," & CStr(itm)
        End If
        nested = False
    Else
        nested = True
        Set o = itm
        If TypeName(o) = "Collection" Then
            'inner = inner & ","
            sb.append ","
            'inner = inner & vbCrLf & parseCollP(o, indent, cdepth + 1)
            sb.append vbCrLf & parseCollP(o, indent, cdepth + 1)
        ElseIf isSeq(o) Then
            'inner = inner & ","
            sb.append ","
            sb.append vbCrLf & parseSeqP(o, indent, cdepth + 1)
        ElseIf TypeName(o) = "Dictionary" Then
            'inner = inner & ","
            sb.append ","
            'inner = inner & vbCrLf & parseDictP(o, indent, cdepth + 1)
            sb.append vbCrLf & parseDictP(o, indent, cdepth + 1)
        End If
    End If
    Set inseq = inseq.more
Wend
inner = sb.toString
sb.Clear
Set sb = Nothing
If Mid(inner, 1, 1) = "," Then inner = Mid(inner, 2, Len(inner) - 1)

str = Space(indent * cdepth) & "(" & inner & ")"
parseSeqP = str

End Function

Public Function wrapstring(instring As String, lwrap As String, Optional rwrap As String) As String
If rwrap = vbNullString Then rwrap = lwrap
wrapstring = lwrap & instring & rwrap
End Function

Function pad(str As String, indent As Long) As String
If indent = 0 Then
    pad = str
Else
    pad = Space(indent) & str
End If
End Function


'generic printing
Public Sub printfn(ByRef itm As Variant)

Dim tn As String
Select Case pickleType(itm)
    Case PickleTypes.primtiveJSON
        Debug.Print itm
    Case PickleTypes.compoundJSON
       tn = TypeName(itm)
       If tn = "Dictionary" Then
          printDict asDict(itm)
       ElseIf tn = "Collection" Then
          printColl asCollection(itm)
       End If
    Case PickleTypes.Serializable
        Debug.Print asSerial(asObject(itm)).asString
    Case Else
        Err.Raise 101, , "Don't know how to print " & TypeName(itm) & " generically!"
End Select
           
End Sub
Public Function asObject(ByRef itm As Variant) As Object
Set asObject = itm
End Function
'pretty print results
Public Sub pprint(ByRef itm As Variant)
Dim tn As String
Select Case pickleType(itm)
    Case PickleTypes.primtiveJSON
        If vartype(itm) = vbString Then
            Debug.Print wrapstring(CStr(itm), """")
        Else
            Debug.Print jParse(itm)
        End If
    Case PickleTypes.compoundJSON
        tn = TypeName(itm)
        If tn = "Dictionary" Then
            Debug.Print parseDictP(asDict(itm))
        ElseIf tn = "Collection" Then
            Debug.Print parseCollP(asCollection(itm))
        ElseIf isSeq(itm) Then
            Debug.Print parseSeqP(seq(itm, ISeq))
        End If
    Case PickleTypes.Serializable
        pprint JSONtoDictionary(asSerial(asObject(itm)).asString)
    Case Else
        If itm Is Nothing Then
            pprint list()
        Else
            pprint JSONtoDictionary(asSerial(asObject(itm)).asString)
        End If
        'Err.Raise 101, , "Don't know how to print " & TypeName(itm) & " generically!"
End Select
End Sub

'print results to a string
Public Function printstr(ByRef itm As Variant) As String
Dim tn As String


Select Case pickleType(itm)
    Case PickleTypes.primtiveJSON
        If vartype(itm) = vbString Then
            printstr = wrapstring(CStr(itm), """")
        Else
            printstr = jParse(itm)
        End If
    Case PickleTypes.compoundJSON
        tn = TypeName(itm)
        If tn = "Dictionary" Then
            printstr = parseDictP(asDict(itm))
        ElseIf tn = "Collection" Then
            printstr = parseCollP(asCollection(itm))
        ElseIf isSeq(itm) Then
            printstr = parseSeqP(seq(itm, ISeq))
        End If
    Case PickleTypes.Serializable
        printstr = parseDictP(JSONtoDictionary(asSerial(asObject(itm)).asString))
    Case Else
        printstr = parseDictP(JSONtoDictionary(asSerial(asObject(itm)).asString))
        'Err.Raise 101, , "Don't know how to print " & TypeName(itm) & " generically!"
End Select
End Function

