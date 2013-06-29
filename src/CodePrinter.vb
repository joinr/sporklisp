'Class that parses VBA code and pretty prints it with html formatting.
Option Explicit
Private reserved As Dictionary
Private separable As Dictionary

Private Enum ParseState
    code
    ResWord
    Comment
End Enum
Private Sub Class_Initialize()
Set reserved = New Dictionary
Set separable = New Dictionary

addReserved "AddHandler"
addReserved "AddressOf"
addReserved "Alias"
addReserved "And"
addReserved "AndAlso"
addReserved "As"
addReserved "Boolean"
addReserved "ByRef"
addReserved "Byte"
addReserved "ByVal"
addReserved "Call"
addReserved "Case"
addReserved "catch"
addReserved "CBool"
addReserved "CByte"
addReserved "CChar"
addReserved "CDate"
addReserved "CDec"
addReserved "CDbl"
addReserved "Char"
addReserved "CInt"
addReserved "Class"
addReserved "CLng"
addReserved "CObj"
addReserved "Const"
addReserved "Continue"
addReserved "CSByte"
addReserved "CShort"
addReserved "CSng"
addReserved "CStr"
addReserved "CType"
addReserved "CUInt"
addReserved "CULng"
addReserved "CUShort"
addReserved "Date"
addReserved "Decimal"
addReserved "Declare"
addReserved "default"
addReserved "Delegate"
addReserved "Dim"
addReserved "DirectCast"
addReserved "Do"
addReserved "Double"
addReserved "Each"
addReserved "Else"
addReserved "ElseIf"
addReserved "End"
addReserved "End If"
addReserved "Enum"
addReserved "Erase"
addReserved "Error"
addReserved "Event"
addReserved "Exit"
addReserved "False"
addReserved "Finally"
addReserved "For"
addReserved "Friend"

addReserved "Function"
addSeparable "Function"

addReserved "Get"
addReserved "GetType"
addReserved "Global"
addReserved "GoSub"
addReserved "GoTo"
addReserved "Handles"
addReserved "If"
addReserved "Implements"
addReserved "Imports"
addReserved "In"
addReserved "Inherits"
addReserved "Integer"
addReserved "Interface"
addReserved "Is"
addReserved "IsNot"
addReserved "Let"
addReserved "Lib"
addReserved "Like"
addReserved "Long"
addReserved "Loop"
addReserved "Me"
addReserved "Mod"
addReserved "Module"
addReserved "MustInherit"
addReserved "MustOverride"
addReserved "MyBase"
addReserved "MyClass"
addReserved "Namespace"
addReserved "Narrowing"
addReserved "New"
addReserved "Next"
addReserved "Not"
addReserved "Nothing"
addReserved "NotInheritable"
addReserved "NotOverridable"
addReserved "Object"
addReserved "Of"
addReserved "On"
addReserved "Operator"
addReserved "Option"
addReserved "Optional"
addReserved "Or"
addReserved "OrElse"
addReserved "Overloads"
addReserved "Overridable"
addReserved "Overrides"
addReserved "ParamArray"
addReserved "Partial"
addReserved "Private"

addReserved "Property"
addSeparable "Property"

addReserved "Protected"
addReserved "Public"
addReserved "RaiseEvent"
addReserved "ReadOnly"
addReserved "ReDim"
addReserved "Rem"
addReserved "RemoveHandler"
addReserved "Resume"
addReserved "Return"
addReserved "SByte"
addReserved "Select"
addReserved "Set"
addReserved "Shadows"
addReserved "Shared"
addReserved "Short"
addReserved "Single"
addReserved "Static"
addReserved "Step"
addReserved "Stop"
addReserved "String"
addReserved "Structure"

addReserved "Sub"
addSeparable "Sub"

addReserved "SyncLock"
addReserved "Then"
addReserved "Throw"
addReserved "To"
addReserved "True"
addReserved "Try"
addReserved "TryCast"
addReserved "TypeOf"
addReserved "Variant"
addReserved "Wend"
addReserved "UInteger"
addReserved "ULong"
addReserved "UShort"
addReserved "Using"
addReserved "When"
addReserved "While"
addReserved "Widening"
addReserved "With"
addReserved "WithEvents"
addReserved "WriteOnly"
addReserved "Xor"

End Sub
Public Sub addReserved(res As String)
reserved.add res, 0
End Sub
Public Sub addSeparable(res As String)
separable.add res, 0
End Sub
Public Function isRes(wrd As String) As Boolean
isRes = reserved.exists(wrd)
End Function
Public Function isSeperable(wrd As String) As Boolean
isSeperable = separable.exists(wrd)
End Function

'Got a simpler way to do this...
'split will parse a line into an array, delimited by spaces.
'if the entry in the array is "" then it's a space.
'we parse a line by splitting it along spaces.
'For each word in the words, we determine if special formatting is required.
    'remarks are special.
    'keywords are special.
    'anything else is not special.

Private Function parseline(line As String, lnumber As Long) As String
Dim words
Dim i As Long
Dim wrd As String
Dim state As ParseState

state = code

words = Split(line, " ")
parseline = separate(words) & "<pre>"
parseline = parseline & newline(lnumber)

i = LBound(words, 1)
While state <> Comment And i <= UBound(words, 1)
    wrd = CStr(words(i))
    If isRes(wrd) Then
        state = ResWord
    ElseIf iscomment(wrd) Then
        state = Comment
    Else
        state = code
    End If
    
    parseline = parseline & parseword(wrd, state, True)
    i = i + 1
Wend

If i <= UBound(words, 1) Then
    While i <= UBound(words, 1)
        wrd = CStr(words(i))
        parseline = parseline & parseword(wrd, code, False)
        i = i + 1
    Wend
End If

If (state = Comment) Or (state = ResWord) Then parseline = parseline & "</span>"

parseline = parseline & "</pre>"

End Function
Private Function separate(words) As String

If UBound(words, 1) > 1 Then
    If checkseparation(CStr(words(0)), CStr(words(1))) Then
        separate = "<hr />" & vbCrLf
    Else
        separate = ""
    End If
End If

End Function
Private Function checkseparation(word1 As String, word2 As String) As Boolean
If SuborFuncorProp(word1) Then
    checkseparation = True
Else
    If isRes(word1) And isRes(word2) And SuborFuncorProp(word2) Then checkseparation = True
End If

End Function
Private Function SuborFuncorProp(word As String) As Boolean
SuborFuncorProp = (word = "Sub" Or word = "Function" Or word = "Property")
End Function
Private Function iscomment(word As String) As Boolean
iscomment = (InStr(1, word, "'") > 0)
End Function
Private Function newline(number As Long) As String

Dim spaces As String
Dim numstr As String

numstr = CStr(number)
If Len(numstr) > 4 Then
    Err.Raise 101, "too many lines of code"
Else
    spaces = Space(4 - Len(numstr))
    numstr = spaces & numstr
    newline = "<span class=""lnum"">" & numstr & ":  </span>"
End If

End Function
Private Function parseword(word As String, state As ParseState, Optional newcomment As Boolean) As String

Select Case state
    Case Is = ParseState.code
        parseword = " " & word
    Case Is = ParseState.ResWord
        parseword = reservedword(word)
    Case Is = ParseState.Comment
        If newcomment Then
            parseword = startcomment(word)
        Else
            parseword = " " & word
        End If
End Select
        
End Function
Private Function reservedword(word As String) As String
reservedword = "<span class=""kwrd"">" & " " & word & "</span>"
End Function
Private Function startcomment(Optional word As String) As String
startcomment = "<span class=""rem"">" & word
End Function
Private Function closeSpan() As String
closeSpan = "</span>"
End Function
'returns the parsed html of the pretty-printed code.  TimeStamped!
Public Function parsemodule(module As VBComponent, Optional path As String) As String
Dim lines As String
With module
    If .CodeModule.CountOfLines > 0 Then
        lines = CStr(Now()) & vbCrLf & .CodeModule.lines(1, .CodeModule.CountOfLines)
    Else
        lines = ""
    End If
End With

parsemodule = parsecode(lines)

End Function
'returns the parsed html of the pretty-printed code.
Public Function parsecode(code As String, Optional path As String) As String
Dim lines
Dim i As Long
Dim line As String

parsecode = startParse
lines = Split(code, vbCrLf)
For i = LBound(lines, 1) To UBound(lines, 1)
    line = CStr(lines(i))
    parsecode = parsecode & vbCrLf
    parsecode = parsecode & parseline(line, i)
Next i

parsecode = closeParse(parsecode)

End Function
Private Function startParse() As String
startParse = csshead & vbCrLf & "<div class=""csharpcode"">"
End Function
Private Function closeParse(parsed As String) As String
closeParse = parsed & vbCrLf & "</div>"
End Function
Private Function csshead() As String
csshead = "<head>" & vbCrLf _
        & "<style type=""text/css"">" & vbCrLf _
       & ".csharpcode, .csharpcode pre" & vbCrLf _
       & "{" & vbCrLf _
           & "font-size: small;" & vbCrLf _
           & "color: black;" & vbCrLf _
           & "font-family: Consolas, ""Courier New"", Courier, Monospace;" & vbCrLf _
           & "background-color: #ffffff;" & vbCrLf _
           & "/*white-space: pre;*/" & vbCrLf _
       & "}" & vbCrLf _
       & ".csharpcode pre { margin: 0em; }" & vbCrLf _
       & ".csharpcode .rem { color: #008000; }" & vbCrLf _
       & ".csharpcode .kwrd { color: #0000ff; }" & vbCrLf _
       & ".csharpcode .str { color: #006080; }" & vbCrLf _
       & ".csharpcode .op { color: #0000c0; }" & vbCrLf _
       & ".csharpcode .preproc { color: #cc6633; }" & vbCrLf _
       & ".csharpcode .asp { background-color: #ffff00; }" & vbCrLf _
       & ".csharpcode .html { color: #800000; }" & vbCrLf _
       & ".csharpcode .attr { color: #ff0000; }" & vbCrLf _
       & ".csharpcode .alt" & vbCrLf _
       & "{" & vbCrLf _
           & "background-color: #f4f4f4;" & vbCrLf _
           & "width: 100%;" & vbCrLf _
           & "margin: 0em;" & vbCrLf
           
csshead = csshead _
       & "}" & vbCrLf _
       & ".csharpcode .lnum { color: #606060; }" & vbCrLf _
       & "</style>" & vbCrLf _
       & "</head>"
        

End Function

 

