'Tom note -> This is an external lib.  Not my favorite coding conventions...might port it.
'I moved the original code from a VB6 module into a class, and adapted some of the sinew.  Notably,
'this class is fairly well wrapped behind a generic serialization module, called, Serialization.
'However, you can just as easily use it by itself.  It depends on the StringBuilder class for
'fast string building (really fast actually).

' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

Option Explicit



Private Const INVALID_JSON      As Long = 1
Private Const INVALID_OBJECT    As Long = 2
Private Const INVALID_ARRAY     As Long = 3
Private Const INVALID_BOOLEAN   As Long = 4
Private Const INVALID_NULL      As Long = 5
Private Const INVALID_KEY       As Long = 6
Private Const INVALID_RPC_CALL  As Long = 7

Private psErrors As String

Public Function GetParserErrors() As String
   GetParserErrors = psErrors
End Function

Public Function ClearParserErrors() As String
   psErrors = ""
End Function


'
'   parse string and create JSON object
'
Public Function parse(ByRef str As String) As Object

   Dim index As Long
   index = 1
   psErrors = ""
   On Error Resume Next
   Call skipChar(str, index)
   Select Case Mid(str, index, 1)
      Case "{"
         Set parse = parseObject(str, index)
      Case "["
         Set parse = parseArray(str, index)
      Case Else
         psErrors = "Invalid JSON"
   End Select


End Function

 '
 '   parse collection of key/value
 '
Private Function parseObject(ByRef str As String, ByRef index As Long) As Dictionary

   Set parseObject = New Dictionary
   Dim sKey As String
   
   ' "{"
   Call skipChar(str, index)
   If Mid(str, index, 1) <> "{" Then
      psErrors = psErrors & "Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf
      Exit Function
   End If
   
   index = index + 1

   Do
      Call skipChar(str, index)
      If "}" = Mid(str, index, 1) Then
         index = index + 1
         Exit Do
      ElseIf "," = Mid(str, index, 1) Then
         index = index + 1
         Call skipChar(str, index)
      End If

      
      ' add key/value pair
      sKey = parseKey(str, index)
      On Error Resume Next
      
      parseObject.add sKey, parseValue(str, index)
      If Err.number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & sKey & vbCrLf
         Exit Do
      End If
   Loop
eh:

End Function

'
'   parse list
'
Private Function parseArray(ByRef str As String, ByRef index As Long) As Collection

   Set parseArray = New Collection

   ' "["
   Call skipChar(str, index)
   If Mid(str, index, 1) <> "[" Then
      psErrors = psErrors & "Invalid Array at position " & index & " : " + Mid(str, index, 20) & vbCrLf
      Exit Function
   End If
   
   index = index + 1

   Do

      Call skipChar(str, index)
      If "]" = Mid(str, index, 1) Then
         index = index + 1
         Exit Do
      ElseIf "," = Mid(str, index, 1) Then
         index = index + 1
         Call skipChar(str, index)
      End If

      ' add value
      On Error Resume Next
      parseArray.add parseValue(str, index)
      If Err.number <> 0 Then
         psErrors = psErrors & Err.Description & ": " & Mid(str, index, 20) & vbCrLf
         Exit Do
      End If
   Loop

End Function

'
'   parse string / number / object / array / true / false / null
'
Private Function parseValue(ByRef str As String, ByRef index As Long)

   Call skipChar(str, index)

   Select Case Mid(str, index, 1)
      Case "{"
         Set parseValue = parseObject(str, index)
      Case "["
         Set parseValue = parseArray(str, index)
      Case """", "'"
         parseValue = parseString(str, index)
      Case "t", "f"
         parseValue = parseBoolean(str, index)
      Case "n"
         parseValue = parseNull(str, index)
      Case Else
         parseValue = parseNumber(str, index)
   End Select

End Function

'
'   parse string
'
Private Function parseString(ByRef str As String, ByRef index As Long) As String

   Dim quote   As String
   Dim Char    As String
   Dim code    As String

   Dim sb As New StringBuilder

   Call skipChar(str, index)
   quote = Mid(str, index, 1)
   index = index + 1
   
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
         Case "\"
            index = index + 1
            Char = Mid(str, index, 1)
            Select Case (Char)
               Case """", "\", "/", "'"
                  sb.append Char
                  index = index + 1
               Case "b"
                  sb.append vbBack
                  index = index + 1
               Case "f"
                  sb.append vbFormFeed
                  index = index + 1
               Case "n"
                  sb.append vbLf
                  index = index + 1
               Case "r"
                  sb.append vbCr
                  index = index + 1
               Case "t"
                  sb.append vbTab
                  index = index + 1
               Case "u"
                  index = index + 1
                  code = Mid(str, index, 4)
                  sb.append ChrW(val("&h" + code))
                  index = index + 4
            End Select
         Case quote
            index = index + 1
            
            parseString = sb.toString
            Set sb = Nothing
            
            Exit Function
            
         Case Else
            sb.append Char
            index = index + 1
      End Select
   Loop
   
   parseString = sb.toString
   Set sb = Nothing
   
End Function

'
'   parse number
'
Private Function parseNumber(ByRef str As String, ByRef index As Long)

   Dim value   As String
   Dim Char    As String

   Call skipChar(str, index)
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      If InStr("+-0123456789.eE", Char) Then
         value = value & Char
         index = index + 1
      Else
         If InStr(value, ".") Or InStr(value, "e") Or InStr(value, "E") Then
            parseNumber = CDbl(value)
         Else
            parseNumber = CLng(value)
         End If
         Exit Function
      End If
   Loop
End Function

'
'   parse true / false
'
Private Function parseBoolean(ByRef str As String, ByRef index As Long) As Boolean

   Call skipChar(str, index)
   If Mid(str, index, 4) = "true" Then
      parseBoolean = True
      index = index + 4
   ElseIf Mid(str, index, 5) = "false" Then
      parseBoolean = False
      index = index + 5
   Else
      psErrors = psErrors & "Invalid Boolean at position " & index & " : " & Mid(str, index) & vbCrLf
   End If

End Function

'
'   parse null
'
Private Function parseNull(ByRef str As String, ByRef index As Long)

   Call skipChar(str, index)
   If Mid(str, index, 4) = "null" Then
      parseNull = Null
      index = index + 4
   Else
      psErrors = psErrors & "Invalid null value at position " & index & " : " & Mid(str, index) & vbCrLf
   End If

End Function

Private Function parseKey(ByRef str As String, ByRef index As Long) As String

   Dim dquote  As Boolean
   Dim squote  As Boolean
   Dim Char    As String

   Call skipChar(str, index)
   Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
         Case """"
            dquote = Not dquote
            index = index + 1
            If Not dquote Then
               Call skipChar(str, index)
               If Mid(str, index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case "'"
            squote = Not squote
            index = index + 1
            If Not squote Then
               Call skipChar(str, index)
               If Mid(str, index, 1) <> ":" Then
                  psErrors = psErrors & "Invalid Key at position " & index & " : " & parseKey & vbCrLf
                  Exit Do
               End If
            End If
         Case ":"
            index = index + 1
            If Not dquote And Not squote Then
               Exit Do
            Else
               parseKey = parseKey & Char
            End If
         Case Else
            If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
            Else
               parseKey = parseKey & Char
            End If
            index = index + 1
      End Select
   Loop

End Function

'
'   skip special character
'
Private Sub skipChar(ByRef str As String, ByRef index As Long)
   Dim bComment As Boolean
   Dim bStartComment As Boolean
   Dim bLongComment As Boolean
   Do While index > 0 And index <= Len(str)
      Select Case Mid(str, index, 1)
      Case vbCr, vbLf
         If Not bLongComment Then
            bStartComment = False
            bComment = False
         End If
         
      Case vbTab, " ", "(", ")"
         
      Case "/"
         If Not bLongComment Then
            If bStartComment Then
               bStartComment = False
               bComment = True
            Else
               bStartComment = True
               bComment = False
               bLongComment = False
            End If
         Else
            If bStartComment Then
               bLongComment = False
               bStartComment = False
               bComment = False
            End If
         End If
         
      Case "*"
         If bStartComment Then
            bStartComment = False
            bComment = True
            bLongComment = True
         Else
            bStartComment = True
         End If
         
      Case Else
         If Not bComment Then
            Exit Do
         End If
      End Select
      
      index = index + 1
   Loop

End Sub

Public Function toString(ByRef obj As Variant) As String
   Dim sb As New StringBuilder
   Select Case vartype(obj)
      Case vbNull
         sb.append "null"
      Case vbDate
         sb.append """" & CStr(obj) & """"
      Case vbString
         sb.append """" & Encode(obj) & """"
      Case vbObject
         
         Dim bFI As Boolean
         Dim i As Long
         Dim tname As String
         
         tname = TypeName(obj)
         
         bFI = True
         Select Case tname
            Case "Dictionary"
                sb.append "{"
                Dim keys
                keys = obj.keys
                For i = 0 To obj.count - 1
                   If bFI Then bFI = False Else sb.append ","
                   Dim key
                   key = keys(i)
                   sb.append """" & key & """:" & toString(obj.item(key))
                Next i
                sb.append "}"
             Case "Collection"
                sb.append "["
                Dim value
                For Each value In obj
                   If bFI Then bFI = False Else sb.append ","
                   sb.append toString(value)
                Next value
                sb.append "]"
            'TOM Change 6 Sep 2012
            Case "ISeq", "SeqConsCell", "SeqConsLazy"
                    sb.append "["
                    Dim nd As ISeq
                    Set nd = obj
                    While exists(nd)
                        sb.append toString(nd.fst)
                        Set nd = nd.more
                        If bFI Then bFI = False Else sb.append ","
                    Wend
                    sb.append "]"
            'Tom change 7 Nov 2012
            Case Else
                sb.append trySerialize(obj)
        End Select
      Case vbBoolean
         If obj Then sb.append "true" Else sb.append "false"
      Case vbVariant, vbArray, vbArray + vbVariant
         Dim sEB
         sb.append multiArray(obj, 1, "", sEB)
      Case Else
         sb.append replace(obj, ",", ".")
   End Select

   toString = sb.toString
   Set sb = Nothing
   
End Function

Private Function Encode(str) As String

   Dim sb As New StringBuilder
   Dim i As Long
   Dim j As Long
   Dim aL1 As Variant
   Dim aL2 As Variant
   Dim c As String
   Dim p As Boolean

   aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
   aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
   For i = 1 To Len(str)
      p = True
      c = Mid(str, i, 1)
      For j = 0 To 7
         If c = Chr(aL1(j)) Then
            sb.append "\" & Chr(aL2(j))
            p = False
            Exit For
         End If
      Next

      If p Then
         Dim a
         a = AscW(c)
         If a > 31 And a < 127 Then
            sb.append c
         ElseIf a > -1 Or a < 65535 Then
            sb.append "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
         End If
      End If
   Next
   
   Encode = sb.toString
   Set sb = Nothing
   
End Function

Private Function multiArray(aBD, iBC, sPS, ByRef SPT)   ' Array BoDy, Integer BaseCount, String PoSition
   
   Dim iDU As Long
   Dim iDL As Long
   Dim i As Long
   
   On Error Resume Next
   iDL = LBound(aBD, iBC)
   iDU = UBound(aBD, iBC)

   Dim sb As New StringBuilder

   Dim sPB1, sPB2  ' String PointBuffer1, String PointBuffer2
   If Err.number = 9 Then
      sPB1 = SPT & sPS
      For i = 1 To Len(sPB1)
         If i <> 1 Then sPB2 = sPB2 & ","
         sPB2 = sPB2 & Mid(sPB1, i, 1)
      Next
      '        multiArray = multiArray & toString(Eval("aBD(" & sPB2 & ")"))
      sb.append toString(aBD(sPB2))
   Else
      SPT = SPT & sPS
      sb.append "["
      For i = iDL To iDU
         sb.append multiArray(aBD, iBC + 1, i, SPT)
         If i < iDU Then sb.append ","
      Next
      sb.append "]"
      SPT = left(SPT, iBC - 2)
   End If
   Err.Clear
   multiArray = sb.toString
   
   Set sb = Nothing
End Function

' Miscellaneous JSON functions

Public Function StringToJSON(st As String) As String
   
   Const FIELD_SEP = "~"
   Const RECORD_SEP = "|"

   Dim sFlds As String
   Dim sRecs As New StringBuilder
   Dim lRecCnt As Long
   Dim lFld As Long
   Dim fld As Variant
   Dim rows As Variant

   lRecCnt = 0
   If st = "" Then
      StringToJSON = "null"
   Else
      rows = Split(st, RECORD_SEP)
      For lRecCnt = LBound(rows) To UBound(rows)
         sFlds = ""
         fld = Split(rows(lRecCnt), FIELD_SEP)
         For lFld = LBound(fld) To UBound(fld) Step 2
            sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
         Next 'fld
         sRecs.append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
      Next 'rec
      StringToJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
   End If
End Function

'NOTE: as of 2007, I believe ADODB is screwed up.  Disabled for now. - TOM

'Public Function RStoJSON(rs As ADODB.Recordset) As String
'   On Error GoTo errHandler
'   Dim sFlds As String
'   Dim sRecs As New cStringBuilder
'   Dim lRecCnt As Long
'   Dim fld As ADODB.field
'
'   lRecCnt = 0
'   If rs.state = adStateClosed Then
'      RStoJSON = "null"
'   Else
'      If rs.EOF Or rs.BOF Then
'         RStoJSON = "null"
'      Else
'         Do While Not rs.EOF And Not rs.BOF
'            lRecCnt = lRecCnt + 1
'            sFlds = ""
'            For Each fld In rs.fields
'               sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.name & """:""" & toUnicode(fld.Value & "") & """")
'            Next 'fld
'            sRecs.Append IIf((Trim(sRecs.toString) <> ""), "," & vbCrLf, "") & "{" & sFlds & "}"
'            rs.moveNext
'         Loop
'         RStoJSON = ("( {""Records"": [" & vbCrLf & sRecs.toString & vbCrLf & "], " & """RecordCount"":""" & lRecCnt & """ } )")
'      End If
'   End If
'
'   Exit Function
'errHandler:
'
'End Function

'Public Function JsonRpcCall(url As String, methName As String, args(), Optional user As String, Optional pwd As String) As Object
'    Dim r As Object
'    Dim cli As Object
'    Dim pText As String
'    Static reqId As Integer
'
'    reqId = reqId + 1
'
'    Set r = CreateObject("Scripting.Dictionary")
'    r("jsonrpc") = "2.0"
'    r("method") = methName
'    r("params") = args
'    r("id") = reqId
'
'    pText = toString(r)
'
'    Set cli = CreateObject("MSXML2.XMLHTTP.6.0")
'   ' Set cli = New MSXML2.XMLHTTP60
'    If Len(user) > 0 Then   ' If Not IsMissing(user) Then
'        cli.Open "POST", url, False, user, pwd
'    Else
'        cli.Open "POST", url, False
'    End If
'    cli.setRequestHeader "Content-Type", "application/json"
'    cli.Send pText
'
'    If cli.Status <> 200 Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL + cli.Status, , cli.statusText
'    End If
'
'    Set r = parse(cli.responseText)
'    Set cli = Nothing
'
'    If r("id") <> reqId Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response id"
'
'    If r.Exists("error") Or Not r.Exists("result") Then
'        Err.Raise vbObjectError + INVALID_RPC_CALL, , "Json-Rpc Response error: " & r("error")("message")
'    End If
'
'    If Not r.Exists("result") Then Err.Raise vbObjectError + INVALID_RPC_CALL, , "Bad Response, missing result"
'
'    Set JsonRpcCall = r("result")
'End Function




Public Function toUnicode(str As String) As String

   Dim x As Long
   Dim uStr As New StringBuilder
   Dim uChrCode As Integer

   For x = 1 To Len(str)
      uChrCode = Asc(Mid(str, x, 1))
      Select Case uChrCode
         Case 8:   ' backspace
            uStr.append "\b"
         Case 9: ' tab
            uStr.append "\t"
         Case 10:  ' line feed
            uStr.append "\n"
         Case 12:  ' formfeed
            uStr.append "\f"
         Case 13: ' carriage return
            uStr.append "\r"
         Case 34: ' quote
            uStr.append "\"""
         Case 39:  ' apostrophe
            uStr.append "\'"
         Case 92: ' backslash
            uStr.append "\\"
         Case 123, 125:  ' "{" and "}"
            uStr.append ("\u" & right("0000" & Hex(uChrCode), 4))
         Case Is < 32, Is > 127: ' non-ascii characters
            uStr.append ("\u" & right("0000" & Hex(uChrCode), 4))
         Case Else
            uStr.append Chr$(uChrCode)
      End Select
   Next
   toUnicode = uStr.toString
   Exit Function

End Function

Private Sub Class_Initialize()
   psErrors = ""
End Sub

