'Borrowed from VB-JSON at:
'http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
' VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
' Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
' BSD Licensed

'I don't really use this class as of yet, since most of my JSON needs are met via serialization stuff.

Option Explicit

Dim dictVars As New Dictionary
Dim plNestCount As Long


Public Function eval(JSON As JSONParser, sJSON As String) As String
   Dim sb As New StringBuilder
   Dim o As Object
   Dim c As Object
   Dim i As Long
   
   Set o = JSON.parse(sJSON)
   If (JSON.GetParserErrors = "") And Not (o Is Nothing) Then
      For i = 1 To o.count
         Select Case vartype(o.item(i))
         Case vbNull
            sb.append "null"
         Case vbDate
            sb.append CStr(o.item(i))
         Case vbString
            sb.append CStr(o.item(i))
         Case Else
            Set c = o.item(i)
            sb.append ExecCommand(c)
         End Select
      Next
   Else
      MsgBox JSON.GetParserErrors, vbExclamation, "Parser Error"
   End If
   eval = sb.toString
End Function

Public Function ExecCommand(ByRef obj As Variant) As String
   Dim sb As New StringBuilder
   
   If plNestCount > 40 Then
      ExecCommand = "ERROR: Nesting level exceeded."
   Else
      plNestCount = plNestCount + 1
      
      Select Case vartype(obj)
         Case vbNull
            sb.append "null"
         Case vbDate
            sb.append CStr(obj)
         Case vbString
            sb.append CStr(obj)
         Case vbObject
            
            Dim i As Long
            Dim j As Long
            Dim this As Object
            Dim key
            Dim paramKeys
            
            If TypeName(obj) = "Dictionary" Then
               Dim sOut As String
               Dim sRet As String
   
               Dim keys
               keys = obj.keys
               For i = 0 To obj.count - 1
                  sRet = ""
             
                  key = keys(i)
                  If vartype(obj.item(key)) = vbString Then
                     sRet = obj.item(key)
                  Else
                     Set this = obj.item(key)
                  End If
                  
                  ' command implementation
                  Select Case LCase(key)
                  Case "alert":
                     MsgBox ExecCommand(this.item("message")), vbInformation, ExecCommand(this.item("title"))
                     
                  Case "input":
                     sb.append InputBox(ExecCommand(this.item("prompt")), ExecCommand(this.item("title")), ExecCommand(this.item("default")))
                     
                  Case "switch"
                     sOut = ExecCommand(this.item("default"))
                     sRet = LCase(ExecCommand(this.item("case")))
                     For j = 0 To this.item("items").count - 1
                        If LCase(this.item("items").item(j + 1).item("case")) = sRet Then
                           sOut = ExecCommand(this.item("items").item(j + 1).item("return"))
                           Exit For
                        End If
                     Next
                     sb.append sOut
                  
                  Case "set":
                     If dictVars.exists(this.item("name")) Then
                        dictVars.item(this.item("name")) = ExecCommand(this.item("value"))
                     Else
                        dictVars.add this.item("name"), ExecCommand(this.item("value"))
                     End If
                     
                  Case "get":
                     sRet = ExecCommand(dictVars(CStr(this.item("name"))))
                     If sRet = "" Then
                        sRet = ExecCommand(this.item("default"))
                     End If
                     
                     sb.append sRet
                     
                  Case "if"
                     Dim val1 As String
                     Dim val2 As String
                     Dim bRes As Boolean
                     val1 = ExecCommand(this.item("value1"))
                     val2 = ExecCommand(this.item("value2"))
                     
                     bRes = False
                     Select Case LCase(this.item("type"))
                     Case "eq" ' =
                        If LCase(val1) = LCase(val2) Then
                           bRes = True
                        End If
                        
                     Case "gt" ' >
                        If val1 > val2 Then
                           bRes = True
                        End If
                     
                     Case "lt" ' <
                        If val1 < val2 Then
                           bRes = True
                        End If
                     
                     Case "gte" ' >=
                        If val1 >= val2 Then
                           bRes = True
                        End If
                     
                     Case "lte" ' <=
                        If val1 <= val2 Then
                           bRes = True
                        End If
                     
                     End Select
                     
                     If bRes Then
                        sb.append ExecCommand(this.item("true"))
                     Else
                        sb.append ExecCommand(this.item("false"))
                     End If
                     
                  Case "return"
                     sb.append obj.item(key)
                  
                     
                  Case Else
                     If TypeName(this) = "Dictionary" Then
                        paramKeys = this.keys
                        For j = 0 To this.count - 1
                           If j > 0 Then
                              sRet = sRet & ","
                           End If
                           sRet = sRet & CStr(this.item(paramKeys(j)))
                        Next
                     End If
                     
                     
                     sb.append "<%" & UCase(key) & "(" & sRet & ")%>"
                     
                  End Select
               Next i
               
            ElseIf TypeName(obj) = "Collection" Then
   
               Dim value
               For Each value In obj
                  sb.append ExecCommand(value)
               Next value
               
            End If
            Set this = Nothing
   
         Case vbBoolean
            If obj Then sb.append "true" Else sb.append "false"
         
         Case vbVariant, vbArray, vbArray + vbVariant
         
         Case Else
            sb.append replace(obj, ",", ".")
      End Select
      plNestCount = plNestCount - 1
   End If
   
   ExecCommand = sb.toString
   Set sb = Nothing
   
End Function


