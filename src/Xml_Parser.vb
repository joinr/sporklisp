'this is a new class designed to wrap typical parsing operations for our datastructures.
'basically, we're implementing our own factory class that can produce data structures from
'xml files.
'currently, we're reading in from Excel.
'we'll replace it with "from xml" as an option.

Private Function indent(str As String) As String
indent = str & "    "
End Function
Private Function newline(str As String) As String
newline = vbCrLf & str
End Function
Public Function tag(tagname As String, str As String) As String
tag = "<" & tagname & ">" & str & "</" & tagname & ">"
End Function
Public Function splitTag(tagname As String, str As String) As String
splitTag = "<" & tagname & ">" & str & "</" & tagname & ">"
End Function
Public Function element(elname As String, str As String) As String
element = newline(tag(elname, str))
End Function
Public Function nelement(elname As String, str As String) As String
nelement = newline(tag(elname, str))
End Function
Public Function makeChild(parent As String, body As String) As String
makeChild = tag(parent, body)
End Function
Public Function xmlheader() As String
xmlheader = "<?xml version=1.0?>"
End Function
Public Function newXml(body As String) As String
newXml = xmlheader & newline(body)
End Function
Function testXML() As String
testXML = xmlheader & _
        (splitTag("EmployeeSales", _
                element("Employee", _
                    element("Empid", "999") & _
                    element("FirstName", "Text") & _
                    element("InvoiceNumber", "999") & _
                    element("InvoiceAmount", "999"))))
End Function
Function getblock(str As String) As String
Dim fst As Long
Dim snd As Long
Dim trd As Long
Dim fth As Long
fst = InStr(1, str, "<")
snd = InStr(fst, str, ">")
trd = InStr(snd, str, "</")
fth = InStr(trd, str, ">")
If fst = 0 Or snd = 0 Or trd = 0 Or fth = 0 Then
    Err.Raise 101, , "poor xml"
Else
    gettag = Mid(str, fst, fth)
End If

End Function

Function parsexml(acc As String, indent As Long, remaining As String)

End Function
Public Function mapToSchema(name As String, Optional map As XmlMap) As String
If Not (map Is Nothing) Then
    mapToSchema = map.Schemas(1).xml
Else
    mapToSchema = ActiveWorkbook.XmlMaps(name).Schemas(1).xml
End If

End Function

Public Sub writeSchema(name As String, Optional filename As String)
If filename <> vbNullString Then filename = name & ".xsd"
Open "C:\Documents and Settings\Tom\My Documents\" & filename For Output As #1
Print #1, mapToSchema(name)

End Sub

Public Function ImportXMLtoList(xmlfilepath As String) As String
'Dim idx As Long
'Application.DisplayAlerts = False
'Set ImportXMLtoList = Workbooks.OpenXML(xmlfilepath, , xlXmlLoadImportToList)
'Application.DisplayAlerts = True
'With ActiveWorkbook.XmlMaps
'    ImportXMLtoList = .item(.count)
'End With
End Function

Public Sub RemoveMap(name As String)
ThisWorkbook.XmlMaps(name).Delete
End Sub

Public Sub Create_XSD(xml As String)
Dim myxml As String, mymap As XmlMap
Dim myschema As String

Application.DisplayAlerts = False
Set mymap = ThisWorkbook.XmlMaps.add(xml)
Application.DisplayAlerts = True

myschema = mapToSchema(vbNullString, mymap)
writeSchema mymap.name

Application.DisplayAlerts = True


End Sub
'recursive function to print an xml string
Public Function prettyprint(xml As String, Optional indent As String)
    
End Function

Public Function data(source As String) As String
data = "<![CDATA[" & source & "]]>"
End Function


