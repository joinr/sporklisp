'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'We're using JSON and CSV files as a primary data interchange format for Marathon.
Private localparser As JSONParser
Public Enum PickleTypes
    primtiveJSON
    compoundJSON
    Serializable
    unknown
End Enum
Option Explicit

Sub tst()
Dim p As Object
Dim jstring As String
Dim JSON As JSONParser
Set JSON = New JSONParser

jstring = "{ width: '200', " & _
            "frame: false, height: 130," & _
            "bodyStyle: 'background-color: #ffffcc;',buttonAlign:'right'," & _
            " items: [{ xtype: 'form',  url: '/content.asp'}," & _
                     "{ xtype: 'form2',  url: '/content2.asp'}] }"
Set p = JSON.parse(jstring)

'Print the text of a nested property '
Debug.Print p.item("width")

'Print the text of a property within an array '
Debug.Print p.item("items")(1).item("xtype")

End Sub
Public Function jParse(ByRef o As Variant) As String
jParse = getparser().toString(o)
End Function
'wrapper function for type casting
Public Function asSerial(o As Object) As ISerializable


If isSeq(o) Then
    Set asSerial = seqToColl(seq(o, ISeq))
Else
    Set asSerial = o
End If
End Function
'TOM added 7 Nov 2012
Public Function trySerialize(ByRef x As Variant) As String
Dim o As Object
If IsObject(x) Then
    Set o = x
    On Error Resume Next
    trySerialize = asSerial(o).asString
    On Error GoTo 0
    Set o = Nothing
Else
    Err.Raise 101, , "Only expected to be called on objects..."
End If

End Function

Sub serialtst()
Dim p As Dictionary
Dim unpickled As Dictionary
Dim jstring As String

Set p = newdict("A", list(1, 2, 3, 4), _
                "B", list(10, 11, 12, 13), _
                "C", newdict("Tom", "Spoon", "Age", 30), _
                "NumberTwo", 2)

jstring = DictionaryToJSON(p)
'produces this output....awesome
'{"A":[1,2,3,4],"B":[10,11,12,13],"C":{"Tom":"Spoon","Age":30},"NumberTwo":2}'

Set unpickled = JSONtoDictionary(jstring) 'from JSON to dictionary
printDict unpickled
End Sub
'Tom Change 9 Sep 2012
Function JSONtoCollection(jstring As String, Optional jp As JSONParser) As Collection
Set jp = getparser(jp)
Set JSONtoCollection = jp.parse(jstring)
End Function
'Tom Change 9 Sep 2012
Function JSONtoPrimitive(jstring As String, Optional jp As JSONParser) As Variant
Set jp = getparser(jp)
bind JSONtoPrimitive, jp.parse(jstring)
End Function
Function JSONtoDictionary(jstring As String, Optional jp As JSONParser) As Dictionary
Set jp = getparser(jp)
Set JSONtoDictionary = jp.parse(jstring)
End Function
Function DictionaryToJSON(indict As Dictionary, Optional jp As JSONParser) As String
Set jp = getparser(jp)
DictionaryToJSON = jp.toString(indict)
End Function

Private Function getparser(Optional jp As JSONParser) As JSONParser

If jp Is Nothing Then
    If localparser Is Nothing Then Set localparser = New JSONParser
    Set getparser = localparser
Else
    Set getparser = jp
End If

End Function
Public Function pickleType(ByRef s As Variant) As PickleTypes
Dim tn As String
pickleType = PickleTypes.unknown
Select Case vartype(s)
    Case vbVariant, vbArray, vbArray + vbVariant, vbNull, vbDate, vbString, vbBoolean, vbDouble, vbDecimal, _
            vbInteger, vbLong, vbSingle
        pickleType = PickleTypes.primtiveJSON
    Case vbObject
        tn = TypeName(s)
        If tn = "Dictionary" Or tn = "Collection" Or tn = "Variant()" Or tn = "ISeq" Then
            pickleType = PickleTypes.compoundJSON
        ElseIf isSeq(s) Then
            pickleType = PickleTypes.compoundJSON
        ElseIf tn = "ISerializable" Then
            pickleType = PickleTypes.Serializable
        End If
End Select
End Function
'Save an object's serialized form as a string, easy for writing files.
Public Function pickle(ByRef s As Variant) As String
Dim ser As ISerializable
Dim tn As String
Select Case pickleType(s)
    Case PickleTypes.primtiveJSON, PickleTypes.compoundJSON
        With getparser(localparser)
            pickle = .toString(s)
        End With
    Case PickleTypes.Serializable
        Set ser = s
        pickle = ser.asString
    Case Else
        Set ser = s
        pickle = ser.asString
        ' Err.Raise 101, , "Don't know howto pickle : " & TypeName(s)
End Select

End Function
'Load a deserializable object, s, with a description of a pickled object
'from a string.  Returns the unpickled object, s.
Public Function unpickle(s As Object, pickled As String) As Object
Dim ser As ISerializable
Select Case pickleType(s)
    Case PickleTypes.compoundJSON
        With getparser(localparser)
            Set s = .parse(pickled)
        End With
    Case PickleTypes.Serializable
        Set ser = s
        ser.FromString pickled
    Case Else
        Set ser = s
        ser.FromString pickled
        'Err.Raise 101, , "Don't know how to unpickle " & TypeName(s)
End Select

Set unpickle = s
End Function
Public Sub pickleTo(ByRef s As Variant, path As String, Optional ts As TextStream)
saveString pickle(s), path, ts
End Sub
Public Function unpickleFrom(s As Object, path As String, Optional ts As TextStream) As Object
Set unpickleFrom = unpickle(s, readString(path, ts))
End Function

Public Sub saveString(str As String, path As String, Optional ts As TextStream, Optional append As Boolean)
If ts Is Nothing Then
    With New FileSystemObject
        Set ts = .CreateTextFile(IOLib.createFolders(path), Not append)
    End With
End If

ts.Write str
ts.Close

End Sub
Public Function readString(path As String, Optional ts As TextStream) As String
If ts Is Nothing Then
    With New FileSystemObject
        Set ts = .OpenTextFile(path, ForReading, False)
    End With
End If

readString = ts.readall
ts.Close

End Function



