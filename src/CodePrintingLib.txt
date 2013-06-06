'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'Utilities for printing code modules in nicely formatted html.
Public Type CodeCount
    LOComments As Long
    LOCode As Long
End Type

Option Explicit
Public Sub ripmar()
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Dim vbp As VBProject
Set vbp = ActiveWorkbook.VBProject

For Each vbcomp In ActiveWorkbook.VBProject.VBComponents
    If InStr(1, vbcomp.name, "Tree") > 0 Then vbp.VBComponents.Remove vbcomp
Next vbcomp
    
        
End Sub

Public Sub parsetest()
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Dim parser As CodePrinter
Dim testcode As String
Dim htmls As Dictionary
Dim html As String
Dim path As String

path = ActiveWorkbook.path & "\html\"

Set htmls = New Dictionary
Set parser = New CodePrinter

With New FileSystemObject
    If Not .FolderExists(path) Then .CreateFolder (path)
    
    For Each vbcomp In ActiveWorkbook.VBProject.VBComponents
        If vbcomp.Type = vbext_ct_ClassModule Then
            html = parser.parsemodule(vbcomp)
            'htmls.add VBComp.name, html
            With .CreateTextFile(path & vbcomp.name & ".html", True)
                .Write html
            End With
            'Debug.Print html
        End If
    Next vbcomp
End With

testcode = "Public Sub parsetest() 'this is a simple line of code" & vbCrLf & _
           "    debug.print hello world!" & vbCrLf & _
           "End Sub"



End Sub
'parse a module into nicely formatted, printable HTML, with syntax highlighting and
'everything!  VBAIDE won't print formatted code, so this gets around all that!
Public Sub parsetarget(modname As String, Optional path As String)
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Dim parser As CodePrinter
Dim html As String

If path = vbNullString Then path = ActiveWorkbook.path & "\html\"

path = getFolder(path)
Set parser = New CodePrinter

With New FileSystemObject
    Set vbcomp = ActiveWorkbook.VBProject.VBComponents(modname)
    html = parser.parsemodule(vbcomp)
    With .CreateTextFile(path & vbcomp.name & ".html", True)
        .Write html
    End With
    Debug.Print "Dumped " & vbcomp.name & " to : " & path & vbcomp.name & ".html"

End With

End Sub

Public Function getCode(name As String) As String
Static module As VBComponent

'For Each VBComp In ActiveWorkbook.VBProject.VBComponents
Set module = ActiveWorkbook.VBProject.VBComponents(name)

'If module.Type = vbext_ct_ClassModule Then
    With module
        If .CodeModule.CountOfLines > 0 Then
            getCode = .CodeModule.lines(1, .CodeModule.CountOfLines)
        Else
            getCode = vbNullString
        End If
    End With
'Else
'    Err.Raise 101, , "Not a class"

'End If

End Function
Public Sub dumpCode(name As String, path As String, Optional ext As String)
Dim ts As TextStream
Dim vbcomp As VBComponent
path = getFolder(path)

If ext = vbNullString Then ext = ".txt"
With New FileSystemObject
    Set vbcomp = ActiveWorkbook.VBProject.VBComponents(name)
    With .CreateTextFile(path & vbcomp.name & ext, True)
        .Write getCode(name)
        .Close
    End With
    Debug.Print "Dumped " & vbcomp.name & " to : " & path & vbcomp.name & ".txt"
End With

End Sub
'Read code from a directory, load it into the
Public Sub readCode(name As String, path As String)

End Sub
'a simple sub to help us dump documentable code!
Public Sub dumpSporkandMarathon()
Dim marclasses As Collection
Dim marmods As Collection
Dim sporkclasses As Collection
Dim sporkmods As Collection
Dim dir As String
Dim path As String
path = getPath() & "Source\"

Dim md As VBComponent
For Each md In ActiveWorkbook.VBProject.VBComponents
    dir = classify(md.name)
    If md.Type = vbext_ct_ClassModule Then
        parsetarget md.name, path & dir & "\Classes\"
    ElseIf md.Type = vbext_ct_StdModule Then
        parsetarget md.name, path & dir & "\Modules\"
    ElseIf md.Type = vbext_ct_MSForm Then
        parsetarget md.name, path & "Forms" & "\"
    End If
Next md
        
    
End Sub
'a simple sub to help us dump documentable code!
Public Sub dumpSporkandMarathonRaw()
Dim marclasses As Collection
Dim marmods As Collection
Dim sporkclasses As Collection
Dim sporkmods As Collection
Dim dir As String
Dim path As String
path = getPath() & "Source\"

Dim md As VBComponent
For Each md In ActiveWorkbook.VBProject.VBComponents
    dir = classify(md.name)
    If md.Type = vbext_ct_ClassModule Then
        dumpCode md.name, path & dir & "\Classes\"
    ElseIf md.Type = vbext_ct_StdModule Then
        dumpCode md.name, path & dir & "\Modules\"
    ElseIf md.Type = vbext_ct_MSForm Then
        dumpCode md.name, path & "Forms" & "\"
    End If
Next md
        
    
End Sub
Private Function classify(nm As String) As String
If InStr(1, nm, "Marathon") > 0 Then
    classify = "Marathon"
ElseIf InStr(1, nm, "TimeStep") > 0 Then
    classify = "Marathon"
ElseIf InStr(1, nm, "lib") > 0 Then
    classify = "SPORK"
ElseIf InStr(1, nm, "Generic") > 0 Then
    classify = "SPORK"
Else
    classify = "SPORK"
End If
End Function


Public Function LOCReview(indict As Dictionary, category As String, Optional errorhandling As Boolean) As Dictionary
Dim itm
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Dim modulecode As CodeCount
Dim count As CodeCount
Dim filter As String
If errorhandling Then filter = "Err.Raise"
      
Set LOCReview = New Dictionary

For Each itm In indict
    Set vbcomp = indict(itm)
    modulecode = countCode(vbcomp.CodeModule.lines(1, vbcomp.CodeModule.CountOfLines), filter)
    LOCReview.add vbcomp.name, newdict("Name", vbcomp.name, _
                                       "CodeLines", modulecode.LOCode, _
                                       "CommentLines", modulecode.LOComments, _
                                       "Category", category)
    count.LOCode = modulecode.LOCode + count.LOCode
    count.LOComments = modulecode.LOComments + count.LOComments
Next itm

LOCReview.add category & "Total", newdict("Name", category & "Total", _
                                       "CodeLines", count.LOCode, _
                                       "CommentLines", count.LOComments, _
                                       "Category", category)

End Function
Public Function getMarathonRelated() As Collection
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Set getMarathonRelated = New Collection

Set getMarathonRelated = New Collection
For Each vbcomp In ActiveWorkbook.VBProject.VBComponents
    Select Case vbcomp.Type
        Case vbext_ct_ClassModule, vbext_ct_StdModule
            If isMarathon(vbcomp) And vbcomp.CodeModule.CountOfLines > 0 Then
                getMarathonRelated.add vbcomp
            End If
    End Select
Next vbcomp
End Function
Public Function getGeneric() As Collection
Dim vbcomp As VBComponent
Dim codemod As CodeModule
Set getGeneric = New Collection

Set getGeneric = New Collection
For Each vbcomp In ActiveWorkbook.VBProject.VBComponents
    Select Case vbcomp.Type
        Case vbext_ct_ClassModule, vbext_ct_StdModule
            If Not (isMarathon(vbcomp)) And vbcomp.CodeModule.CountOfLines > 0 Then
                getGeneric.add vbcomp
            End If
    End Select
Next vbcomp
End Function
Public Function ReviewALL() As Dictionary

Dim itm
Dim rec As Dictionary

Dim marathons As Dictionary
Dim generics As Dictionary

Dim reviewed As Dictionary

Dim comp As VBComponent
Set marathons = New Dictionary
Set generics = New Dictionary

For Each comp In getMarathonRelated()
    marathons.add comp.name, comp
Next comp

For Each comp In getGeneric()
    generics.add comp.name, comp
Next comp

Set reviewed = SetLib.union(LOCReview(marathons, "Marathon"), LOCReview(generics, "Generic"))


Set ReviewALL = reviewed

End Function
Public Function ReviewErrorHandling() As Dictionary

Dim itm
Dim rec As Dictionary

Dim marathons As Dictionary
Dim generics As Dictionary

Dim reviewed As Dictionary

Dim comp As VBComponent
Set marathons = New Dictionary
Set generics = New Dictionary

For Each comp In getMarathonRelated()
    marathons.add comp.name, comp
Next comp

For Each comp In getGeneric()
    generics.add comp.name, comp
Next comp

Set reviewed = SetLib.union(LOCReview(marathons, "Marathon", True), LOCReview(generics, "Generic", True))


Set ReviewErrorHandling = reviewed

End Function
Public Sub CodeReview()
Dim res As Dictionary
Set res = ReviewALL
pprint res("MarathonTotal")
pprint res("GenericTotal")


End Sub
Public Sub errorReview()
Dim errorhandlers As Dictionary
Dim itm
Dim handlercount As Long
Dim classcount As Long
Dim rec As Dictionary
Set errorhandlers = ReviewErrorHandling()
For Each itm In errorhandlers
    Set rec = errorhandlers(itm)
    If rec("CodeLines") > 0 Then
        handlercount = handlercount + 1
    End If
    classcount = classcount + 1
Next itm

Debug.Print "There are " & handlercount & " classes or modules with error handling, out of " & classcount & " classes or modules."
'quickview errorhandlers

End Sub
Public Function isMarathon(cmp As VBComponent) As Boolean
isMarathon = InStr(1, cmp.name, "TimeStep") Or InStr(1, cmp.name, "Marathon")
End Function
Public Function countCode(codeblock As String, Optional filter As String) As CodeCount

Dim lines
Dim ln
Dim trimmed As String

lines = Split(codeblock, vbCrLf)
    With countCode
        .LOCode = 0
        .LOComments = 0
    For Each ln In lines
        trimmed = CStr(ln)
        trimmed = Trim(trimmed)
        If Mid(trimmed, 1, 1) = "'" Then
            .LOComments = .LOComments + 1
        Else
            If filter = vbNullString Then
                .LOCode = .LOCode + 1
            ElseIf InStr(1, trimmed, filter) > 0 Then
                .LOCode = .LOCode + 1
            End If
        End If
    Next ln
End With

End Function
