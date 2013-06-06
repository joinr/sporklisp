'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'Some primitive version control stuff.  Meant to facilitate use of spork with git.
'Typical workflow is to spit source files out to git for an initial commit.
'As changes are made, blast out the source again, and let git sort the differences.
'Git commits are - currently - handled manually by responsible adults.
Option Explicit

Public Sub dumpAllCode(Optional path As String, Optional ext As String)
Dim vbc As VBComponent
If path = vbNullString Then path = ActiveWorkbook.path & "\"
For Each vbc In ActiveWorkbook.VBProject.VBComponents
    dumpCode vbc.name, path, ext
Next vbc
End Sub
'tries to refresh code modules with serialized code.
Public Sub readAllCode(Optional path As String, Optional ext As String)
Dim modules As Dictionary
Dim vbc As VBComponent
Dim sourcepath As String

If path = vbNullString Then path = ActiveWorkbook.path & "\"

If ext = vbNullString Then ext = ".txt"
For Each vbc In ActiveWorkbook.VBProject.VBComponents
    sourcepath = path & vbc.name & ext

    
Next vbc

End Sub