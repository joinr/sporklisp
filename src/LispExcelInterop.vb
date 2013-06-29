'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit

'Allow excel formulas to be s-expressions
Function lispFunc(ByRef expr As String, Optional inval As Variant, Optional outval As Variant) As Variant
On Error GoTo clean
Dim args As Collection
If IsMissing(inval) Then
    Set args = list()
Else
    Set args = list(checkRange(inval))
End If
bind lispFunc, apply(eval(read(expr)), args)
Set args = Nothing
Exit Function
clean:
resetLisp
End Function
'Allow results from s-expressions to be serialized...
Sub printRange(inval As Variant)

End Sub

'provide hashed bindings for lisp forms in excel
Function lispVar(ByRef varname As String, ByRef expr As String) As String
With Lisp.getGlobalLispEnv
    If .exists(varname) Then
            .Remove varname
    End If
    .add varname, eval(read(expr))
End With

lispVar = varname
End Function

'bind a range to a table in the global lisp environment
Function lispTable(inrng As Range) As String

End Function