'Private Declare Function GetAsyncKeyState Lib "user32" ( _
'ByVal vKey As Long) As Integer
'
'Function ShiftCtrl() As Boolean
'If GetAsyncKeyState(vbKeyControl) Then
'ShiftCtrl = GetAsyncKeyState(vbKeyShift)
'End If
'
'End Function

'A Wrapped Lisp environment that can read and evaluate expressions.
Option Explicit
Public localenv As Dictionary

Public Function read(s As String) As Variant
bind read, Lisp.read(s)
End Function

Public Function eval(exp As Variant, Optional env As Dictionary) As Variant
Dim res As Collection

If env Is Nothing Then Set env = localenv
bind eval, Lisp.eval(exp, env)

End Function

Private Sub Class_Initialize()
Set localenv = Lisp.getGlobalLispEnv
End Sub
