
Option Explicit

Public interpreter As LispInterpreter
Public expr As Variant

Private Sub InputBox_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If InputBox.value <> vbNullString Then
    bind expr, interpreter.eval(interpreter.read(InputBox.value))
    PrintBox.value = printstr(expr)
Else
    PrintBox.value = vbNullString
End If
End Sub

Private Sub PrintBox_AfterUpdate()
InputBox.SetFocus
End Sub


Private Sub UserForm_Initialize()
Set interpreter = New LispInterpreter
InputBox.EnterKeyBehavior = True
End Sub


