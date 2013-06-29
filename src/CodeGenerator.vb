Option Explicit

Private reserved As Dictionary
Private Const vartype As String = "[vartype]"
Private Const setter As String = "[setter]"
Private Const pname As String = "[pname]"
Private Const scope As String = "[scope]"
Private Const subclass As String = "[subclass]"


Private Const letbase As String = "[scope] Property Let [subclass][pname] (rhs as [vartype])" & vbCrLf & _
                                  "     [pname] = rhs" & vbCrLf & _
                                  "End Property"
                                  
Private Const setbase As String = "[scope] Property Set [subclass][pname] (rhs as [vartype])" & vbCrLf & _
                                  "     set [pname] = rhs" & vbCrLf & _
                                  "End Property"
                                  
Private Const getBase As String = "[scope] Property Get [subclass][pname] () as [vartype]" & vbCrLf & _
                                  "     [pname] = [pname]" & vbCrLf & _
                                  "End Property"

Private Const getsetbase As String = "[scope] Property Get [subclass][pname] () as [vartype]" & vbCrLf & _
                                     "     set [pname] = [pname]" & vbCrLf & _
                                     "End Property"
                                     

Private Sub Class_Initialize()
Set reserved = New Dictionary

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
Public Function getClassCode(name As String) As String
Static module As VBComponent

Set module = ActiveWorkbook.VBProject.VBComponents(name) 'For Each VBComp In ActiveWorkbook.VBProject.VBComponents

If module.Type = vbext_ct_ClassModule Then
    With module
        If .CodeModule.CountOfLines > 0 Then
            getClassCode = .CodeModule.lines(1, .CodeModule.CountOfLines)
        Else
            getClassCode = vbNullString
        End If
    End With
Else
    Err.Raise 101, , "Not a class"
End If

End Function

Public Function getPublicVars(lines As String) As String

Dim firstfun As Long
Dim firstsub As Long
Dim firstpfun As Long
Dim firstpsub As Long

firstsub = InStr(1, lines, "Public Sub")
firstfun = InStr(1, lines, "Public Function")
firstpsub = InStr(1, lines, "Private Sub")
firstpfun = InStr(1, lines, "Private Function")

getPublicVars = Mid(lines, 1, min(firstsub, firstfun, firstpsub, firstpfun))

End Function
Private Function min(ParamArray args()) As Long
Dim itm
For Each itm In args
    If CLng(itm) < min Then min = CLng(itm)
Next itm

End Function

'Takes an interface, with Public members, and generates required code to fill in members of said interface,
'assuming a delegated internal class.
Public Function buildProperties(vars As String, Optional notpublic As Boolean, Optional subclassed As String) As String
Dim lines
Dim itm
Dim line
Dim property As String
Dim varname As String, vartype As String

lines = Split(vars, vbCrLf)

buildProperties = vars & vbCrLf
    
    
For Each itm In lines
    line = Split(itm, " ")
    If line(0) = "Public" Then
        varname = CStr(line(1))
        vartype = clean(CStr(line(3)))
        buildProperties = buildProperties & buildProperty(varname, vartype, subclassed)
    End If
Next itm

End Function
Public Function clean(inval As String) As String
clean = LCase(inval)
clean = UCase(Mid(inval, 1, 1)) & Mid(inval, 2, Len(inval))
End Function

Public Function buildProperty(varname As String, vtype As String, Optional subclassed As String) As String

buildProperty = buildLet(varname, vtype, subclassed) & vbCrLf & buildGet(varname, vtype, subclassed)
  
End Function

Public Function buildLet(varname As String, vtype As String, Optional subclassed As String) As String
Static base As String

If reserved.exists(vtype) Then base = letbase Else base = setbase

If subclassed <> vbNullString Then subclassed = "_" & subclassed

buildLet = replace(base, pname, varname)
buildLet = replace(buildLet, vartype, vtype)
buildLet = replace(buildLet, subclass, subclassed)

End Function

Public Function buildGet(varname As String, vtype As String, Optional subclassed As String) As String
Static base As String

If reserved.exists(vtype) Then base = getBase Else base = getsetbase

If subclassed <> vbNullString Then subclassed = "_" & subclassed

buildGet = replace(base, pname, varname)
buildGet = replace(buildGet, vartype, vtype)
buildGet = replace(buildGet, subclass, subclassed)


End Function
'grabs code from an interface class, and generates a templated classfile with public variables and private properties
'pointing the to original classname.
Public Function extendInterface(classname As String, Optional path As String) As String
extendInterface = buildProperties(getClassCode(classname), True, classname)
End Function
