'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit

'Tom change 6 Nov 2012
'Traverses a map in depth first order.
'To effeciently do this (and to make it easy to serialize/deserialize using JSON),
'we split our maps into 2 levels...
'Associated values that are also maps become folders with their own stuff...
'Primitive values (i.e. values that are NOT maps), are serialized as a separate
'dictionary under "entries.json".

'Keyvals that are also maps become folders, contents are the vals.
Public Sub mapToFolders(path As String, data As Dictionary, Optional flat As Boolean)
Dim key
Dim itm
Dim d As Dictionary
Dim entries As Dictionary

Set d = data

If flat = False Then
    Set entries = New Dictionary
    
    For Each key In d
        If Not IsObject(data(key)) Then
            'SerializationLib.pickleTo data(key), path & "\" & CStr(key) & ".json"
            entries.add key, data(key)
        Else
            Select Case TypeName(data(key))
                Case "Dictionary"
                    mapToFolders path & "\" & CStr(key), data(key), flat
                Case Else
                    'SerializationLib.pickleTo data(key), path & "\" & CStr(key) & ".json"
                    entries.add key, data(key)
            End Select
        End If
    Next key
    
    If entries.count > 0 Then SerializationLib.pickleTo entries, path & "\entries.json"
Else 'produce a flattened segment of entries.  Each entry is a file, even if it's only a primitive val.
    For Each key In d
        If Not IsObject(data(key)) Then
            SerializationLib.pickleTo data(key), path & "\" & CStr(key) & ".json"
        Else
            Select Case TypeName(data(key))
                Case "Dictionary"
                    mapToFolders path & "\" & CStr(key), data(key), flat
                Case Else
                    SerializationLib.pickleTo data(key), path & "\" & CStr(key) & ".json"
            End Select
        End If
    Next key
End If

Set entries = Nothing
Set d = Nothing

End Sub

'TOM change 7 Nov 2012 -> converts a directory tree into a nested map, where subfolders are
'nested maps as well.  If flat is false or not supplied, non-dictionary values are flattend into a
'subdictionary called entries.json.  If flat is true, every non-dictionary entry is expanded into an
'individual JSON file.
Public Function folderToMap(rootpath As String, Optional flat As Boolean) As Dictionary
Dim key
Dim itm
Dim d As Dictionary
Dim fl As String
Dim entries As Dictionary
Dim folds As Dictionary

Set entries = getEntries(rootpath, flat)
'first thing we do is snarf the entries in the root directory, if any exist...

Set folds = listFolders(rootpath)
For Each key In folds
    Set d = folderToMap(CStr(folds(key)), flat)
    entries.add CStr(key), d
Next key
Set folds = Nothing

Set folderToMap = entries
Set entries = Nothing
Set d = Nothing

End Function

'TOM change 7 Nov 2012 -> loads the root entries from a folder-backed map.
Public Function getEntries(rootpath As String, Optional flat As Boolean) As Dictionary
Dim fl
Dim filepaths As Dictionary
Dim s As String
Dim ents As StringBuilder
If flat = False Then
    If fileExists(rootpath & "\entries.json") Then
        Set getEntries = SerializationLib.unpickleFrom(New Dictionary, rootpath & "\entries.json")
    Else
        Set getEntries = New Dictionary
    End If
Else

    Set ents = New StringBuilder
    ents.append "{"
    Set filepaths = listFiles(rootpath)
    If filepaths.count > 0 Then
        For Each fl In filepaths
            s = SerializationLib.readString(filepaths(fl)) 'build a keyval from the file/string
            ents.append "'" & replace(CStr(fl), ".json", vbNullString) & "':" & s & ","
        Next fl
        ents.Remove ents.length - 1, 1 'drop the last ","
        ents.append "}"
    
        Set getEntries = SerializationLib.JSONtoDictionary(ents.toString)
        Set ents = Nothing
    Else
        Set getEntries = New Dictionary
    End If
    Set filepaths = Nothing
End If
End Function
Public Function fileExists(path As String) As Boolean
'With New Scripting.FileSystemObject
'    fileExists = .fileExists(path)
'End With
fileExists = dir$(path) <> vbNullString
End Function
'Tom change 7 Nov 2012 -> added...
Public Function listFolders(path As String) As Dictionary
Dim fldr As Folder
Dim f As Folder
Set listFolders = New Dictionary
With New FileSystemObject
    If .FolderExists(path) Then
        Set fldr = .getFolder(path)
        For Each f In fldr.SubFolders
            listFolders.add f.name, f.path
        Next f
    End If
End With

End Function
'Tom change 7 Nov 2012 -> added...
Public Function listFiles(path As String, Optional ext As String) As Dictionary
Dim fldr As Folder
Dim f As File
Set listFiles = New Dictionary
With New FileSystemObject
    If .FolderExists(path) Then
        Set fldr = .getFolder(path)
        If ext <> vbNullString Then
            For Each f In fldr.Files
                If InStr(1, f.name, ext) > 0 Then
                    listFiles.add f.name, f.path
                End If
            Next f
        Else
            For Each f In fldr.Files
                    listFiles.add f.name, f.path
            Next f
        End If
    End If
End With

End Function

'eliminate all files in the paths, paths can include directories.  only files in directories will be
'recursively removed, with the directories left unharmed.
Public Function wipeFiles(ParamArray paths()) As Boolean
Dim path As String
Dim p
Dim dir As Folder
Dim fl As File
Dim fldr As Folder
Dim success As Boolean

With New FileSystemObject
    For Each p In paths
        path = CStr(p)
        If .fileExists(path) Then
            .DeleteFile path
        ElseIf .FolderExists(path) Then
            Set dir = .getFolder(path)
            For Each fl In dir.Files
                fl.Delete
            Next fl
            For Each fldr In dir.SubFolders
                success = wipeFiles(fldr.path)
            Next fldr
        End If
    Next p
End With
        
wipeFiles = success
End Function
'eliminate the entire directory structure, including files.
Public Sub wipeAll(ParamArray paths())
Dim path As String
Dim p
Dim dir As Folder
Dim fl As File
Dim fldr As Folder

With New FileSystemObject
    For Each p In paths
        path = CStr(p)
        If .fileExists(path) Then
            .DeleteFile path
        ElseIf .FolderExists(path) Then
            .DeleteFolder path
        End If
    Next p
End With
        
End Sub

Public Sub writeFiles(kvps As Dictionary)

End Sub
Public Function getPath() As String
getPath = ActiveWorkbook.path & "\"
End Function
Public Function createFolders(path As String) As String
Dim tmp
Dim i
Dim current As String
tmp = Split(path, "\")
current = tmp(0)

With New FileSystemObject
    For i = 1 To UBound(tmp, 1) - 1
        current = current & "\" & CStr(tmp(i))
        If Not .FolderExists(current) Then
            Call .CreateFolder(current)
        End If
    Next i
    'Tom change 6 Nov 2012...
    If Not hasExtension(CStr(tmp(i))) Then
        current = current & "\" & CStr(tmp(i))
        If Not .FolderExists(current) Then
            Call .CreateFolder(current)
        End If
    End If
End With
createFolders = path

End Function
Public Function hasExtension(path As String) As Boolean
hasExtension = InStr(1, path, ".")
End Function
Public Function getFolder(path As String) As String

With New FileSystemObject
    If Not .FolderExists(path) Then getFolder = createFolders(path)
End With

getFolder = path
End Function

Public Function emptyFolder(path As String) As String
With New FileSystemObject
    If Not .FolderExists(path) Then
        createFolders (path)
    Else
        wipeFiles (path)
    End If
End With
emptyFolder = path
End Function

Public Function newFolder(path As String) As String
With New FileSystemObject
    If Not .FolderExists(path) Then
        .CreateFolder path
    Else
        .DeleteFolder path
        .CreateFolder path
    End If
End With
newFolder = path

End Function
