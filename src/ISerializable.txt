'A simple interface, for use in defining objects that can serialize to and from string-based formats.
'I haven't really used toFile or fromFile yet, but there might be binary serialization formats at
'some point, where this would be useful, or custom serialization dependent on file type.
'For now, 99% of the serialization is handle by the pickle and unpickle functions in the
'Serialization module.  This interface is tied heavily to that module.
Option Explicit
'instantiate a serializable object from a string
Public Sub FromString(code As String)
End Sub
'convert a serializable object to its string representation
Public Function asString() As String
End Function
'write a serialized object to a resource path....usually a file.
Public Sub toFile(path As String)
End Sub
'read a serialized object from a resource path...usually a file.
Public Sub fromFile(path As String)
End Sub
'added this for our complex classes.  everything can be built from a dictionary.
Public Sub fromDictionary(indict As Dictionary)
End Sub