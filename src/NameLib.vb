'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'29 Aug 2012
'Simple library for generating names.
Public namecount As Double
Option Explicit

Public Function randName(Optional base As String) As String
If base = vbNullString Then base = "NAME"
randName = base & Rnd()
End Function

Public Function getName(base As String) As String
getName = base & "_" & namecount
namecount = namecount + 1
End Function
'generate an aggregated name from a set of names
Public Function genName(base As String, nameset As Dictionary) As String
Dim nm
Dim merged As String

merged = base
For Each nm In nameset
    merged = "_" & merged
Next nm

End Function