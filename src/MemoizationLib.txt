'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

Option Explicit
'simple global memoization routines...
'allows us to cache values for certain functions.

'Public memo As Dictionary
'Public Function getMemoKey(ParamArray inargs()) As String
'
'End Function
'
'Public Sub memoize(key As String, val As Variant)
'If memo Is Nothing Then Set memo = New Dictionary
'memo.add key, val
'
'End Sub
'
'Public Function memoized(key As String) As Variant
'
'If memo Is Nothing Then
'    Set memo = New Dictionary
'    memoized = Null
'Else
'    If IsObject(memo(key)) Then
'        Set memoized = memo(key)
'    Else
'        memoized = memo(key)
'    End If
'End If
'
'
'End Function