'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'Integrated 29 Aug 2012
Option Explicit

Public Const inf As Single = 999999999
Public Const infinite As Single = inf
Public Const infLong As Long = 999999999
Public Const infnegative As Single = -999999999


Public Function parseInf(inval) As Single
Select Case CStr(inval)
    Case "inf", "inf+"
        parseInf = LibEnumsAndConstants.inf
    Case "infnegative", "inf-"
        parseInf = LibEnumsAndConstants.infnegative
    Case Else
        parseInf = CSng(inval)
End Select
        

End Function