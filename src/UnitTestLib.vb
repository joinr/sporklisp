'------SPORK library, by Tom Spoon 6 June 2013-----------------------
'------Licensed under the Eclipse Public License - v 1.0-------------
'See the License module, or License.txt for more information.

'This is a REALLY basic unit testing module.  It's better than nothing!  Makes it easy to bake assertions into your
'tests.
Option Explicit
'simple test rig for unit tests...
'each pair of values is a testing pair, usually a function call, paired with the expected value.
'I'm already doing this implicitly, just going to bake it in formally.  For every test/expect that
'is not equal (fails), we add the pair to the dictionary.
Public Function makeAssertions(ParamArray NameTestExpect()) As Dictionary
Dim i As Long
Dim testval
Dim expected
Dim count As Long
Dim testname As String
testname = CStr(NameTestExpect(LBound(NameTestExpect, 1)))
Set makeAssertions = newdict("TestName", testname)
For i = LBound(NameTestExpect, 1) + 1 To UBound(NameTestExpect, 1) Step 2
    count = count + 1
    testval = NameTestExpect(i)
    expected = NameTestExpect(i + 1)
    If testval <> expected Then
        makeAssertions.add testname & "_" & count, "TEST FAIL on Test #" & count & ", Expected: '" & expected & "' Not Equal to Actual: '" & testval & "'"
    End If
Next i
If makeAssertions.count = 1 Then
    makeAssertions.add "Passed", True
Else
    makeAssertions.add "Passed", False
End If

End Function
Public Function testsPassed(results As Dictionary) As Boolean
testsPassed = results("Passed")
End Function

Public Sub sampleTest()
pprint makeAssertions("Tom Test!", 2, 2, "This will fail", "sure did", 2 + 2, 5)
End Sub
'Allows us to insert generic asserts anywhere.  Can be used to compare objects, values.  Object assertions are
'based on referential equality.
Public Function assert(ByRef testval As Variant, ByRef expected As Variant, Optional noError As Boolean) As Boolean
Dim msg As String
If (IsObject(testval) And Not IsObject(expected)) Or Not IsObject(testval) And IsObject(expected) Then
    assert = False
    msg = "Assertion failed! Mixing objects and primitives!"
ElseIf IsObject(testval) And IsObject(expected) Then
    assert = testval Is expected
    msg = "Assertion failed! Objects do not refer to the same object!"
Else
    assert = testval = expected
    msg = "Assertion failed! Value: '" & testval & "' does not equal '" & expected & "'"
End If

If assert = False And noError = False Then _
    Err.Raise 101, , msg
    
End Function
