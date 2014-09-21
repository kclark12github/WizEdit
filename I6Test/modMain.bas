Attribute VB_Name = "modMain"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Sub BugAssert(ByVal fExpression As Boolean, Optional sExpression As String)
#If afDebug Then
    If fExpression Then Exit Sub
    BugMessage "BugAssert failed: " & sExpression
    Stop
#End If
End Sub
Public Sub DtoI6(ByVal x As Double, Data() As Integer)
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    
    r3 = x \ 100000000#
    Data(3) = CInt(r3)
    x = x - (r3 * 100000000#)
    r2 = x \ 10000#
    Data(2) = r2
    x = x - (r2 * 10000#)
    r1 = x
    Data(1) = r1
End Sub
Public Function I6toD(Data() As Integer) As Double
    Dim r1 As Double
    Dim r2 As Double
    Dim r3 As Double
    
    r1 = Data(1)
    r2 = Data(2) * 10000#
    r3 = Data(3) * 100000000#
    
    I6toD = r1 + r2 + r3
End Function
Public Sub BytesFromWord(w As Integer, abBuf() As Byte, iOffset As Long)
    BugAssert iOffset <= UBound(abBuf)
    CopyMemory abBuf(iOffset), w, 2
End Sub
Public Function lo(x As Integer) As Byte
    Dim b(1 To 2) As Byte
    Call BytesFromWord(x, b, 1)
    lo = b(1)
End Function
Public Function hi(x As Integer) As Byte
    Dim b(1 To 2) As Byte
    Call BytesFromWord(x, b, 1)
    hi = b(2)
End Function
Public Function NiManI6toD(num() As Integer) As Double
    NiManI6toD = _
          (lo(num(3)) And 15) * 100000000# + _
          (lo(num(3)) And 240) * 1600000000# + _
          (hi(num(2)) And 15) * 2560000# + _
          (hi(num(2)) And 240) * 40960000# + _
          (lo(num(2)) And 15) * 10000# + _
          (lo(num(2)) And 240) * 160000# + _
          (hi(num(1)) And 15) * 256# + _
          (hi(num(1)) And 240) * 4096# + _
          (lo(num(1)) And 15) * 1# + _
          (lo(num(1)) And 240) * 16#
End Function
Public Sub Main()
    Dim n(1 To 3) As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim x As Double
    Dim SaveX As Double
    
    Call DtoI6(32898#, n)
    Call DtoI6(2146000#, n)
    
    n(1) = &H1770
    n(2) = &HD6
    n(3) = &H0
    
'    'Debug.Print NiManI6toD(n)
    Debug.Print I6toD(n)
'    For i = 0 To 9999 'Step 1000
'        n(3) = i
'        Debug.Print "n(3): " & i
'        For j = 0 To 9999 'Step 1000
'            n(2) = j
'            Debug.Print "n(2): " & j
'            For k = 0 To 9999 'Step 1000
'                n(1) = k
'                x = I6toD(n)
'                If SaveX > 0 And SaveX <> x - 1 Then
'                    Debug.Print x
'                End If
'                SaveX = x
'                DoEvents
'            Next k
'        Next j
'    Next i
End Sub
