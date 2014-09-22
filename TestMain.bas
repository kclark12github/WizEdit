Attribute VB_Name = "TestMain"
Option Explicit
Type TP6
    be As Byte              'Biased Exponent
    v1 As Integer           'Lower 16 bits of mantissa
    v2 As Integer           'Middle 16 bits of mantissa
    v3 As Byte              'Signed upper 7 bits of mantissa
End Type

Const DBL_BIAS As Long = &H3FE
Const REAL_BIAS As Long = &H80
Const TP_REAL_BIAS As Long = (DBL_BIAS - REAL_BIAS)    ' 0x37E

Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Sub ClearTP6(TP6buffer As TP6)
    TP6buffer.be = 0
    TP6buffer.v1 = 0
    TP6buffer.v2 = 0
    TP6buffer.v3 = 0
End Sub
Public Function DtoTP6(x As Double) As TP6
    Dim i As Integer
    Dim i4 As Long
    Dim wp(0 To 3) As Integer
    Dim fNegative As Boolean

    Debug.Print String(80, "=")
    If x = 0# Then
        Call ClearTP6(DtoTP6)
        Exit Function
    End If

    'wp = (unsigned short int *)&x;      // Break down double into words
    For i = 0 To 3
        wp(i) = 0
        Call CopyMemory(ByVal VarPtr(wp(i)), ByVal VarPtr(x) + (i * 2), 2)
        'Debug.Print "wp(" & i & ") = " & Hex(wp(i))
    Next i
    
    '===============================================================
    'typedef struct {
    '    unsigned char be   PAK;     /* biased exponent */
    '    unsigned int  v1   PAK;     /* lower 16 bits of mantissa */
    '    unsigned int  v2   PAK;     /* next  16 bits of mantissa */
    '    unsigned int  v3:7 PAK;     /* upper  7 bits of mantissa */
    '    unsigned int  s :1 PAK;     /* sign bit */
    '} tp_real_t;
    '===============================================================
    
    'r.s  = wp[3] >> 15;                 // High bit set for sign
    fNegative = (wp(3) And 2 ^ 15) = 2 ^ 15
    If fNegative Then DtoTP6.v3 = (DtoTP6.v3 Or &H80)
    wp(3) = (wp(3) And &H7FFF)
    Debug.Print "After Sign Check; wp(3) = " & Hex(wp(3))
    
    '// ------------------------------------------------------------------
    '// Grab biased exponent -- exclude sign and shift out the MSB
    '// mantissa bits.
    '
    'r.be = (unsigned char)(((wp[3] & 0x7FFF) >> 4) - TP_REAL_BIAS);
    DtoTP6.be = ((wp(3) \ 2 ^ 4) - TP_REAL_BIAS)
    Debug.Print "DtoTP6.be = 0x" & Hex(DtoTP6.be)
    
    '// ------------------------------------------------------------------
    '// Now...just assign the mantissa after shifting the bits to conform
    '// with the layout for the TP 6-byte real.
    '
    'r.v3 = ((wp[3] & 0x0F) << 3) | (wp[2] >> 13);
    DtoTP6.v3 = ((wp(3) And &HF) * 2 ^ 3) Or (wp(2) \ 2 ^ 13)
    Debug.Print "DtoTP6.v3 = 0x" & Hex(DtoTP6.v3)
    
    'r.v2 = (wp[2] << 3) | (wp[1] >> 13);
    i4 = (wp(2) * 2 ^ 3) Or (wp(1) \ 2 ^ 13)
    'Debug.Print "i4 = 0x" & Hex(i4)
    Call CopyMemory(ByVal VarPtr(DtoTP6.v2), ByVal VarPtr(i4), 2)
    Debug.Print "DtoTP6.v2 = 0x" & Hex(DtoTP6.v2)
    
    'r.v1 = (wp[1] << 3) | (wp[0] >> 13);
    i4 = (wp(1) * 2 ^ 3) Or (wp(0) \ 2 ^ 13)
    'Debug.Print "i4 = 0x" & Hex(i4)
    Call CopyMemory(ByVal VarPtr(DtoTP6.v1), ByVal VarPtr(i4), 2)
    Debug.Print "DtoTP6.v1 = 0x" & Hex(DtoTP6.v1)
    Exit Function
End Function
Public Function TP6toD(r As TP6) As Double
    'if (r.be == 0)return 0.0;
    If r.be = 0 Then
        TP6toD = 0#
        Exit Function
    End If
    
    'return ((((128 + r.v3) * 65536.0) + r.v2) * 65536.0 + r.v1) *
    '    ldexp((r.s ? -1.0 : 1.0), r.be - (129 + 39));
    If (r.v3 And &H80) = &H80 Then
        TP6toD = ((((128 + r.v3) * 65536#) + r.v2) * 65536# + r.v1) * -10 ^ (r.be - (129 + 39))
    Else
        TP6toD = ((((128 + r.v3) * 65536#) + r.v2) * 65536# + r.v1) * 10 ^ (r.be - (129 + 39))
    End If
End Function
Public Sub Main()
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim r As TP6
    Dim x As Double
    Dim errorCode As Long
'553599992787
'23 A7 0F 27 FF FF
    
    'On Error GoTo ErrorHandler
'    Unit = FreeFile
'    Open "Input.dat" For Binary Access Read Write Lock Read Write As #Unit
'    Get #Unit, , Data
'    Close #Unit
    
    x = CDbl(-553599992787#)
    r = DtoTP6(x)
    x = TP6toD(r)
    Debug.Print "Test: " & x
    'Debug.Print "Converted Data: " & TP6toD(Data)
    'Data = DtoTP6(x)
'    Unit = FreeFile
'    Open "Output.dat" For Binary Access Read Write Lock Read Write As #Unit
'    Put #Unit, , x
'    Close #Unit

ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Test"
    Exit Sub
    Resume Next
End Sub


