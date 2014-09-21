Attribute VB_Name = "modWiz01Main"
'modWiz01Main - modWiz01Main.bas
'   Main module for Proving Grounds of the Mad Overlord...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit

Global Const Wiz01ItemMapMax As Integer = 569
Global Const Wiz01RaceMapMax As Integer = 10
Global Const Wiz01ProfessionMapMax As Integer = 13
Global Const Wiz01ConditionMapMax As Integer = 11
Global Const Wiz01SpellMapMax As Integer = 95

Type Wiz01Item
    Equipped As Integer
    Cursed As Integer
    Identified As Integer
    ItemCode As Integer
End Type

Type Wiz01Points
    Current As Integer
    Maximum As Integer
End Type

Type Wiz01Character
    NameLength As Byte                  'Pascal Varying Length String Format...
    Name As String * 15                 '
    PasswordLength As Byte              'Pascal Varying Length String Format...
    Password As String * 15             '
    
    Out As Integer                      '00 00 = No; 01 00 = Yes;
    Race As Integer
    Profession As Integer
    AgeInWeeks As Integer
    Status As Integer
    Alignment As Integer
    Statistics As Long
    Unknown1(1 To 4) As Byte
    GP As Long
    Unknown2(1 To 2) As Byte
    ItemCount As Integer
    ItemList(1 To 8) As Wiz01Item       'List of Items (stowing not an option in Wiz01...)
    EXP As Long
    Unknown3(1 To 2) As Byte
    LVL As Wiz01Points
    HP As Wiz01Points
    SpellBooks(1 To 8) As Byte         'Need to mask as bits...
    MageSpellPoints(1 To 7) As Integer
    PriestSpellPoints(1 To 7) As Integer
    Unknown4(1 To 34) As Byte
End Type
Private Spells(1 To 49) As String
Public Sub DumpWiz01(ByVal strFile As String)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim iChar As Long
    Dim Unit As Integer
    Dim Offset As Long
    Dim errorCode As Long
    Dim xCharacters(1 To 20) As Wiz01Character
    Dim strTemp As String
    
    'Mage Spell Book...
    Spells(1) = "HALITO"
    Spells(2) = "MOGREF"
    Spells(3) = "KATINO"
    Spells(4) = "DUMAPIC"
    Spells(5) = "DILTO"
    Spells(6) = "SOPIC"
    Spells(7) = "MAHALITO"
    Spells(8) = "MOLITO"
    Spells(9) = "MORLIS"
    Spells(10) = "DALTO"
    Spells(11) = "LAHALITO"
    Spells(12) = "MAMORLIS"
    Spells(13) = "MAKANITO"
    Spells(14) = "MADALTO"
    Spells(15) = "LAKANITO"
    Spells(16) = "ZILWAN"
    Spells(17) = "MASOPIC"
    Spells(18) = "HAMAN"
    Spells(19) = "MALOR"
    Spells(20) = "MAHAMAN"
    Spells(21) = "TILTOWAIT"
    'Priest Spell Book...
    Spells(22) = "KALKI"
    Spells(23) = "DIOS"
    Spells(24) = "BADIOS"
    Spells(25) = "MILWA"
    Spells(26) = "PORFIC"
    Spells(27) = "MATU"
    Spells(28) = "CALFO"
    Spells(29) = "MANIFO"
    Spells(30) = "MONTINO"
    Spells(31) = "DIALKO"
    Spells(32) = "LATUMAPIC"
    Spells(33) = "BAMATU"
    Spells(34) = "DIAL"
    Spells(35) = "BADIAL"
    Spells(36) = "LATUMOFIS"
    Spells(37) = "MAPORFIC"
    Spells(38) = "DIALMA"
    Spells(39) = "BADIALMA"
    Spells(40) = "LITOKAN"
    Spells(41) = "KANDI"
    Spells(42) = "DI"
    Spells(43) = "BADI"
    Spells(44) = "LORTO"
    Spells(45) = "MADI"
    Spells(46) = "MABADI"
    Spells(47) = "LOKTOFEIT"
    Spells(48) = "MALIKTO"
    Spells(49) = "KADORTO"
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open strFile For Binary Access Read Write Lock Read Write As #Unit
    Offset = &H1D801
    For i = 1 To 5
        iChar = (i * 4) - 3
        Get #Unit, Offset, xCharacters(iChar)
        Get #Unit, , xCharacters(iChar + 1)
        Get #Unit, , xCharacters(iChar + 2)
        Get #Unit, , xCharacters(iChar + 3)
        Offset = Offset + 1024
    Next i
    Close #Unit
    
    For i = 7 To 7 '1 To 20
        Debug.Print String(80, "=")
        Debug.Print "Character #" & i
        With xCharacters(i)
            If Left(.Name, .NameLength) = vbNullString Then GoTo NextCharacter
            
            .Name = Replace(.Name, Chr(0), " ")
            
            Debug.Print "Name:               " & vbTab & Left(.Name, .NameLength)
            Debug.Print "Password:           " & vbTab & Left(.Password, .PasswordLength)

            If .Out = 1 Then
                Debug.Print "Out (Left in Maze): " & vbTab & "YES"
            Else
                Debug.Print "Out (Left in Maze): " & vbTab & "NO"
            End If

            Debug.Print "Race:               " & vbTab & strRace(.Race)
            Debug.Print "Profession:         " & vbTab & strProfession(.Profession)
            Debug.Print "Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
            Debug.Print "Status:             " & vbTab & strStatus(.Status)
            Debug.Print "Alignment:          " & vbTab & strAlignment(.Alignment)
            Debug.Print "Level:              " & vbTab & strPoints(.LVL)
            Debug.Print "Hit Points:         " & vbTab & strPoints(.HP)
            Debug.Print "Gold Pieces:        " & vbTab & .GP & vbTab & "0x" & Hex(.GP)
            Debug.Print "Experience Points:  " & vbTab & .EXP & vbTab & "0x" & Hex(.EXP)

            Debug.Print vbCrLf & "Basic Statistics..."
            Debug.Print "Strength:          " & vbTab & icvtStatistic(.Statistics, 1)
            Debug.Print "Intellegence:      " & vbTab & icvtStatistic(.Statistics, 2)
            Debug.Print "Piety:             " & vbTab & icvtStatistic(.Statistics, 3)
            Debug.Print "Vitality:          " & vbTab & icvtStatistic(.Statistics, 4)
            Debug.Print "Agility:           " & vbTab & icvtStatistic(.Statistics, 5)
            Debug.Print "Luck:              " & vbTab & icvtStatistic(.Statistics, 6)

            Debug.Print vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
            For j = 1 To .ItemCount
                Debug.Print strItem(.ItemList(j))
            Next j

            Debug.Print vbCrLf & "SpellBooks..."
            For j = 1 To 8
                For k = 1 To 8
                    If ((j - 1) * 8) + k <= UBound(Spells) Then Debug.Print vbTab & strSpell(((j - 1) * 8) + k, .SpellBooks(j), k - 1)
                Next k
            Next j
        
            Debug.Print " "
            strTemp = "Mage Spell Points: "
            For j = 1 To 7
                strTemp = strTemp & .MageSpellPoints(j) & "/"
            Next j
            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
            Debug.Print strTemp
            
            strTemp = "Priest Spell Points: "
            For j = 1 To 7
                strTemp = strTemp & .PriestSpellPoints(j) & "/"
            Next j
            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
            Debug.Print strTemp
            
            Debug.Print " "
            Debug.Print "Unknown Region #1 (4 bytes): "
            Debug.Print strHex(.Unknown1, 4) & vbCrLf

            Debug.Print "Unknown Region #2 (2 bytes): "
            Debug.Print strHex(.Unknown2, 2) & vbCrLf

            Debug.Print "Unknown Region #3 (2 bytes): "
            Debug.Print strHex(.Unknown3, 2) & vbCrLf

            Debug.Print "Unknown Region #4 (34 bytes): "
            Debug.Print strHex(.Unknown4, 34) & vbCrLf
        End With

NextCharacter:
    Next i
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "DumpWiz01"
    Exit Sub
    Resume Next
End Sub
Public Sub PopulateCondition(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01ConditionMapMax
            .AddItem strStatus(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateProfession(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01ProfessionMapMax
            .AddItem strProfession(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateRace(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01RaceMapMax
            .AddItem strRace(i), CInt(i)
        Next i
    End With
End Sub
Public Sub ReadWiz01(ByVal strFile As String, xCharacters() As Wiz01Character)
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    
    'Proving Grounds of the Mad Overlord supports up to 20 characters...
    'The layout is a little funky in that the Character structure seems
    'not to have lined-up evenly with the disk layout (1024 byte blocks
    'on disk)... So, characters start at offset 0x0001D800 and then are
    'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
    '(completing the 1K disk block), then another 4 blocks... for 5
    'total blocks of 4, making 20 characters.
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open strFile For Binary Access Read Write Lock Read Write As #Unit
    Offset = &H1D801
    For i = 1 To 5
        iChar = (i * 4) - 3
        Get #Unit, Offset, xCharacters(iChar)
        Get #Unit, , xCharacters(iChar + 1)
        Get #Unit, , xCharacters(iChar + 2)
        Get #Unit, , xCharacters(iChar + 3)
        Offset = Offset + 1024
    Next i
    Close #Unit
    
    For i = 1 To 20
        xCharacters(i).Name = Replace(xCharacters(i).Name, Chr(0), " ")
    Next i
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "ReadWiz01"
    Exit Sub
    Resume Next
End Sub
Private Function strAlignment(ByVal x As Integer) As String
    Select Case x
        Case 0
            strAlignment = "Unaligned"
        Case 1
            strAlignment = "Good"
        Case 2
            strAlignment = "Neutral"
        Case 3
            strAlignment = "Evil"
        Case Else
            strAlignment = "Unknown"
    End Select
End Function
Private Function strStatus(ByVal x As Integer) As String
    Select Case x
        Case 0
            strStatus = "OK"
        Case 1
            strStatus = "Afraid"
        Case 2
            strStatus = "Asleep"
        Case 3
            strStatus = "Paralyzed"
        Case 4
            strStatus = "Stoned"
        Case 5
            strStatus = "Dead"
        Case 6
            strStatus = "Ashes"
        Case 7
            strStatus = "Lost"
        Case Else
            strStatus = "Unknown"
    End Select
End Function
Private Function strHex(ByRef xBytes() As Byte, nBytes As Integer) As String
    Dim i As Integer

    strHex = ""
    For i = 1 To nBytes
        strHex = strHex & Format(Hex(xBytes(i)), "00") ' & " "
        If i Mod 4 = 0 Then strHex = strHex & " "
        If i Mod 32 = 0 Then strHex = strHex & vbCrLf
    Next i
End Function
Private Function strProfession(ByVal x As Integer) As String
    Select Case x
        Case 0
            strProfession = "Fighter"
        Case 1
            strProfession = "Mage"
        Case 2
            strProfession = "Priest"
        Case 3
            strProfession = "Thief"
        Case 4
            strProfession = "Bishop"
        Case 5
            strProfession = "Samurai"
        Case 6
            strProfession = "Lord"
        Case 7
            strProfession = "Ninja"
        Case Else
            strProfession = "Unknown"
    End Select
End Function
Private Function strRace(ByVal x As Byte) As String
    Select Case x
        Case 1
            strRace = "Human"
        Case 2
            strRace = "Elf"
        Case 3
            strRace = "Dwarf"
        Case 4
            strRace = "Gnome"
        Case 5
            strRace = "Hobbit"
        Case Else
            strRace = "Unknown"
    End Select
End Function
Private Function strItem(x As Wiz01Item) As String
    strItem = vbTab & "Code: " & x.ItemCode
End Function
Private Function strPoints(x As Wiz01Points) As String
    strPoints = x.Current & "/" & x.Maximum
End Function
Private Function strSpell(Spell As Integer, Data As Byte, Offset As Integer) As String
    Dim Temp As String
    If (Data And 2 ^ Offset) = 2 ^ Offset Then Temp = "[X]" Else Temp = "[ ]"
    strSpell = Temp & " " & Spells(Spell) '& vbTab & "[Spell: " & Spell & "; Data: " & Hex(Data) & "; Offset: " & Offset & "]"
End Function
Private Function icvtStatistic(xStatistics As Long, WhichStat As Integer) As Integer
    Select Case WhichStat
        Case 1  'Strength
            icvtStatistic = ((xStatistics \ (2 ^ 0)) And &H1F)
        Case 2  'Intelligence
            icvtStatistic = ((xStatistics \ (2 ^ 5)) And &H1F)
        Case 3  'Piety
            icvtStatistic = ((xStatistics \ (2 ^ 10)) And &H1F)
        Case 4  'Vitality
            icvtStatistic = ((xStatistics \ (2 ^ 16)) And &H1F)
        Case 5  'Agility
            icvtStatistic = ((xStatistics \ (2 ^ 21)) And &H1F)
        Case 6  'Luck
            icvtStatistic = ((xStatistics \ (2 ^ 26)) And &H1F)
        Case Else
            icvtStatistic = 0
    End Select
End Function
Private Sub Test()
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim Data As Long
    Dim errorCode As Long
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open "Test.dat" For Binary Access Read Write Lock Read Write As #Unit
    For i = 1 To 3
        Get #Unit, , Data
        Debug.Print "Data: " & Hex(Data)
        Debug.Print vbTab & "STR: " & ((Data \ (2 ^ 0)) And &H1F)
        Debug.Print vbTab & "INT: " & ((Data \ (2 ^ 5)) And &H1F)
        Debug.Print vbTab & "PIE: " & ((Data \ (2 ^ 10)) And &H1F)
        Debug.Print vbTab & "VIT: " & ((Data \ (2 ^ 16)) And &H1F)
        Debug.Print vbTab & "AGL: " & ((Data \ (2 ^ 21)) And &H1F)
        Debug.Print vbTab & "LUC: " & ((Data \ (2 ^ 26)) And &H1F)
    Next i
    Close #Unit

ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Test"
    Exit Sub
    Resume Next
End Sub

