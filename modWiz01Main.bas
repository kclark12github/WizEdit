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
Global Const Wiz01CharactersMax As Integer = 20
Global Const Wiz01ItemListMax As Integer = 8
Global Const Wiz01ItemMapMax As Integer = 100
Global Const Wiz01RaceMapMax As Integer = 5
Global Const Wiz01ProfessionMapMax As Integer = 7
Global Const Wiz01StatusMapMax As Integer = 7
Global Const Wiz01SpellLevelMax As Integer = 7
Global Const Wiz01SpellMapMax As Integer = 50

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
Private ItemList(0 To Wiz01ItemMapMax) As String
Private Spells(0 To Wiz01SpellMapMax) As String
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
    Dim bString As String
    
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
    
    InitializeWiz01ItemList
    InitializeWiz01Spells
    For i = 1 To 20
        Debug.Print String(80, "=")
        Debug.Print "Character #" & i
        With xCharacters(i)
            If Left(.Name, .NameLength) = vbNullString Then GoTo NextCharacter
            
            .Name = Replace(.Name, Chr(0), " ")
            
            Debug.Print "Name:               " & vbTab & Left(.Name, .NameLength)
'            Debug.Print "Password:           " & vbTab & Left(.Password, .PasswordLength)
'
'            If .Out = 1 Then
'                Debug.Print "Out (Left in Maze): " & vbTab & "YES"
'            Else
'                Debug.Print "Out (Left in Maze): " & vbTab & "NO"
'            End If
'
'            Debug.Print "Race:               " & vbTab & strRace(.Race)
'            Debug.Print "Profession:         " & vbTab & strProfession(.Profession)
'            Debug.Print "Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
'            Debug.Print "Status:             " & vbTab & strStatus(.Status)
'            Debug.Print "Alignment:          " & vbTab & strAlignment(.Alignment)
'            Debug.Print "Level:              " & vbTab & strPoints(.LVL)
'            Debug.Print "Hit Points:         " & vbTab & strPoints(.HP)
'            Debug.Print "Gold Pieces:        " & vbTab & .GP
'            Debug.Print "Experience Points:  " & vbTab & .EXP
'
'            Debug.Print vbCrLf & "Basic Statistics..."
'            Debug.Print "Strength:           " & vbTab & icvtStatistic(.Statistics, 1)
'            Debug.Print "Intellegence:       " & vbTab & icvtStatistic(.Statistics, 2)
'            Debug.Print "Piety:              " & vbTab & icvtStatistic(.Statistics, 3)
'            Debug.Print "Vitality:           " & vbTab & icvtStatistic(.Statistics, 4)
'            Debug.Print "Agility:            " & vbTab & icvtStatistic(.Statistics, 5)
'            Debug.Print "Luck:               " & vbTab & icvtStatistic(.Statistics, 6)
'
'            Debug.Print vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
'            For j = 1 To .ItemCount
'                Debug.Print strItem(.ItemList(j))
'            Next j
'
'            Debug.Print vbCrLf & "SpellBooks..."
'            bString = icvtSpellsToBin(.SpellBooks)
'            For j = 1 To Wiz01SpellMapMax
'                If Mid(bString, j + 1, 1) = "1" Then
'                    Debug.Print "[X] " & GetSpell(CInt(j))
'                Else
'                    Debug.Print "[ ] " & GetSpell(CInt(j))
'                End If
'            Next j
'
'            Debug.Print " "
'            strTemp = "Mage Spell Points:    " & vbTab
'            For j = 1 To 7
'                strTemp = strTemp & .MageSpellPoints(j) & "/"
'            Next j
'            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
'            Debug.Print strTemp
'
'            strTemp = "Priest Spell Points:  " & vbTab
'            For j = 1 To 7
'                strTemp = strTemp & .PriestSpellPoints(j) & "/"
'            Next j
'            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
'            Debug.Print strTemp
            
            Debug.Print "Unknown Region #1 (4 bytes): "
            Debug.Print strHex(.Unknown1, 4) '& vbCrLf

            Debug.Print "Unknown Region #2 (2 bytes): "
            Debug.Print strHex(.Unknown2, 2) '& vbCrLf

            Debug.Print "Unknown Region #3 (2 bytes): "
            Debug.Print strHex(.Unknown3, 2) '& vbCrLf

            Debug.Print "Unknown Region #4 (34 bytes): "
            Debug.Print strHex(.Unknown4, 34) '& vbCrLf
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
Public Function GetSpell(i As Integer) As String
    GetSpell = Spells(i)
End Function
Public Function icvtSpell(Spell As Integer, Data As Byte, Offset As Integer) As Boolean
    If (Data And 2 ^ Offset) = 2 ^ Offset Then icvtSpell = True Else icvtSpell = False
End Function
Public Function icvtSpellsToBin(Data() As Byte) As String
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    Dim bString As String
    
    bString = vbNullString
    For i = 1 To UBound(Data)
        For Offset = 0 To 7
            If (Data(i) And 2 ^ Offset) = 2 ^ Offset Then bString = bString & "1" Else bString = bString & "0"
        Next Offset
    Next i
    icvtSpellsToBin = bString
    
ExitSub:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "icvtSpellsToBin"
    Exit Function
    Resume Next
End Function
Public Function icvtStatistic(xStatistics As Long, WhichStat As Integer) As Integer
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
Public Sub InitializeWiz01ItemList()
    ItemList(0) = "Broken Item"
    ItemList(1) = "Long Sword"
    ItemList(2) = "Short Sword"
    ItemList(3) = "Anointed Mace"
    ItemList(4) = "Anointed Flail"
    ItemList(5) = "Staff"
    ItemList(6) = "Dagger"
    ItemList(7) = "Small Shield"
    ItemList(8) = "Large Shield"
    ItemList(9) = "Robes"
    ItemList(10) = "Leather Armor"
    ItemList(11) = "Chain Mail"
    ItemList(12) = "Breast Plate"
    ItemList(13) = "Plate Mail"
    ItemList(14) = "Helm"
    ItemList(15) = "Potion Of Dios"
    ItemList(16) = "Potion Of Latumofis"
    ItemList(17) = "Long Sword +1"
    ItemList(18) = "Short Sword +1"
    ItemList(19) = "Mace +1"
    ItemList(20) = "Staff of Mogref"
    ItemList(21) = "Scroll of Katino"
    ItemList(22) = "Leather +1"
    ItemList(23) = "Chain Mail +1"
    ItemList(24) = "Plate Mail +1"
    ItemList(25) = "Shield +1"
    ItemList(26) = "Breast Plate +1"
    ItemList(27) = "Scroll Of Badios"
    ItemList(28) = "Scroll Of Halito"
    ItemList(29) = "Long Sword –1"
    ItemList(30) = "Short Sword –1"
    ItemList(31) = "Mace –1"
    ItemList(32) = "Staff +2"
    ItemList(33) = "Dragon Slayer"
    ItemList(34) = "Helm +1"
    ItemList(35) = "Leather -1"
    ItemList(36) = "Chain -1"
    ItemList(37) = "Breast Plate -1"
    ItemList(38) = "Shield –1"
    ItemList(39) = "Jeweled Amulet"
    ItemList(40) = "Scroll of Badios"
    ItemList(41) = "Potion of Sopic"
    ItemList(42) = "Long Sword +2"
    ItemList(43) = "Short Sword +2"
    ItemList(44) = "Mace +2"
    ItemList(45) = "Scroll Of Lomilw"
    ItemList(46) = "Scroll Of Dilto"
    ItemList(47) = "Copper Gloves"
    ItemList(48) = "Leather +2"
    ItemList(49) = "Chain +2"
    ItemList(50) = "Plate Mail +2"
    ItemList(51) = "Shield +2"
    ItemList(52) = "Helm +2 (E)"
    ItemList(53) = "Potion Of Dial"
    ItemList(54) = "Ring of Porfic"
    ItemList(55) = "Were Slayer"
    ItemList(56) = "Mage Masher"
    ItemList(57) = "Mace Pro Poison"
    ItemList(58) = "Staff Of Montino"
    ItemList(59) = "Blade Cusinart"
    ItemList(60) = "Amulet Of Manifo"
    ItemList(61) = "Rod Of Flame"
    ItemList(62) = "Chain +2 (E)"
    ItemList(63) = "Plate +2 (N)"
    ItemList(64) = "Shield +3 (E)"
    ItemList(65) = "Amulet Of Makanito"
    ItemList(66) = "Helm of Malor"
    ItemList(67) = "Scroll of Badial"
    ItemList(68) = "Short Sword -2"
    ItemList(69) = "Dagger +2"
    ItemList(70) = "Mace –2"
    ItemList(71) = "Staff –2"
    ItemList(72) = "Dagger Of Speed"
    ItemList(73) = "Cursed Robe"
    ItemList(74) = "Leather -2"
    ItemList(75) = "Chain -2"
    ItemList(76) = "Breastplate -2"
    ItemList(77) = "Shield –2"
    ItemList(78) = "Cursed Helmet"
    ItemList(79) = "Breast Plate +2"
    ItemList(80) = "Silver Gloves"
    ItemList(81) = "Evil +3 Sword"
    ItemList(82) = "+3 Evil Short Sword"
    ItemList(83) = "Thieves Dagger"
    ItemList(84) = "+3 Breast Plate"
    ItemList(85) = "Lord's Garb"
    ItemList(86) = "Muramasa Blade"
    ItemList(87) = "Shiriken"
    ItemList(88) = "Chain Pro Fire"
    ItemList(89) = "+3 Evil Plate"
    ItemList(90) = "+3 Shield"
    ItemList(91) = "Ring of Healing"
    ItemList(92) = "Ring Pro Undead"
    ItemList(93) = "Deadly Ring"
    ItemList(94) = "Werdna's Amulet"
    ItemList(95) = "Statuette/Bear"
    ItemList(96) = "Statuette/Frog"
    ItemList(97) = "Bronze Key"
    ItemList(98) = "Silver Key"
    ItemList(99) = "Gold Key"
    ItemList(100) = "Blue Ribbon"
End Sub
Public Sub InitializeWiz01Spells()
    'Mage Spell Book...
    Spells(0) = "Unknown"
    Spells(1) = "Halito"
    Spells(2) = "Mogref"
    Spells(3) = "Katino"
    Spells(4) = "Dumapic"
    Spells(5) = "Dilto"
    Spells(6) = "Sopic"
    Spells(7) = "Mahalito"
    Spells(8) = "Molito"
    Spells(9) = "Morlis"
    Spells(10) = "Dalto"
    Spells(11) = "Lahalito"
    Spells(12) = "Mamorlis"
    Spells(13) = "Makanito"
    Spells(14) = "Madalto"
    Spells(15) = "Lakanito"
    Spells(16) = "Zilwan"
    Spells(17) = "Masopic"
    Spells(18) = "Haman"
    Spells(19) = "Malor"
    Spells(20) = "Mahaman"
    Spells(21) = "Tiltowait"
    'Priest Spell Book...
    Spells(22) = "Kalki"
    Spells(23) = "Dios"
    Spells(24) = "Badios"
    Spells(25) = "Milwa"
    Spells(26) = "Porfic"
    Spells(27) = "Matu"
    Spells(28) = "Calfo"
    Spells(29) = "Manifo"
    Spells(30) = "Montino"
    Spells(31) = "Lomilwa"
    Spells(32) = "Dialko"
    Spells(33) = "Latumapic"
    Spells(34) = "Bamatu"
    Spells(35) = "Dial"
    Spells(36) = "Badial"
    Spells(37) = "Latumofis"
    Spells(38) = "Maporfic"
    Spells(39) = "Dialma"
    Spells(40) = "Badialma"
    Spells(41) = "Litokan"
    Spells(42) = "Kandi"
    Spells(43) = "Di"
    Spells(44) = "Badi"
    Spells(45) = "Lorto"
    Spells(46) = "Madi"
    Spells(47) = "Mabadi"
    Spells(48) = "Loktofeit"
    Spells(49) = "Malikto"
    Spells(50) = "Kadorto"
End Sub
Public Sub PopulateWiz01Item(x As ComboBox)
    Dim i As Integer
    With x
        .Clear
        For i = 0 To Wiz01ItemMapMax
            .AddItem ItemList(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz01Profession(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01ProfessionMapMax
            .AddItem strProfession(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz01Race(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01RaceMapMax
            .AddItem strRace(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz01SpellBooks(lstMageSpells As ListBox, lstPriestSpells As ListBox)
    Dim i As Integer
    With lstMageSpells
        .Clear
        For i = 1 To 21
            .AddItem Spells(i)
        Next i
    End With

    With lstPriestSpells
        .Clear
        For i = 22 To 50
            .AddItem Spells(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz01Status(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01StatusMapMax
            .AddItem strStatus(i), CInt(i)
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
            strStatus = "Lost/Deleted"
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
'    strItem = vbTab & ItemList(x.ItemCode) & "; Code: " & x.ItemCode & "; Equipped: "
'    If x.Identified Then strItem = strItem & "; Identified"
'    If x.Equipped Then strItem = strItem & "; **EQUIPPED**"
'    If x.Cursed Then strItem = strItem & "; --CURSED--"

    strItem = vbTab
    If x.Cursed Then
        strItem = strItem & "-"
    ElseIf x.Equipped Then
        strItem = strItem & "*"
    Else
        strItem = strItem & " "
    End If
    strItem = strItem & ItemList(x.ItemCode)
End Function
Private Function strPoints(x As Wiz01Points) As String
    strPoints = x.Current & "/" & x.Maximum
End Function
Public Function strSpell(Spell As Integer, Data As Byte, Offset As Integer) As String
    Dim Temp As String
    If (Data And 2 ^ Offset) = 2 ^ Offset Then Temp = "[X]" Else Temp = "[ ]"
    strSpell = Temp & " " & Spells(Spell) '& vbTab & "[Spell: " & Spell & "; Data: " & Hex(Data) & "; Offset: " & Offset & "]"
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
Public Sub TestSpells()
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    'Dim Data As Long
    Dim errorCode As Long
    Dim Data(1 To 8) As Byte
    Dim bString As String
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open "spells.dat" For Binary Access Read Write Lock Read Write As #Unit
    For i = 1 To 8
        Get #Unit, , Data(i)
    Next i
    Close #Unit

    bString = vbNullString
    For i = 1 To 8
        bString = bString & Hex(Data(i))
    Next i
    Debug.Print "Hex: " & bString
    
    bString = vbNullString
    For i = 1 To 8
        For Offset = 0 To 7
            If (Data(i) And 2 ^ Offset) = 2 ^ Offset Then bString = bString & "1" Else bString = bString & "0"
        Next Offset
    Next i
    Debug.Print "Bin: " & bString
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Test"
    Exit Sub
    Resume Next
End Sub

