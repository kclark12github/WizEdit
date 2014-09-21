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

Global Const Wiz01ScenarioName As String = "PROVING GROUNDS OF THE MAD OVERLORD!"
Global Const Wiz01ScenarioDataOffset As Long = &H1D401
Global Const Wiz01CharacterDataOffset As Long = &H1D801
Global Const Wiz01CharactersMax As Integer = 20
Global Const Wiz01ItemListMax As Integer = 8
Global Const Wiz01ItemMapMin As Integer = 0
Global Const Wiz01ItemMapMax As Integer = 100
Global Const Wiz01AlignmentMapMax As Integer = 3
Global Const Wiz01RaceMapMax As Integer = 5
Global Const Wiz01ProfessionMapMax As Integer = 7
Global Const Wiz01StatusMapMax As Integer = 7
Global Const Wiz01SpellLevelMax As Integer = 7
Global Const Wiz01SpellMapMax As Integer = 50

Type Wiz01ScenarioTag
    Length As Byte
    Name As String * 48
End Type

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
    NameLength As Byte                  '0x1D800   Pascal Varying Length String Format...
    Name As String * 15                 '
    PasswordLength As Byte              '0x1D810   Pascal Varying Length String Format...
    Password As String * 15             '
    
    Out As Integer                      '0x1D820    00 00 = No; 01 00 = Yes;
    Race As Integer                     '0x1D822    01 00 = Human
    Profession As Integer               '0x1D824    06 00 = Lord
    AgeInWeeks As Integer               '0x1D826    0C 03 = Weeks Alive...
    Status As Integer                   '0x1D828    00 00 = OK
    Alignment As Integer                '0x1D82A    02 00 = Neutral
    Statistics As Long                  '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
    Unknown1(1 To 4) As Byte            '0x1D830
    GP(1 To 3) As Integer               '0x1D834
    ItemCount As Integer                '0x1D83A
    ItemList(1 To 8) As Wiz01Item       '0x1D83C    List of Items (stowing not an option in Wiz01...)
    EXP(1 To 3) As Integer              '0x1D87C
    LVL As Wiz01Points                  '0x1D882
    HP As Wiz01Points                   '0x1D886
    SpellBooks(1 To 8) As Byte          '0x1D88A    Need to mask as bits...
    MageSpellPoints(1 To 7) As Integer  '0x1D892
    PriestSpellPoints(1 To 7) As Integer '0x1D8A0
    Unknown2(1 To 2) As Byte            '0x1D8AE
    AC As Integer                       '0x1D8B0
    Unknown3(1 To 24) As Byte           '0x1D8B2
    Location As Integer                 '0x1D8CA    Some sort of packed variable...
    Down As Integer                     '0x1D8CC    Seems to be a simple 2-byte integer...
    Honors As Integer                   '0x1D8CE    Need more testing, but 1 = ">"
                                        '0x1D8D0    Next Character Record...
End Type
Private ItemList(Wiz01ItemMapMin To Wiz01ItemMapMax) As String
Private Spells(0 To Wiz01SpellMapMax) As String
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
Public Sub Wiz01cvtBstrToSpells(bString As String, Data() As Byte)
    Dim i As Long
    Dim iByte As Integer
    Dim iOffset As Integer
    
    'First clear the target area...
    For i = 1 To UBound(Data)
        Data(i) = 0
    Next i
    
    'Now inch through the bit-string setting the appropriate bits in the appropriate bytes...
    For i = 0 To (UBound(Data) * 8) - 1
        iOffset = i Mod 8
        iByte = (i \ 8) + 1
        If Mid(bString, i + 1, 1) = "1" Then
            Data(iByte) = (Data(iByte) Or (2 ^ iOffset))
        End If
    Next i
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01cvtSpellsToBstr"
    Exit Sub
    Resume Next
End Sub
Public Function Wiz01cvtSpellsToBstr(Data() As Byte) As String
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
    Wiz01cvtSpellsToBstr = bString
    
ExitSub:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01cvtSpellsToBstr"
    Exit Function
    Resume Next
End Function
Public Function Wiz01cvtStatisticsToLong(Stat1 As String, Stat2 As String, Stat3 As String, Stat4 As String, Stat5 As String, Stat6 As String) As Long
    Wiz01cvtStatisticsToLong = (CLng(Stat1) * (2 ^ 0)) + (CLng(Stat2) * (2 ^ 5)) + (CLng(Stat3) * (2 ^ 10)) + (CLng(Stat4) * (2 ^ 16)) + (CLng(Stat5) * (2 ^ 21)) + (CLng(Stat6) * (2 ^ 26))
End Function
Public Function Wiz01cvtStatisticToInt(xStatistics As Long, WhichStat As Integer) As Integer
    Select Case WhichStat
        Case 1  'Strength
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 0)) And &H1F)
        Case 2  'Intelligence
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 5)) And &H1F)
        Case 3  'Piety
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 10)) And &H1F)
        Case 4  'Vitality
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 16)) And &H1F)
        Case 5  'Agility
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 21)) And &H1F)
        Case 6  'Luck
            Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 26)) And &H1F)
        Case Else
            Wiz01cvtStatisticToInt = 0
    End Select
End Function
Private Function Wiz01Dumapic(xCharacter As Wiz01Character) As String
    With xCharacter
        ' Always seems to be facing North when Quiting from within the Maze...
        Wiz01Dumapic = "Facing North; " & (.Location \ 100) & " East; " & (.Location Mod 100) & " North; " & .Down & " Down"  ' from the steps leading to the castle"
    End With
End Function
Public Sub Wiz01DumpCharacter(xCharacter As Wiz01Character)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim iChar As Long
    Dim errorCode As Long
    Dim strTemp As String
    Dim bString As String
    
    On Error GoTo ErrorHandler
    Debug.Print String(80, "=")
    With xCharacter
        Debug.Print "Name:               " & vbTab & Left(.Name, .NameLength)
        Debug.Print "Password:           " & vbTab & Left(.Password, .PasswordLength)

        If .Out = 1 Then
            Debug.Print "On Expidition:      " & vbTab & "YES"
        Else
            Debug.Print "On Expidition:      " & vbTab & "NO"
        End If
        Debug.Print Wiz01Dumapic(xCharacter)

        Debug.Print "Race:               " & vbTab & strRace(.Race)
        Debug.Print "Profession:         " & vbTab & strProfession(.Profession)
        Debug.Print "Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
        Debug.Print "Status:             " & vbTab & strStatus(.Status)
        Debug.Print "Alignment:          " & vbTab & strAlignment(.Alignment)
        Debug.Print "Level:              " & vbTab & strPoints(.LVL)
        Debug.Print "Hit Points:         " & vbTab & strPoints(.HP)
        Debug.Print "Gold Pieces:        " & vbTab & I6toD(.GP)
        Debug.Print "Experience Points:  " & vbTab & I6toD(.EXP)
        Debug.Print "Armor Class:        " & vbTab & .AC
        If .Honors > 0 Then
            Select Case .Honors
                Case 1
                    Debug.Print "Honors:             " & vbTab & """" > """"
                Case Else
            End Select
        End If
        
        Debug.Print vbCrLf & "Basic Statistics..."
        Debug.Print "Strength:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 1)
        Debug.Print "Intellegence:       " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 2)
        Debug.Print "Piety:              " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 3)
        Debug.Print "Vitality:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 4)
        Debug.Print "Agility:            " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 5)
        Debug.Print "Luck:               " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 6)

        Debug.Print vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
        For i = 1 To .ItemCount
            Debug.Print strItem(.ItemList(i))
        Next i

        Debug.Print " "
        strTemp = "Mage Spell Points:    " & vbTab
        For i = 1 To 7
            strTemp = strTemp & .MageSpellPoints(i) & "/"
        Next i
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        Debug.Print strTemp

        strTemp = "Priest Spell Points:  " & vbTab
        For i = 1 To 7
            strTemp = strTemp & .PriestSpellPoints(i) & "/"
        Next i
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        Debug.Print strTemp
            
        Debug.Print "Unknown Region #1 (4 bytes): "
        Debug.Print strHex(.Unknown1, 4) '& vbCrLf
        Debug.Print "Unknown Region #2 (2 bytes): "
        Debug.Print strHex(.Unknown2, 2) '& vbCrLf
        Debug.Print "Unknown Region #3 (30 bytes): "
        Debug.Print strHex(.Unknown3, 30) '& vbCrLf
    End With

ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01DumpCharacter"
    Exit Sub
    Resume Next
End Sub
Public Function Wiz01GetSpell(i As Integer) As String
    Wiz01GetSpell = Spells(i)
End Function
Public Sub Wiz01InitializeItemList()
    ItemList(0) = "Special: Broken Item"
    ItemList(1) = "Weapon: Long Sword"
    ItemList(2) = "Weapon: Short Sword"
    ItemList(3) = "Weapon: Anointed Mace"
    ItemList(4) = "Weapon: Anointed Flail"
    ItemList(5) = "Weapon: Staff"
    ItemList(6) = "Weapon: Dagger"
    ItemList(7) = "Shield: Small Shield"
    ItemList(8) = "Shield: Large Shield"
    ItemList(9) = "Armor: Robes"
    ItemList(10) = "Armor: Leather Armor"
    ItemList(11) = "Armor: Chain Mail"
    ItemList(12) = "Armor: Breast Plate"
    ItemList(13) = "Armor: Plate Mail"
    ItemList(14) = "Helm: Helm"
    ItemList(15) = "Magic: Potion Of Dios"
    ItemList(16) = "Magic: Potion Of Latumofis"
    ItemList(17) = "Weapon: Long Sword +1"
    ItemList(18) = "Weapon: Short Sword +1"
    ItemList(19) = "Weapon: Mace +1"
    ItemList(20) = "Weapon: Staff of Mogref"
    ItemList(21) = "Magic: Scroll of Katino"
    ItemList(22) = "Armor: Leather +1"
    ItemList(23) = "Armor: Chain Mail +1"
    ItemList(24) = "Armor: Plate Mail +1"
    ItemList(25) = "Shield: Shield +1"
    ItemList(26) = "Armor: Breast Plate +1"
    ItemList(27) = "Magic: Scroll Of Badios"
    ItemList(28) = "Magic: Scroll Of Halito"
    ItemList(29) = "Weapon: Long Sword -1"
    ItemList(30) = "Weapon: Short Sword -1"
    ItemList(31) = "Weapon: Mace -1"
    ItemList(32) = "Weapon: Staff +2"
    ItemList(33) = "Weapon: Dragon Slayer"
    ItemList(34) = "Helm: Helm +1"
    ItemList(35) = "Armor: Leather -1"
    ItemList(36) = "Armor: Chain -1"
    ItemList(37) = "Armor: Breast Plate -1"
    ItemList(38) = "Shield: Shield -1"
    ItemList(39) = "Magic: Jeweled Amulet"
    ItemList(40) = "Magic: Scroll of Badios"
    ItemList(41) = "Magic: Potion of Sopic"
    ItemList(42) = "Weapon: Long Sword +2"
    ItemList(43) = "Weapon: Short Sword +2"
    ItemList(44) = "Weapon: Mace +2"
    ItemList(45) = "Magic: Scroll Of Lomilw"
    ItemList(46) = "Magic: Scroll Of Dilto"
    ItemList(47) = "Gauntlets: Copper Gloves"
    ItemList(48) = "Armor: Leather +2"
    ItemList(49) = "Armor: Chain +2"
    ItemList(50) = "Armor: Plate Mail +2"
    ItemList(51) = "Shield: Shield +2"
    ItemList(52) = "Helm: Helm +2 (E)"
    ItemList(53) = "Magic: Potion Of Dial"
    ItemList(54) = "Magic: Ring of Porfic"
    ItemList(55) = "Weapon: Were Slayer"
    ItemList(56) = "Weapon: Mage Masher"
    ItemList(57) = "Weapon: Mace Pro Poison"
    ItemList(58) = "Weapon: Staff Of Montino"
    ItemList(59) = "Weapon: Blade Cusinart'"
    ItemList(60) = "Magic: Amulet Of Manifo"
    ItemList(61) = "Weapon: Rod Of Flame"
    ItemList(62) = "Armor: Chain +2 (E)"
    ItemList(63) = "Armor: Plate +2 (N)"
    ItemList(64) = "Shield: Shield +3 (E)"
    ItemList(65) = "Magic: Amulet Of Makanito"
    ItemList(66) = "Helm: Helm of Malor"
    ItemList(67) = "Magic: Scroll of Badial"
    ItemList(68) = "Weapon: Short Sword -2"
    ItemList(69) = "Weapon: Dagger +2"
    ItemList(70) = "Weapon: Mace -2"
    ItemList(71) = "Weapon: Staff -2"
    ItemList(72) = "Weapon: Dagger Of Speed"
    ItemList(73) = "Armor: Cursed Robe"
    ItemList(74) = "Armor: Leather -2"
    ItemList(75) = "Armor: Chain -2"
    ItemList(76) = "Armor: Breastplate -2"
    ItemList(77) = "Shield: Shield -2"
    ItemList(78) = "Helm: Cursed Helmet"
    ItemList(79) = "Armor: Breast Plate +2"
    ItemList(80) = "Gauntlets: Gloves of Silver"
    ItemList(81) = "Weapon: Evil +3 Sword"
    ItemList(82) = "Weapon: +3 Evil Short Sword"
    ItemList(83) = "Weapon: Thieves Dagger"
    ItemList(84) = "Armor: +3 Breast Plate"
    ItemList(85) = "Armor: Lord's Garb"
    ItemList(86) = "Weapon: Muramasa Blade"
    ItemList(87) = "Weapon: Shiriken"
    ItemList(88) = "Armor: Chain Pro Fire"
    ItemList(89) = "Armor: +3 Evil Plate"
    ItemList(90) = "Shield: +3 Shield"
    ItemList(91) = "Magic: Ring of Healing"
    ItemList(92) = "Magic: Ring Pro Undead"
    ItemList(93) = "Magic: Deadly Ring"
    ItemList(94) = "Special: Werdna's Amulet"
    ItemList(95) = "Special: Statuette/Bear"
    ItemList(96) = "Special: Statuette/Frog"
    ItemList(97) = "Special: Bronze Key"
    ItemList(98) = "Special: Silver Key"
    ItemList(99) = "Special: Gold Key"
    ItemList(100) = "Special: Blue Ribbon"
End Sub
Public Sub Wiz01InitializeSpells()
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
Public Function Wiz01IsBishop(x As Integer) As Boolean
    Wiz01IsBishop = (strProfession(x) = "Bishop")
End Function
Public Function Wiz01IsMage(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Bishop", "Mage", "Samurai"
            Wiz01IsMage = True
        Case Else
            Wiz01IsMage = False
    End Select
End Function
Public Function Wiz01IsPriest(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Bishop", "Priest", "Lord"
            Wiz01IsPriest = True
        Case Else
            Wiz01IsPriest = False
    End Select
End Function
Public Function Wiz01IsLord(x As Integer) As Boolean
    Wiz01IsLord = (strProfession(x) = "Lord")
End Function
Public Function Wiz01IsSamurai(x As Integer) As Boolean
    Wiz01IsSamurai = (strProfession(x) = "Samurai")
End Function
Public Function Wiz01IsSpellCaster(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Mage", "Priest", "Bishop", "Lord", "Samurai"
            Wiz01IsSpellCaster = True
        Case Else
            Wiz01IsSpellCaster = False
    End Select
End Function
Public Function Wiz01lkupItemByCbo(x As Integer, cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If ItemList(x) = cbo.List(i) Then
            Wiz01lkupItemByCbo = i
            Exit Function
        End If
    Next i
End Function
Public Function Wiz01lkupItemByName(x As String) As Integer
    Dim i As Integer
    For i = Wiz01ItemMapMin To Wiz01ItemMapMax
        If ItemList(i) = x Then
            Wiz01lkupItemByName = i
            Exit Function
        End If
    Next i
End Function
Public Sub Wiz01PopulateAlignment(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01AlignmentMapMax
            .AddItem strAlignment(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz01PopulateItem(x As ComboBox)
    Dim i As Integer
    With x
        .Clear
        For i = Wiz01ItemMapMin To Wiz01ItemMapMax
            .AddItem ItemList(i)    ', i    'Removed to allow ComboBox to be Sorted
        Next i
    End With
End Sub
Public Sub Wiz01PopulateProfession(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01ProfessionMapMax
            .AddItem strProfession(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz01PopulateRace(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01RaceMapMax
            .AddItem strRace(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz01PopulateSpellBooks(lstMageSpells As ListBox, lstPriestSpells As ListBox)
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
Public Sub Wiz01PopulateStatus(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz01StatusMapMax
            .AddItem strStatus(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz01PrintCharacter(oUnit As Integer, xCharacter As Wiz01Character)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim iChar As Long
    Dim errorCode As Long
    Dim strTemp As String
    Dim bString As String
    Dim MageSpell As String * 20
    Dim PriestSpell As String * 20
    
    On Error GoTo ErrorHandler
    
    With xCharacter
        Print #oUnit, String(100, "=")
        Print #oUnit, "Character Print: " & Left(.Name, .NameLength) & " Generated on " & Format(Date, "Long Date")
        Print #oUnit, String(100, "-")
            
        Print #oUnit, "Name:               " & vbTab & Left(.Name, .NameLength)
        Print #oUnit, "Password:           " & vbTab & Left(.Password, .PasswordLength)

        If .Out = 1 Then
            Print #oUnit, "On Expidition:      " & vbTab & "YES"
        Else
            Print #oUnit, "On Expidition:      " & vbTab & "NO"
        End If
        If .Location = 0 Then
            Print #oUnit, "Location:           " & vbTab & "Castle"
        Else
            Print #oUnit, "Location:           " & vbTab & Wiz01Dumapic(xCharacter)
        End If

        Print #oUnit, "Race:               " & vbTab & strRace(.Race)
        Print #oUnit, "Profession:         " & vbTab & strProfession(.Profession)
        Print #oUnit, "Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
        Print #oUnit, "Status:             " & vbTab & strStatus(.Status)
        Print #oUnit, "Alignment:          " & vbTab & strAlignment(.Alignment)
        Print #oUnit, "Level:              " & vbTab & strPoints(.LVL)
        Print #oUnit, "Hit Points:         " & vbTab & strPoints(.HP)
        Print #oUnit, "Gold Pieces:        " & vbTab & I6toD(.GP)
        Print #oUnit, "Experience Points:  " & vbTab & I6toD(.EXP)
        Print #oUnit, "Armor Class:        " & vbTab & .AC
        If .Honors > 0 Then
            Select Case .Honors
                Case 1
                    Print #oUnit, "Honors:             " & vbTab & """" > """"
                Case Else
            End Select
        End If
        
        Print #oUnit, vbCrLf & "Basic Statistics..."
        Print #oUnit, "Strength:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 1)
        Print #oUnit, "Intellegence:       " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 2)
        Print #oUnit, "Piety:              " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 3)
        Print #oUnit, "Vitality:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 4)
        Print #oUnit, "Agility:            " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 5)
        Print #oUnit, "Luck:               " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 6)

        Print #oUnit, vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
        For i = 1 To .ItemCount
            Print #oUnit, strItem(.ItemList(i))
        Next i

        Print #oUnit, vbCrLf & "SpellBooks..."
        bString = Wiz01cvtSpellsToBstr(.SpellBooks)
        
        Print #oUnit, "Mage:               " & vbTab & "Priest:"
        i = i
        Do While i <= Wiz01SpellMapMax
            If i + 21 > Wiz01SpellMapMax Then Exit Do
            MageSpell = String(20, " ")
            PriestSpell = String(20, " ")
            If i <= 21 Then
                If Mid(bString, i + 1, 1) = "1" Then
                    MageSpell = "[X] " & Wiz01GetSpell(CInt(i))
                Else
                    MageSpell = "[ ] " & Wiz01GetSpell(CInt(i))
                End If
            End If
            If Mid(bString, i + 1 + 21, 1) = "1" Then
                PriestSpell = "[X] " & Wiz01GetSpell(CInt(i + 21))
            Else
                PriestSpell = "[ ] " & Wiz01GetSpell(CInt(i + 21))
            End If
            Print #oUnit, MageSpell & vbTab & PriestSpell
            i = i + 1
        Loop
        
        Print #oUnit, " "
        strTemp = "Mage Spell Points:    " & vbTab
        For i = 1 To 7
            strTemp = strTemp & .MageSpellPoints(i) & "/"
        Next i
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        Print #oUnit, strTemp

        strTemp = "Priest Spell Points:  " & vbTab
        For i = 1 To 7
            strTemp = strTemp & .PriestSpellPoints(i) & "/"
        Next i
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        Print #oUnit, strTemp
            
        Print #oUnit, "Unknown Region #1 (4 bytes): "
        Print #oUnit, strHex(.Unknown1, 4) '& vbCrLf
        Print #oUnit, "Unknown Region #2 (2 bytes): "
        Print #oUnit, strHex(.Unknown2, 2) '& vbCrLf
        Print #oUnit, "Unknown Region #3 (24 bytes): "
        Print #oUnit, strHex(.Unknown3, 24) '& vbCrLf
    End With

ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01PrintCharacter"
    Exit Sub
    Resume Next
End Sub
Public Sub Wiz01Read(ByVal strFile As String, xCharacters() As Wiz01Character)
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    Dim ScenarioName As Wiz01ScenarioTag
    
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
    Offset = Wiz01CharacterDataOffset
    For i = 1 To 5
        iChar = (i * 4) - 3
        Get #Unit, Offset, xCharacters(iChar)
        Get #Unit, , xCharacters(iChar + 1)
        Get #Unit, , xCharacters(iChar + 2)
        Get #Unit, , xCharacters(iChar + 3)
        Offset = Offset + 1024
    Next i
    
    For i = 1 To 20
        xCharacters(i).Name = Replace(xCharacters(i).Name, Chr(0), " ")
        xCharacters(i).Name = Left(xCharacters(i).Name, xCharacters(i).NameLength)
    Next i
    
ExitSub:
    Close #Unit
    Call SaveRegSetting("Environment", "UWAPath01", ParsePath(strFile, DrvDirNoSlash))
    Call SaveRegSetting("Environment", "Wiz01DataFile", ParsePath(strFile, FileNameBaseExt))
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01Read"
    Exit Sub
    Resume Next
End Sub
Public Function Wiz01ValidateScenario(xFileName As String) As Boolean
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    Dim ScenarioName As Wiz01ScenarioTag
    
    Wiz01ValidateScenario = False
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open xFileName For Binary Access Read Lock Read As #Unit
    Get #Unit, Wiz01ScenarioDataOffset, ScenarioName
    If Left(ScenarioName.Name, ScenarioName.Length) <> Wiz01ScenarioName Then
        Call MsgBox("Save game file specified is not a valid Ultimate Wizardry Archives: Proving Grounds of the Mad Overlord! save game file.", vbExclamation, Screen.ActiveForm.Caption)
        GoTo ExitSub
    End If
    Wiz01ValidateScenario = True

ExitSub:
    Close #Unit
    Exit Function
    
ErrorHandler:
    Call MsgBox(Err.Description, vbExclamation, "Wiz01ValidateScenario")
    Exit Function
    Resume Next
End Function
Public Sub Wiz01Write(ByVal strFile As String, xCharacters() As Wiz01Character)
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
    Offset = Wiz01CharacterDataOffset
    For i = 1 To 5
        iChar = (i * 4) - 3
        Put #Unit, Offset, xCharacters(iChar)
        Put #Unit, , xCharacters(iChar + 1)
        Put #Unit, , xCharacters(iChar + 2)
        Put #Unit, , xCharacters(iChar + 3)
        Offset = Offset + 1024
    Next i
    Close #Unit
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz01Write"
    Exit Sub
    Resume Next
End Sub
