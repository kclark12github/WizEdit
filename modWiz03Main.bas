Attribute VB_Name = "modWiz03Main"
'modWiz03Main - modWiz03Main.bas
'   Main module for The Legacy of Llylgamyn...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   09/16/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit

Global Const Wiz03ScenarioName As String = "THE LEGACY OF LLYLGAMYN"
Global Const Wiz03ScenarioDataOffset As Long = &H1DA01
Global Const Wiz03CharacterDataOffset As Long = &H1DE01
Global Const Wiz03CharactersMax As Integer = 20
Global Const Wiz03ItemListMax As Integer = 8
Global Const Wiz03ItemMapMin As Integer = 1000
Global Const Wiz03ItemMapMax As Integer = 1103
Global Const Wiz03AlignmentMapMax As Integer = 3
Global Const Wiz03RaceMapMax As Integer = 5
Global Const Wiz03ProfessionMapMax As Integer = 7
Global Const Wiz03StatusMapMax As Integer = 7
Global Const Wiz03SpellLevelMax As Integer = 7
Global Const Wiz03SpellMapMax As Integer = 50

Type Wiz03ScenarioTag
    Length As Byte
    Name As String * 48
End Type

Type Wiz03Item
    Equipped As Integer
    Cursed As Integer
    Identified As Integer
    ItemCode As Integer
End Type

Type Wiz03Points
    Current As Integer
    Maximum As Integer
End Type

Type Wiz03Character
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
    ItemList(1 To 8) As Wiz03Item       '0x1D83C    List of Items (stowing not an option in Wiz03...)
    EXP(1 To 3) As Integer              '0x1D87C
    LVL As Wiz03Points                  '0x1D882
    HP As Wiz03Points                   '0x1D886
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
Private ItemList(Wiz03ItemMapMin To Wiz03ItemMapMax) As String
Private Spells(0 To Wiz03SpellMapMax) As String
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
Private Function strItem(x As Wiz03Item) As String
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
Private Function strPoints(x As Wiz03Points) As String
    strPoints = x.Current & "/" & x.Maximum
End Function
Public Function strSpell(Spell As Integer, Data As Byte, Offset As Integer) As String
    Dim Temp As String
    If (Data And 2 ^ Offset) = 2 ^ Offset Then Temp = "[X]" Else Temp = "[ ]"
    strSpell = Temp & " " & Spells(Spell) '& vbTab & "[Spell: " & Spell & "; Data: " & Hex(Data) & "; Offset: " & Offset & "]"
End Function
Private Function Wiz03Dumapic(xCharacter As Wiz03Character) As String
    With xCharacter
        ' Always seems to be facing North when Quiting from within the Maze...
        Wiz03Dumapic = "Facing North; " & (.Location \ 100) & " East; " & (.Location Mod 100) & " North; " & .Down & " Down"  ' from the steps leading to the castle"
    End With
End Function
Public Sub Wiz03DumpCharacter(xCharacter As Wiz03Character)
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
        Debug.Print Wiz03Dumapic(xCharacter)

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
        Debug.Print "Strength:           " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 1)
        Debug.Print "Intellegence:       " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 2)
        Debug.Print "Piety:              " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 3)
        Debug.Print "Vitality:           " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 4)
        Debug.Print "Agility:            " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 5)
        Debug.Print "Luck:               " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 6)

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
    MsgBox Err.Description, vbExclamation, "Wiz03DumpCharacter"
    Exit Sub
    Resume Next
End Sub
Public Sub Wiz03cvtBstrToSpells(bString As String, Data() As Byte)
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
    MsgBox Err.Description, vbExclamation, "Wiz03cvtSpellsToBstr"
    Exit Sub
    Resume Next
End Sub
Public Function Wiz03cvtSpellsToBstr(Data() As Byte) As String
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
    Wiz03cvtSpellsToBstr = bString

ExitSub:
    Exit Function

ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz03cvtSpellsToBstr"
    Exit Function
    Resume Next
End Function
Public Function Wiz03cvtStatisticsToLong(Stat1 As String, Stat2 As String, Stat3 As String, Stat4 As String, Stat5 As String, Stat6 As String) As Long
    Wiz03cvtStatisticsToLong = (CLng(Stat1) * (2 ^ 0)) + (CLng(Stat2) * (2 ^ 5)) + (CLng(Stat3) * (2 ^ 10)) + (CLng(Stat4) * (2 ^ 16)) + (CLng(Stat5) * (2 ^ 21)) + (CLng(Stat6) * (2 ^ 26))
End Function
Public Function Wiz03cvtStatisticToInt(xStatistics As Long, WhichStat As Integer) As Integer
    Select Case WhichStat
        Case 1  'Strength
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 0)) And &H1F)
        Case 2  'Intelligence
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 5)) And &H1F)
        Case 3  'Piety
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 10)) And &H1F)
        Case 4  'Vitality
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 16)) And &H1F)
        Case 5  'Agility
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 21)) And &H1F)
        Case 6  'Luck
            Wiz03cvtStatisticToInt = ((xStatistics \ (2 ^ 26)) And &H1F)
        Case Else
            Wiz03cvtStatisticToInt = 0
    End Select
End Function
Public Function Wiz03GetSpell(i As Integer) As String
    Wiz03GetSpell = Spells(i)
End Function
Public Sub Wiz03InitializeItemList()
    ItemList(1000) = "Misc: Broken Item"
    ItemList(1001) = "Special: Orb of Earithin"
    ItemList(1002) = "Special: Neutral Crystal"
    ItemList(1003) = "Special: Crystal of Evil"
    ItemList(1004) = "Special: Crystal of Good"
    ItemList(1005) = "Special: Ship in a Bottle"
    ItemList(1006) = "Weapon: Staff of Earth"
    ItemList(1007) = "Magic: Amulet of Air"
    ItemList(1008) = "Misc: Holy Water"
    ItemList(1009) = "Magic: Rod of Fire"
    ItemList(1010) = "Special: Gold Medallion"
    ItemList(1011) = "Special: Orb of Mhuuzfes"
    ItemList(1012) = "Weapon: Butterfly Knife"
    ItemList(1013) = "Weapon: Short Sword"
    ItemList(1014) = "Weapon: Broad Sword"
    ItemList(1015) = "Weapon: Mace"
    ItemList(1016) = "Weapon: Staff of Earth"
    ItemList(1017) = "Weapon: Hand Axe"
    ItemList(1018) = "Weapon: Battle Axe"
    ItemList(1019) = "Weapon: Dagger"
    ItemList(1020) = "Weapon: Flail"
    ItemList(1021) = "Shield: Round Shield"
    ItemList(1022) = "Shield: Heater Shield"
    ItemList(1023) = "Armor: Mage's Robes"
    ItemList(1024) = "Armor: Cuirass"
    ItemList(1025) = "Armor: Hauberk"
    ItemList(1026) = "Armor: Breast Plate"
    ItemList(1027) = "Armor: Plate Armor"
    ItemList(1028) = "Helm: Sallet"
    ItemList(1029) = "Magic: Potion of Dios"
    ItemList(1030) = "Magic: Latumofis Oil"
    ItemList(1031) = "Weapon: Short Sword +1"
    ItemList(1032) = "Weapon: Broad Sword +1"
    ItemList(1033) = "Weapon: Mace +1"
    ItemList(1034) = "Weapon: Battle Axe +1"
    ItemList(1035) = "Weapon: Nunchuka"
    ItemList(1036) = "Weapon: Dagger +1"
    ItemList(1037) = "Magic: Scroll of Katino"
    ItemList(1038) = "Armor: Cuirass +1"
    ItemList(1039) = "Armor: Hauberk +1"
    ItemList(1040) = "Armor: Breast Plate +1"
    ItemList(1041) = "Armor: Plate Armor +1"
    ItemList(1042) = "Shield: Heater +1"
    ItemList(1043) = "Helm: Bascinet"
    ItemList(1044) = "Gauntlets: Gloves of Iron"
    ItemList(1045) = "Magic: Scroll of Badios"
    ItemList(1046) = "Magic: Potion of Halito"
    ItemList(1047) = "Weapon: Short Sword -1"
    ItemList(1048) = "Weapon: Broadsword -1"
    ItemList(1049) = "Weapon: Mace -1"
    ItemList(1050) = "Weapon: Dagger -1"
    ItemList(1051) = "Weapon: Battle Axe -1"
    ItemList(1052) = "Weapon: Margaux's Flail"
    ItemList(1053) = "Special: Bag of Gems"
    ItemList(1054) = "Weapon: Wizard's Staff"
    ItemList(1055) = "Weapon: Flametongue"
    ItemList(1056) = "Shield: Round Shield -1"
    ItemList(1057) = "Armor: Cuirass -1"
    ItemList(1058) = "Armor: Hauberk -1"
    ItemList(1059) = "Armor: Breast Plate -1"
    ItemList(1060) = "Armor: Plate Armor -1"
    ItemList(1061) = "Helm: Sallet -1"
    ItemList(1062) = "Magic: Potion of Sopic"
    ItemList(1063) = "Special: Gold Ring"
    ItemList(1064) = "Special: Salamander Ring"
    ItemList(1065) = "Special: Serpent's Tooth"
    ItemList(1066) = "Weapon: Short Sword +2"
    ItemList(1067) = "Weapon: Broad Sword +2"
    ItemList(1068) = "Weapon: Battle Axe +2"
    ItemList(1069) = "Weapon: Ivory Blade (G)"
    ItemList(1070) = "Weapon: Ebony Blade (E)"
    ItemList(1071) = "Weapon: Amber Blade (N)"
    ItemList(1072) = "Weapon: Mace +2"
    ItemList(1073) = "Gauntlets: Gloves of Mithril"
    ItemList(1074) = "Magic: Amulet of Dailko"
    ItemList(1075) = "Armor: Cuirass +2"
    ItemList(1076) = "Shield: Heater +2"
    ItemList(1077) = "Armor: Displacer Robes"
    ItemList(1078) = "Armor: Hauberk +2"
    ItemList(1079) = "Armor: Breast Plate +2"
    ItemList(1080) = "Armor: Plate Armor +2"
    ItemList(1081) = "Armor: Armet"
    ItemList(1082) = "Armor: Wargan Robes"
    ItemList(1083) = "Weapon: Giant's Club"
    ItemList(1084) = "Weapon: Blade Cuisinart'"
    ItemList(1085) = "Weapon: Shepherd Crook"
    ItemList(1086) = "Weapon: Unholy Axe"
    ItemList(1087) = "Weapon: Rod of Death"
    ItemList(1088) = "Special: Gem of Exorcism"
    ItemList(1089) = "Special: Bag of Emeralds"
    ItemList(1090) = "Special: Bag of Garnets"
    ItemList(1091) = "Special: Blue Pearl"
    ItemList(1092) = "Special: Ruby Slippers"
    ItemList(1093) = "Weapon: Necrology Rod"
    ItemList(1094) = "Misc: Book of Life"
    ItemList(1095) = "Misc: Book of Death"
    ItemList(1096) = "Special: Dragon's Tooth"
    ItemList(1097) = "Special: Trollkin Ring"
    ItemList(1098) = "Special: Rabbit's Foot"
    ItemList(1099) = "Special: Thief's Pick"
    ItemList(1100) = "Misc: Book of Demons"
    ItemList(1101) = "Weapon: Butterfly Knife"
    ItemList(1102) = "Special: Gold Tiara"
    ItemList(1103) = "Gauntlets: Mantis Gloves"
End Sub
Public Sub Wiz03InitializeSpells()
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
Public Function Wiz03IsBishop(x As Integer) As Boolean
    Wiz03IsBishop = (strProfession(x) = "Bishop")
End Function
Public Function Wiz03IsMage(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Bishop", "Mage", "Samurai"
            Wiz03IsMage = True
        Case Else
            Wiz03IsMage = False
    End Select
End Function
Public Function Wiz03IsPriest(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Bishop", "Priest", "Lord"
            Wiz03IsPriest = True
        Case Else
            Wiz03IsPriest = False
    End Select
End Function
Public Function Wiz03IsLord(x As Integer) As Boolean
    Wiz03IsLord = (strProfession(x) = "Lord")
End Function
Public Function Wiz03IsSamurai(x As Integer) As Boolean
    Wiz03IsSamurai = (strProfession(x) = "Samurai")
End Function
Public Function Wiz03IsSpellCaster(x As Integer) As Boolean
    Select Case strProfession(x)
        Case "Mage", "Priest", "Bishop", "Lord", "Samurai"
            Wiz03IsSpellCaster = True
        Case Else
            Wiz03IsSpellCaster = False
    End Select
End Function
Public Function Wiz03lkupItemByCbo(x As Integer, cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If ItemList(x) = cbo.List(i) Then
            Wiz03lkupItemByCbo = i
            Exit Function
        End If
    Next i
End Function
Public Function Wiz03lkupItemByName(x As String) As Integer
    Dim i As Integer
    For i = Wiz03ItemMapMin To Wiz03ItemMapMax
        If ItemList(i) = x Then
            Wiz03lkupItemByName = i
            Exit Function
        End If
    Next i
End Function
Public Sub Wiz03PopulateAlignment(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz03AlignmentMapMax
            .AddItem strAlignment(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz03PopulateItem(x As ComboBox)
    Dim i As Integer
    With x
        .Clear
        For i = Wiz03ItemMapMin To Wiz03ItemMapMax
            .AddItem ItemList(i)    ', i    'Removed to allow ComboBox to be Sorted
        Next i
    End With
End Sub
Public Sub Wiz03PopulateProfession(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz03ProfessionMapMax
            .AddItem strProfession(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz03PopulateRace(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz03RaceMapMax
            .AddItem strRace(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz03PopulateSpellBooks(lstMageSpells As ListBox, lstPriestSpells As ListBox)
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
Public Sub Wiz03PopulateStatus(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz03StatusMapMax
            .AddItem strStatus(i), CInt(i)
        Next i
    End With
End Sub
Public Sub Wiz03PrintCharacter(oUnit As Integer, xCharacter As Wiz03Character)
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
            Print #oUnit, "Location:           " & vbTab & Wiz03Dumapic(xCharacter)
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
        Print #oUnit, "Strength:           " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 1)
        Print #oUnit, "Intellegence:       " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 2)
        Print #oUnit, "Piety:              " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 3)
        Print #oUnit, "Vitality:           " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 4)
        Print #oUnit, "Agility:            " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 5)
        Print #oUnit, "Luck:               " & vbTab & Wiz03cvtStatisticToInt(.Statistics, 6)

        Print #oUnit, vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
        For i = 1 To .ItemCount
            Print #oUnit, strItem(.ItemList(i))
        Next i

        Print #oUnit, vbCrLf & "SpellBooks..."
        bString = Wiz03cvtSpellsToBstr(.SpellBooks)
        
        Print #oUnit, "Mage:               " & vbTab & "Priest:"
        i = i
        Do While i <= Wiz03SpellMapMax
            If i + 21 > Wiz03SpellMapMax Then Exit Do
            MageSpell = String(20, " ")
            PriestSpell = String(20, " ")
            If i <= 21 Then
                If Mid(bString, i + 1, 1) = "1" Then
                    MageSpell = "[X] " & Wiz03GetSpell(CInt(i))
                Else
                    MageSpell = "[ ] " & Wiz03GetSpell(CInt(i))
                End If
            End If
            If Mid(bString, i + 1 + 21, 1) = "1" Then
                PriestSpell = "[X] " & Wiz03GetSpell(CInt(i + 21))
            Else
                PriestSpell = "[ ] " & Wiz03GetSpell(CInt(i + 21))
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
    MsgBox Err.Description, vbExclamation, "Wiz03PrintCharacter"
    Exit Sub
    Resume Next
End Sub
Public Sub Wiz03Read(ByVal strFile As String, xCharacters() As Wiz03Character)
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    Dim ScenarioName As Wiz03ScenarioTag
    
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
    Offset = Wiz03CharacterDataOffset
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
    Call SaveRegSetting("Environment", "UWAPath03", ParsePath(strFile, DrvDirNoSlash))
    Call SaveRegSetting("Environment", "Wiz03DataFile", ParsePath(strFile, FileNameBaseExt))
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "Wiz03Read"
    Exit Sub
    Resume Next
End Sub
Public Function Wiz03ValidateScenario(xFileName As String) As Boolean
    Dim i As Long
    Dim iChar As Integer
    Dim Offset As Long
    Dim Unit As Integer
    Dim errorCode As Long
    Dim ScenarioName As Wiz03ScenarioTag
    
    Wiz03ValidateScenario = False
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open xFileName For Binary Access Read Lock Read As #Unit
    Get #Unit, Wiz03ScenarioDataOffset, ScenarioName
    If Left(ScenarioName.Name, ScenarioName.Length) <> Wiz03ScenarioName Then
        Call MsgBox("Save game file specified is not a valid Ultimate Wizardry Archives: Proving Grounds of the Mad Overlord! save game file.", vbExclamation, Screen.ActiveForm.Caption)
        GoTo ExitSub
    End If
    Wiz03ValidateScenario = True

ExitSub:
    Close #Unit
    Exit Function
    
ErrorHandler:
    Call MsgBox(Err.Description, vbExclamation, "Wiz03ValidateScenario")
    Exit Function
    Resume Next
End Function
Public Sub Wiz03Write(ByVal strFile As String, xCharacters() As Wiz03Character)
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
    Offset = Wiz03CharacterDataOffset
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
    MsgBox Err.Description, vbExclamation, "Wiz03Write"
    Exit Sub
    Resume Next
End Sub
