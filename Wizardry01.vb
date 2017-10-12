'Wizardry01.cls
'   Main Class for Proving Grounds of the Mad Overlord...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class Wizardry01
    Inherits WizEditBase
    Public Sub New(FileName As String, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(FileName, Caption, Icon, BoxArt, Parent)
        Read()
    End Sub
#Region "Properties"
#Region "Declarations"
    Const ScenarioName As String = "PROVING GROUNDS OF THE MAD OVERLORD!"
    Const ScenarioDataOffset As UInt32 = &H1D400
    Const CharacterDataOffset As UInt32 = &H1D800
    'Const ItemMapMax As Integer = 100
    'Const SpellMapMax As Integer = 50

    'Private mItemList(ItemMapMax) As String
    'Private mSpells(SpellMapMax) As String

#End Region
    Public Overrides ReadOnly Property HonorsList As String()
        Get
            HonorsList = {
                "> Chevron of Trebor"
                }
        End Get
    End Property
#End Region
#Region "Methods"
    Public Sub Read()
        Dim binReader As BinaryReader = Nothing
        Try
            If Not File.Exists(MyBase.Path) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", MyBase.Path))
            binReader = New BinaryReader(File.Open(MyBase.Path, FileMode.Open))
            binReader.BaseStream.Position = ScenarioDataOffset
            Dim myScenarioName As String = binReader.ReadString()
            If myScenarioName <> ScenarioName Then Throw New NotSupportedException("Save game file specified is not a valid Ultimate Wizardry Archives: Proving Grounds of the Mad Overlord! save game file.")

            'Proving Grounds of the Mad Overlord supports up to 20 characters...
            'The layout is a little funky in that the Character structure seems
            'not to have lined-up evenly with the disk layout (1024 byte blocks
            'on disk)... So, characters start at offset 0x0001D800 and then are
            'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
            '(completing the 1K disk block), then another 4 blocks... for 5
            'total blocks of 4, making 20 characters.
            binReader.BaseStream.Position = CharacterDataOffset
            For iBlock As Short = 1 To 5
                Dim iChar As Short = (iBlock * 4) - 4
                Characters(iChar).Read(binReader)
                Characters(iChar + 1).Read(binReader)
                Characters(iChar + 2).Read(binReader)
                Characters(iChar + 3).Read(binReader)
                binReader.BaseStream.Position += 192
            Next iBlock
            For iChar As Short = 0 To 19
                If Characters(iChar).Name <> "" Then
                    Debug.WriteLine(New String("="c, 132))
                    Debug.WriteLine(String.Format("{0:00}) {1}", iChar + 1, Characters(iChar).Name))
                    Debug.WriteLine(New String("-"c, 132))
                    Debug.WriteLine(Characters(iChar).ToString)
                End If
            Next iChar
            MyBase.SaveRegSetting("Environment", "UWAPath01", MyBase.DirectoryName)
            MyBase.SaveRegSetting("Environment", "Wiz01DataFile", MyBase.FileName)
        Finally
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
    Public Sub Save()
        Dim binWriter As BinaryWriter = Nothing
        Try
            Backup()
            binWriter = New BinaryWriter(File.Open(MyBase.Path, FileMode.Open, FileAccess.Write, FileShare.None))
            binWriter.BaseStream.Position = CharacterDataOffset
            For iBlock As Short = 1 To 5
                Dim iChar As Short = (iBlock * 4) - 4
                Characters(iChar).Save(binWriter)
                Characters(iChar + 1).Save(binWriter)
                Characters(iChar + 2).Save(binWriter)
                Characters(iChar + 3).Save(binWriter)
                binWriter.BaseStream.Position += 192
            Next iBlock
        Finally
            If binWriter IsNot Nothing Then binWriter.Close() : binWriter = Nothing
        End Try
    End Sub
#End Region
    '    Public Function lkupItemByCbo(x As Integer, cbo As ComboBox) As Integer
    '        For i As Integer = 0 To cbo.Items.Count - 1
    '            If ItemList(x) = cbo.Items(i) Then
    '                Wiz01lkupItemByCbo = i
    '                Exit Function
    '            End If
    '        Next i
    '    End Function
    '    Public Function lkupItemByName(x As String) As Integer
    '        For i As Integer = Wiz01ItemMapMin To Wiz01ItemMapMax
    '            If ItemList(i) = x Then
    '                Wiz01lkupItemByName = i
    '                Exit Function
    '            End If
    '        Next i
    '    End Function
    '    Public Sub PopulateAlignment(x As ComboBox)
    '        With x
    '            .Items.Clear()
    '            For i As Byte = 0 To Wiz01AlignmentMapMax
    '                .Items.Add(Alignment(i))
    '            Next i
    '        End With
    '    End Sub
    '    Public Sub PopulateHonors(x As ListBox)
    '        With x
    '            .Items.Clear()
    '            .Items.Add("> - Chevron of Trebor")
    '            '.Items.Add("G - Mark of Gnilda")
    '            '.Items.Add("K - Knight of Gnilda")
    '            '.Items.Add("D - Descendant of Heroes")
    '            '.Items.Add("* - Star of Llylgamyn")
    '        End With
    '    End Sub
    '    Public Sub PopulateItem(x As ComboBox)
    '        With x
    '            .Items.Clear()
    '            For i As Integer = Wiz01ItemMapMin To Wiz01ItemMapMax
    '                .Items.Add(ItemList(i))    ', i    'Removed to allow ComboBox to be Sorted
    '            Next i
    '        End With
    '    End Sub
    '    Public Sub PopulateProfession(x As ComboBox)
    '        With x
    '            .Items.Clear()
    '            For i As Byte = 0 To Wiz01ProfessionMapMax
    '                .Items.Add(Profession(i))
    '            Next i
    '        End With
    '    End Sub
    '    Public Sub PopulateRace(x As ComboBox)
    '        With x
    '            .Items.Clear()
    '            For i As Byte = 0 To Wiz01RaceMapMax
    '                .Items.Add(Race(i))
    '            Next i
    '        End With
    '    End Sub
    '    Public Sub PopulateSpellBooks(lstMageSpells As ListBox, lstPriestSpells As ListBox)
    '        With lstMageSpells
    '            .Items.Clear()
    '            For i As Integer = 1 To 21
    '                .Items.Add(Spells(i))
    '            Next i
    '        End With
    '        With lstPriestSpells
    '            .Items.Clear()
    '            For i As Integer = 22 To 50
    '                .Items.Add(Spells(i))
    '            Next i
    '        End With
    '    End Sub
    '    Public Sub PopulateStatus(x As ComboBox)
    '        With x
    '            .Items.Clear()
    '            For i As Byte = 0 To Wiz01StatusMapMax
    '                .Items.Add(Status(i))
    '            Next i
    '        End With
    '    End Sub
    Public Sub PrintCharacter(oUnit As Integer, xCharacter As Character)
        '        Dim i As Long
        '        Dim j As Long
        '        Dim k As Long
        '        Dim iChar As Long
        '        Dim errorCode As Long
        '        Dim strTemp As String
        '        Dim bString As String
        '        Dim MageSpell As String * 20
        '    Dim PriestSpell As String * 20

        '    On Error GoTo ErrorHandler

        '        With xCharacter
        '            Print #oUnit, String(100, "=")
        '        Print #oUnit, "Character Print: " & Left(.Name, .NameLength) & " Generated on " & Format(Of Date, "Long Date")()
        '        Print #oUnit, String(100, "-")

        '        Print #oUnit, "Name:               " & vbTab & Left(.Name, .NameLength)
        '        Print #oUnit, "Password:           " & vbTab & Left(.Password, .PasswordLength)
        '        Print #oUnit, "Honors:             " & vbTab & strHonors(.Honors)

        '        If .Out = 1 Then
        '                Print #oUnit, "On Expedition:      " & vbTab & "YES"
        '        Else
        '                Print #oUnit, "On Expedition:      " & vbTab & "NO"
        '        End If
        '            If .Location = 0 Then
        '                Print #oUnit, "Location:           " & vbTab & "Castle"
        '        Else
        '                Print #oUnit, "Location:           " & vbTab & Wiz01Dumapic(xCharacter)
        '        End If

        '            Print #oUnit, "Race:               " & vbTab & strRace(.Race)
        '        Print #oUnit, "Profession:         " & vbTab & strProfession(.Profession)
        '        Print #oUnit, "Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
        '        Print #oUnit, "Status:             " & vbTab & strStatus(.Status)
        '        Print #oUnit, "Alignment:          " & vbTab & strAlignment(.Alignment)
        '        Print #oUnit, "Level:              " & vbTab & strPoints(.LVL)
        '        Print #oUnit, "Hit Points:         " & vbTab & strPoints(.HP)
        '        Print #oUnit, "Gold Pieces:        " & vbTab & I6toD(.GP)
        '        Print #oUnit, "Experience Points:  " & vbTab & I6toD(.EXP)
        '        Print #oUnit, "Armor Class:        " & vbTab & .AC
        '        If .Honors > 0 Then
        '                Select Case .Honors
        '                    Case 1
        '                        Print #oUnit, "Honors:             " & vbTab & """" > """"
        '                Case Else
        '                End Select
        '            End If

        '            Print #oUnit, vbCrLf & "Basic Statistics..."
        '        Print #oUnit, "Strength:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 1)
        '        Print #oUnit, "Intelligence:       " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 2)
        '        Print #oUnit, "Piety:              " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 3)
        '        Print #oUnit, "Vitality:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 4)
        '        Print #oUnit, "Agility:            " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 5)
        '        Print #oUnit, "Luck:               " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 6)

        '        Print #oUnit, vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)..."
        '        For i = 1 To .ItemCount
        '                Print #oUnit, strItem(.ItemList(i))
        '        Next i

        '            Print #oUnit, vbCrLf & "SpellBooks..."
        '        bString = Wiz01cvtSpellsToBstr(.SpellBooks)

        '            Print #oUnit, "Mage:               " & vbTab & "Priest:"
        '        i = i
        '            Do While i <= Wiz01SpellMapMax
        '                If i + 21 > Wiz01SpellMapMax Then Exit Do
        '                MageSpell = String(20, " ")
        '                PriestSpell = String(20, " ")
        '                If i <= 21 Then
        '                    If Mid(bString, i + 1, 1) = "1" Then
        '                        MageSpell = "[X] " & Wiz01GetSpell(CInt(i))
        '                    Else
        '                        MageSpell = "[ ] " & Wiz01GetSpell(CInt(i))
        '                    End If
        '                End If
        '                If Mid(bString, i + 1 + 21, 1) = "1" Then
        '                    PriestSpell = "[X] " & Wiz01GetSpell(CInt(i + 21))
        '                Else
        '                    PriestSpell = "[ ] " & Wiz01GetSpell(CInt(i + 21))
        '                End If
        '                Print #oUnit, MageSpell & vbTab & PriestSpell
        '            i = i + 1
        '            Loop

        '            Print #oUnit, " "
        '        strTemp = "Mage Spell Points:    " & vbTab
        '            For i = 1 To 7
        '                strTemp = strTemp & .MageSpellPoints(i) & "/"
        '            Next i
        '            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        '            Print #oUnit, strTemp

        '        strTemp = "Priest Spell Points:  " & vbTab
        '            For i = 1 To 7
        '                strTemp = strTemp & .PriestSpellPoints(i) & "/"
        '            Next i
        '            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        '            Print #oUnit, strTemp

        '        Print #oUnit, "Unknown Region #1 (4 bytes): "
        '        Print #oUnit, strHex(.Unknown1, 4) '& vbCrLf
        '            Print #oUnit, "Unknown Region #2 (2 bytes): "
        '        Print #oUnit, strHex(.Unknown2, 2) '& vbCrLf
        '            Print #oUnit, "Unknown Region #3 (24 bytes): "
        '        Print #oUnit, strHex(.Unknown3, 24) '& vbCrLf
        '        End With
    End Sub
End Class
