'clsWiz01.cls
'   Main Class for Proving Grounds of the Mad Overlord...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class clsWiz01
    Public Sub New(FileName As String)
        InitializeItemList()
        InitializeSpells()
        mFileName = FileName
        For iChar As Short = 0 To CharactersMax - 1
            mCharacters(iChar) = New Character()
        Next iChar
        Read()
    End Sub
#Region "#Internal Class(es)"
    Public Class Character
        Public Sub New()
            mName = ""
            mPassword = ""

            mOut = 0S
            mRace = 0S
            mProfession = 0S
            mAgeInWeeks = 0S
            mStatus = 0S
            mAlignment = 0S
            mStatistics = 0
            mUnknown1 = {0, 0, 0, 0}
            mGP = {0S, 0S, 0S}
            mItemCount = 0S
            mItemList = {New Item(), New Item(), New Item(), New Item(), New Item(), New Item(), New Item(), New Item()}
            mEXP = {0S, 0S, 0S}
            mLVL = New Points
            mHP = New Points
            mSpellBooks = {0, 0, 0, 0, 0, 0, 0, 0}
            mMageSpellPoints = {0S, 0S, 0S, 0S, 0S, 0S, 0S}
            mPriestSpellPoints = {0S, 0S, 0S, 0S, 0S, 0S, 0S}
            mUnknown2 = {0, 0}
            mAC = 0S
            mUnknown3 = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
            mLocation = 0S
            mDown = 0S
            mHonors = 0S
        End Sub
#Region "Properties"
#Region "Declarations"
        Const NamePasswordLengthMax As Short = 15

        Const ItemListMax As Integer = 8
        Const AlignmentMapMax As Integer = 3
        Const RaceMapMax As Integer = 5
        Const ProfessionMapMax As Integer = 7
        Const StatusMapMax As Integer = 7
        Const SpellLevelMax As Integer = 7

        Private mName As String                                 '0x1D800   Pascal Varying Length String Format...
        Private mPassword As String                             '0x1D810   Pascal Varying Length String Format...

        Private mOut As Int16                                   '0x1D820    00 00 = No; 01 00 = Yes;
        Private mRace As Int16                                  '0x1D822    01 00 = Human
        Private mProfession As Int16                            '0x1D824    06 00 = Lord
        Private mAgeInWeeks As Int16                            '0x1D826    0C 03 = Weeks Alive...
        Private mStatus As Int16                                '0x1D828    00 00 = OK
        Private mAlignment As Int16                             '0x1D82A    02 00 = Neutral
        Private mStatistics As Int32                            '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
        Private mUnknown1(3) As Byte                            '0x1D830
        Private mGP(2) As Int16                                 '0x1D834
        Private mItemCount As Int16                             '0x1D83A
        Private mItemList(ItemListMax - 1) As Item              '0x1D83C    List of Items (stowing not an option in Wiz01...)
        Private mEXP(2) As Int16                                '0x1D87C
        Private mLVL As Points                                  '0x1D882
        Private mHP As Points                                   '0x1D886
        Private mSpellBooks(7) As Byte                          '0x1D88A    Need to mask as bits...
        Private mMageSpellPoints(SpellLevelMax - 1) As Int16    '0x1D892
        Private mPriestSpellPoints(SpellLevelMax - 1) As Int16  '0x1D8A0
        Private mUnknown2(1) As Byte                            '0x1D8AE
        Private mAC As Int16                                    '0x1D8B0
        Private mUnknown3(23) As Byte                           '0x1D8B2
        Private mLocation As Int16                              '0x1D8CA    Some sort of packed variable...
        Private mDown As Int16                                  '0x1D8CC    Seems to be a simple 2-byte Int16...
        Private mHonors As Int16                                '0x1D8CE    Need more testing, but 1 = ">"
        '                                                       '0x1D8D0    Next Character Record...
#End Region
        Public Property Name As String
            Get
                Return mName
            End Get
            Set(value As String)
                If value.Length > NamePasswordLengthMax Then Throw New ArgumentException(String.Format("Name length is limited to {0} characters!", NamePasswordLengthMax))
                mName = value
            End Set
        End Property
#End Region
#Region "Methods"
        Friend Sub Read(binReader As BinaryReader)
            Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", binReader.BaseStream.Position))
            mName = binReader.ReadString()                          '0x1D800   Pascal Varying Length String Format...
            binReader.BaseStream.Position += NamePasswordLengthMax - mName.Length
            mPassword = binReader.ReadString()                      '0x1D810   Pascal Varying Length String Format...
            binReader.BaseStream.Position += NamePasswordLengthMax - mPassword.Length

            mOut = binReader.ReadInt16()                            '0x1D820    00 00 = No; 01 00 = Yes;
            mRace = binReader.ReadInt16()                           '0x1D822    01 00 = Human
            mProfession = binReader.ReadInt16()                     '0x1D824    06 00 = Lord
            mAgeInWeeks = binReader.ReadInt16()                     '0x1D826    0C 03 = Weeks Alive...
            mStatus = binReader.ReadInt16()                         '0x1D828    00 00 = OK
            mAlignment = binReader.ReadInt16()                      '0x1D82A    02 00 = Neutral
            mStatistics = binReader.ReadInt32()                     '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
            binReader.BaseStream.Position += 4                      '0x1D830
            For i As Short = 0 To 2                                 '0x1D834
                mGP(i) = binReader.ReadInt16()
            Next i
            mItemCount = binReader.ReadInt16()                      '0x1D83A
            For i As Short = 0 To ItemListMax - 1                   '0x1D83C    List of Items (stowing not an option in Wiz01...)
                mItemList(i).Read(binReader)
            Next i
            For i As Short = 0 To 2                                 '0x1D87C
                mEXP(i) = binReader.ReadInt16()
            Next i
            mLVL.Read(binReader)                                    '0x1D882
            mHP.Read(binReader)                                     '0x1D886
            For i As Short = 0 To 7                                 '0x1D88A    Need to mask as bits...
                mSpellBooks(i) = binReader.ReadByte()
            Next i
            For i As Short = 0 To SpellLevelMax - 1                 '0x1D892
                mMageSpellPoints(i) = binReader.ReadInt16()
            Next i
            For i As Short = 0 To SpellLevelMax - 1                 '0x1D8A0
                mPriestSpellPoints(i) = binReader.ReadInt16()
            Next i
            binReader.BaseStream.Position += 2                      '0x1D8AE
            mAC = binReader.ReadInt16()                             '0x1D8B0
            binReader.BaseStream.Position += 24                     '0x1D8B2
            mLocation = binReader.ReadInt16()                       '0x1D8CA    Some sort of packed variable...
            mDown = binReader.ReadInt16()                           '0x1D8CC    Seems to be a simple 2-byte Int16...
            mHonors = binReader.ReadInt16()                         '0x1D8CE    Need more testing, but 1 = ">"
            '                                                       '0x1D8D0    Next Character Record...
        End Sub
#End Region
    End Class
    Public Class Item
        Public Sub New()
            mEquipped = 0
            mCursed = 0
            mIdentified = 0
            mItemCode = 0
        End Sub
        Public Sub New(ByVal ItemCode As Integer, ByVal equipped As Integer, ByVal Cursed As Integer, ByVal Identified As Integer)
            mEquipped = equipped
            mCursed = Cursed
            mIdentified = Identified
            mItemCode = ItemCode
        End Sub
        Private mEquipped As Integer
        Private mCursed As Integer
        Private mIdentified As Integer
        Private mItemCode As Integer
        Public Property ItemCode As Integer
            Get
                Return mItemCode
            End Get
            Set(value As Integer)
                mItemCode = value
            End Set
        End Property
        Public Property Cursed As Integer
            Get
                Return mCursed
            End Get
            Set(value As Integer)
                mCursed = value
            End Set
        End Property
        Public Property Equipped As Integer
            Get
                Return mEquipped
            End Get
            Set(value As Integer)
                mEquipped = value
            End Set
        End Property
        Public Property Identified As Integer
            Get
                Return mIdentified
            End Get
            Set(value As Integer)
                mIdentified = value
            End Set
        End Property
        Friend Sub Read(binReader As BinaryReader)
            mEquipped = binReader.ReadInt16()
            mCursed = binReader.ReadInt16()
            mIdentified = binReader.ReadInt16()
            mItemCode = binReader.ReadInt16()
        End Sub
    End Class
    Public Class Points
        Public Sub New()
            mCurrent = 0
            mMaximum = 0
        End Sub
        Private mCurrent As Integer = 0
        Private mMaximum As Integer = 0
        Public Property Current As Integer
            Get
                Return mCurrent
            End Get
            Set(value As Integer)
                mCurrent = value
            End Set
        End Property
        Public Property Maximum As Integer
            Get
                Return mMaximum
            End Get
            Set(value As Integer)
                mMaximum = value
            End Set
        End Property
        Friend Sub Read(binReader As BinaryReader)
            mCurrent = binReader.ReadInt16()
            mMaximum = binReader.ReadInt16()
        End Sub
    End Class
#End Region
#Region "Properties"
#Region "Enumerations"
    Public Enum enumAlignment As Integer
        Unaligned = 0
        Good = 1
        Neutral = 2
        Evil = 3
    End Enum
    Public Enum enumProfession As Integer
        Fighter = 0
        Mage = 1
        Priest =2
        Thief = 3
        Bishop =4
        Samurai = 5
        Lord =6
        Ninja = 7
    End Enum
    Public Enum enumRace As Byte
        Human = 1
        Elf =2
        Dwarf = 3
        Gnome =4
        Hobbit = 5
    End Enum
    Public Enum enumStatus As Integer
        OK = 0
        Afraid = 1
        Asleep = 2
        Paralyzed = 3
        Stoned = 4
        Dead = 5
        Ashes = 6
        LostDeleted = 7
    End Enum
#End Region
#Region "Declarations"
    Const CharactersMax As Integer = 20
    Const ItemMapMax As Integer = 100
    Const SpellMapMax As Integer = 50

    Private ItemList(ItemMapMax) As String
    Private Spells(SpellMapMax) As String

    Private mFileName As String = vbNullString
    Private mCharacters(19) As Character
#End Region
    Public ReadOnly Property FileName As String
        Get
            Return mFileName
        End Get
    End Property
    Public ReadOnly Property Characters As Character()
        Get
            Return mCharacters
        End Get
    End Property
#End Region
#Region "Methods"
    Public Sub Read()
        Const ScenarioDataOffset As Int32 = &H1D400
        Const CharacterDataOffset As Int32 = &H1D800
        Const ScenarioName As String = "PROVING GROUNDS OF THE MAD OVERLORD!"
        Dim binReader As BinaryReader = Nothing
        Try
            If Not File.Exists(mFileName) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", mFileName))
            binReader = New BinaryReader(File.Open(mFileName, FileMode.Open))
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
                mCharacters(iChar).Read(binReader) : Debug.WriteLine(String.Format("{0}{1}", vbTab, mCharacters(iChar).Name))
                mCharacters(iChar + 1).Read(binReader) : Debug.WriteLine(String.Format("{0}{1}", vbTab, mCharacters(iChar).Name))
                mCharacters(iChar + 2).Read(binReader) : Debug.WriteLine(String.Format("{0}{1}", vbTab, mCharacters(iChar).Name))
                mCharacters(iChar + 3).Read(binReader) : Debug.WriteLine(String.Format("{0}{1}", vbTab, mCharacters(iChar).Name))
                binReader.BaseStream.Position += 192
            Next iBlock
            For iChar As Short = 0 To 19
                Debug.WriteLine(String.Format("{0:00}) {1}", iChar + 1, mCharacters(iChar).Name))
            Next iChar
        Finally
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
#End Region


    Private Function Alignment(ByVal x As enumAlignment) As String
        Select Case x
            Case enumAlignment.Unaligned : Return "Unaligned"
            Case enumAlignment.Good : Return "Good"
            Case enumAlignment.Neutral : Return "Neutral"
            Case enumAlignment.Evil : Return "Evil"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function Honors(ByVal x As Integer) As String
        Select Case x
            Case &H0 : Return "> Chevron of Trebor"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function Status(ByVal x As enumStatus) As String
        Select Case x
            Case enumStatus.OK : Return "OK"
            Case enumStatus.Afraid : Return "Afraid"
            Case enumStatus.Asleep : Return "Asleep"
            Case enumStatus.Paralyzed : Return "Paralyzed"
            Case enumStatus.Stoned : Return "Stoned"
            Case enumStatus.Dead : Return "Dead"
            Case enumStatus.Ashes : Return "Ashes"
            Case enumStatus.LostDeleted : Return "Lost/Deleted"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function Profession(ByVal x As enumProfession) As String
        Select Case x
            Case enumProfession.Fighter : Return "Fighter"
            Case enumProfession.Mage : Return "Mage"
            Case enumProfession.Priest : Return "Priest"
            Case enumProfession.Thief : Return "Thief"
            Case enumProfession.Bishop : Return "Bishop"
            Case enumProfession.Samurai : Return "Samurai"
            Case enumProfession.Lord : Return "Lord"
            Case enumProfession.Ninja : Return "Ninja"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function Race(ByVal x As enumRace) As String
        Select Case x
            Case enumRace.Human : Return "Human"
            Case enumRace.Elf : Return "Elf"
            Case enumRace.Dwarf : Return "Dwarf"
            Case enumRace.Gnome : Return "Gnome"
            Case enumRace.Hobbit : Return "Hobbit"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function FormatItem(x As Item) As String
        '    Item = vbTab & ItemList(x.ItemCode) & "; Code: " & x.ItemCode & "; Equipped: "
        '    If x.Identified Then Item &= "; Identified"
        '    If x.Equipped Then Item &= "; **EQUIPPED**"
        '    If x.Cursed Then strItem &= "; --CURSED--"

        Return String.Format("{0}{1}{2}", vbTab, IIf(x.Cursed, "-", IIf(x.Equipped, "*", " ")), ItemList(x.ItemCode))
    End Function
    Private Function FormatPoints(x As Points) As String
        Return String.Format("{0:D}/{1:D}", x.Current, x.Maximum)
    End Function
    Public Function FormatSpell(Spell As Integer, Data As Byte, Offset As Integer) As String
        Return String.Format("[{0}] {1}", IIf((Data And 2 ^ Offset) = 2 ^ Offset, "X", " "), Spells(Spell)) '& vbTab & "[Spell: " & Spell & "; Data: " & Hex(Data) & "; Offset: " & Offset & "]"
    End Function
    Public Sub cvtBstrToSpells(bString As String, Data() As Byte)
        '        Dim i As Long
        '        Dim iByte As Integer
        '        Dim iOffset As Integer

        '        'First clear the target area...
        '        For i = 1 To UBound(Data)
        '            Data(i) = 0
        '        Next i

        '        'Now inch through the bit-string setting the appropriate bits in the appropriate bytes...
        '        For i = 0 To (UBound(Data) * 8) - 1
        '            iOffset = i Mod 8
        '            iByte = (i \ 8) + 1
        '            If Mid(bString, i + 1, 1) = "1" Then
        '                Data(iByte) = (Data(iByte) Or (2 ^ iOffset))
        '            End If
        '        Next i
    End Sub
    Public Function cvtSpellsToBstr(Data() As Byte) As String
        '        Dim i As Long
        '        Dim Offset As Long
        '        Dim bString As String

        '        bString = vbNullString
        '        For i = 1 To UBound(Data)
        '            For Offset = 0 To 7
        '                If (Data(i) And 2 ^ Offset) = 2 ^ Offset Then bString = bString & "1" Else bString = bString & "0"
        '            Next Offset
        '        Next i
        '        Wiz01cvtSpellsToBstr = bString
    End Function
    Public Function cvtStatisticsToLong(Stat1 As String, Stat2 As String, Stat3 As String, Stat4 As String, Stat5 As String, Stat6 As String) As Long
        '        cvtStatisticsToLong = (CLng(Stat1) * (2 ^ 0)) + (CLng(Stat2) * (2 ^ 5)) + (CLng(Stat3) * (2 ^ 10)) + (CLng(Stat4) * (2 ^ 16)) + (CLng(Stat5) * (2 ^ 21)) + (CLng(Stat6) * (2 ^ 26))
    End Function
    Public Function cvtStatisticToInt(xStatistics As Long, WhichStat As Integer) As Integer
        '        Select Case WhichStat
        '            Case 1  'Strength
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 0)) And &H1F)
        '            Case 2  'Intelligence
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 5)) And &H1F)
        '            Case 3  'Piety
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 10)) And &H1F)
        '            Case 4  'Vitality
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 16)) And &H1F)
        '            Case 5  'Agility
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 21)) And &H1F)
        '            Case 6  'Luck
        '                Wiz01cvtStatisticToInt = ((xStatistics \ (2 ^ 26)) And &H1F)
        '            Case Else
        '                Wiz01cvtStatisticToInt = 0
        '        End Select
    End Function
    Private Function Dumapic(xCharacter As Character) As String
        '        With xCharacter
        '            ' Always seems to be facing North when Quiting from within the Maze...
        '            Wiz01Dumapic = "Facing North; " & (.Location \ 100) & " East; " & (.Location Mod 100) & " North; " & .Down & " Down"  ' from the steps leading to the castle"
        '        End With
    End Function
    Public Sub DumpCharacter(xCharacter As Character)
        '        Dim i As Long
        '        Dim j As Long
        '        Dim k As Long
        '        Dim iChar As Long
        '        Dim errorCode As Long
        '        Dim strTemp As String
        '        Dim bString As String

        '        On Error GoTo ErrorHandler
        '        Debug.WriteLine(New String("="c, 80))
        '        With xCharacter
        '            Debug.WriteLine("Name:               " & vbTab & Left(.Name, .NameLength))
        '            Debug.WriteLine("Password:           " & vbTab & Left(.Password, .PasswordLength))

        '            If .Out = 1 Then
        '                Debug.WriteLine("On Expedition:      " & vbTab & "YES")
        '            Else
        '                Debug.WriteLine("On Expedition:      " & vbTab & "NO")
        '            End If
        '            Debug.WriteLine(Wiz01Dumapic(xCharacter))

        '            Debug.WriteLine("Race:               " & vbTab & Race(.Race))
        '            Debug.WriteLine("Profession:         " & vbTab & Profession(.Profession))
        '            Debug.WriteLine("Age:                " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)")
        '            Debug.WriteLine("Status:             " & vbTab & Status(.Status))
        '            Debug.WriteLine("Alignment:          " & vbTab & Alignment(.Alignment))
        '            Debug.WriteLine("Honors:             " & vbTab & Honors(.Honors))
        '            Debug.WriteLine("Level:              " & vbTab & strPoints(.LVL))
        '            Debug.WriteLine("Hit Points:         " & vbTab & strPoints(.HP))
        '            Debug.WriteLine("Gold Pieces:        " & vbTab & I6toD(.GP))
        '            Debug.WriteLine("Experience Points:  " & vbTab & I6toD(.EXP))
        '            Debug.WriteLine("Armor Class:        " & vbTab & .AC)
        '            'If .Honors > 0 Then
        '            '    Select Case .Honors
        '            '        Case 1 : Debug.WriteLine("Honors:             " & vbTab & """" > """")
        '            '        Case Else
        '            '    End Select
        '            'End If

        '            Debug.WriteLine(vbCrLf & "Basic Statistics...")
        '            Debug.WriteLine("Strength:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 1))
        '            Debug.WriteLine("Intelligence:       " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 2))
        '            Debug.WriteLine("Piety:              " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 3))
        '            Debug.WriteLine("Vitality:           " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 4))
        '            Debug.WriteLine("Agility:            " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 5))
        '            Debug.WriteLine("Luck:               " & vbTab & Wiz01cvtStatisticToInt(.Statistics, 6))

        '            Debug.WriteLine(vbCrLf & "List of Items (Currently carrying " & .ItemCount & " items)...")
        '            For i = 1 To .ItemCount
        '                Debug.WriteLine(strItem(.ItemList(i)))
        '            Next i

        '            Debug.WriteLine(" ")
        '            strTemp = "Mage Spell Points:    " & vbTab
        '            For i = 1 To 7
        '                strTemp = strTemp & .MageSpellPoints(i) & "/"
        '            Next i
        '            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        '            Debug.WriteLine(strTemp)

        '            strTemp = "Priest Spell Points:  " & vbTab
        '            For i = 1 To 7
        '                strTemp = strTemp & .PriestSpellPoints(i) & "/"
        '            Next i
        '            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        '            Debug.WriteLine(strTemp)

        '            Debug.WriteLine("Unknown Region #1 (4 bytes): ")
        '            Debug.WriteLine(strHex(.Unknown1, 4)) '& vbCrLf
        '            Debug.WriteLine("Unknown Region #2 (2 bytes): ")
        '            Debug.WriteLine(strHex(.Unknown2, 2)) '& vbCrLf
        '            Debug.WriteLine("Unknown Region #3 (30 bytes): ")
        '            Debug.WriteLine(strHex(.Unknown3, 30)) '& vbCrLf
        '        End With
    End Sub
    Public Function GetHonors(x As ListBox) As Integer
        '        GetHonors = 0
        '        If x.Selected(0) Then Wiz01GetHonors = (Wiz01GetHonors Or &H1)      ' >
        '        If x.Selected(1) Then Wiz01GetHonors = (Wiz01GetHonors Or &H4000)   ' G
        '        If x.Selected(2) Then Wiz01GetHonors = (Wiz01GetHonors Or &H800)    ' K
        '        If x.Selected(3) Then Wiz01GetHonors = (Wiz01GetHonors Or &H2000)   ' D
        '        If x.Selected(4) Then Wiz01GetHonors = (Wiz01GetHonors Or &H20)     ' *
    End Function
    Public Sub InitializeItemList()
        ItemList(0) = "Misc: Broken Item"
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
        ItemList(45) = "Magic: Scroll Of Lomilwa"
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
    Public Sub InitializeSpells()
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
    Public Function IsBishop(x As enumProfession) As Boolean
        Return CBool(x = enumProfession.Bishop)
    End Function
    Public Function IsMage(x As enumProfession) As Boolean
        Select Case x
            Case enumProfession.Bishop, enumProfession.Mage, enumProfession.Samurai : Return True
        End Select
        Return False
    End Function
    Public Function IsPriest(x As enumProfession) As Boolean
        Select Case x
            Case enumProfession.Priest, enumProfession.Bishop, enumProfession.Lord : Return True
        End Select
        Return False
    End Function
    Public Function IsLord(x As enumProfession) As Boolean
        Return CBool(x = enumProfession.Lord)
    End Function
    Public Function IsSamurai(x As enumProfession) As Boolean
        Return CBool(x = enumProfession.Samurai)
    End Function
    Public Function IsSpellCaster(x As enumProfession) As Boolean
        Select Case x
            Case enumProfession.Mage, enumProfession.Priest, enumProfession.Bishop, enumProfession.Lord, enumProfession.Samurai : Return True
        End Select
        Return False
    End Function
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
    '    Public Sub Read(ByVal strFile As String, xCharacters() As Wiz01Character)
    '        Dim i As Long
    '        Dim iChar As Integer
    '        Dim Offset As Long
    '        Dim Unit As Integer
    '        Dim errorCode As Long
    '        Dim ScenarioName As Wiz01ScenarioTag

    '        'Proving Grounds of the Mad Overlord supports up to 20 characters...
    '        'The layout is a little funky in that the Character structure seems
    '        'not to have lined-up evenly with the disk layout (1024 byte blocks
    '        'on disk)... So, characters start at offset 0x0001D800 and then are
    '        'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
    '        '(completing the 1K disk block), then another 4 blocks... for 5
    '        'total blocks of 4, making 20 characters.

    '        On Error GoTo ErrorHandler
    '        Unit = FreeFile()
    '        Open strFile For Binary Access Read Write Lock Read Write As #Unit
    '    Offset = Wiz01CharacterDataOffset
    '        For i = 1 To 5
    '            iChar = (i * 4) - 3
    '        Get #Unit, Offset, xCharacters(iChar)
    '        Get #Unit, , xCharacters(iChar + 1)
    '        Get #Unit, , xCharacters(iChar + 2)
    '        Get #Unit, , xCharacters(iChar + 3)
    '        Offset = Offset + 1024
    '        Next i

    '        For i = 1 To 20
    '            xCharacters(i).Name = Replace(xCharacters(i).Name, Chr(0), " ")
    '            xCharacters(i).Name = Left(xCharacters(i).Name, xCharacters(i).NameLength)
    '        Next i

    'ExitSub:
    '        Close #Unit
    '    Call SaveRegSetting("Environment", "UWAPath01", ParsePath(strFile, DrvDirNoSlash))
    '        Call SaveRegSetting("Environment", "Wiz01DataFile", ParsePath(strFile, FileNameBaseExt))
    '        Exit Sub

    'ErrorHandler:
    '        MsgBox Err.Description, vbExclamation, "Wiz01Read"
    '    Exit Sub
    '        Resume Next
    '    End Sub
    '    Public Sub SetHonors(x As ListBox, y As Integer)
    '        Dim i As Integer
    '        For i = 0 To x.ListCount - 1
    '            x.Selected(i) = False
    '        Next i

    '        If (y And &H1) = &H1 Then x.Selected(0) = True          ' >
    '        If (y And &H4000) = &H4000 Then x.Selected(1) = True    ' G
    '        If (y And &H800) = &H800 Then x.Selected(2) = True      ' K
    '        If (y And &H2000) = &H2000 Then x.Selected(3) = True    ' D
    '        If (y And &H20) = &H20 Then x.Selected(4) = True        ' *

    '        x.ListIndex = -1
    '    End Sub
    '    Public Sub Write(ByVal strFile As String, xCharacters() As Wiz01Character)
    '        Dim i As Long
    '        Dim iChar As Integer
    '        Dim Offset As Long
    '        Dim Unit As Integer
    '        Dim errorCode As Long

    '        'Proving Grounds of the Mad Overlord supports up to 20 characters...
    '        'The layout is a little funky in that the Character structure seems
    '        'not to have lined-up evenly with the disk layout (1024 byte blocks
    '        'on disk)... So, characters start at offset 0x0001D800 and then are
    '        'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
    '        '(completing the 1K disk block), then another 4 blocks... for 5
    '        'total blocks of 4, making 20 characters.

    '        On Error GoTo ErrorHandler
    '        Unit = FreeFile()
    '        Open strFile For Binary Access Read Write Lock Read Write As #Unit
    '    Offset = Wiz01CharacterDataOffset
    '        For i = 1 To 5
    '            iChar = (i * 4) - 3
    '            Put #Unit, Offset, xCharacters(iChar)
    '        Put #Unit, , xCharacters(iChar + 1)
    '        Put #Unit, , xCharacters(iChar + 2)
    '        Put #Unit, , xCharacters(iChar + 3)
    '        Offset = Offset + 1024
    '        Next i
    '        Close #Unit

    'ExitSub:
    '        Exit Sub

    'ErrorHandler:
    '        MsgBox Err.Description, vbExclamation, "Wiz01Write"
    '    Exit Sub
    '        Resume Next
    '    End Sub
End Class
#Region "Support Classes"
#End Region