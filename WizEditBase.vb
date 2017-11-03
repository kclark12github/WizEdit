'WizEditBase.vb
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/11/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Imports WizEdit.CharacterBase
Public Class WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        mBoxArt = BoxArt
        mCaption = Caption
        mIcon = Icon
        mParent = Parent
        ReDim mCharacters(Me.CharactersMax - 1)
        For iChar As Short = 0 To Me.CharactersMax - 1
            mCharacters(iChar) = New CharacterBase(Me)
        Next iChar
        LoadMasterSpellBooks()
    End Sub
#Region "Properties"
#Region "Declarations"
    Protected mBoxArt As Image = Nothing
    Protected mCaption As String = Nothing
    Protected mForm As Form = Nothing
    Protected mIcon As Icon = Nothing
    Protected mParent As Form = Nothing

    Protected mCharacters As CharacterBase()
    Protected mMageSpellBook As SpellBase()
    Protected mMasterItemList As ItemData()
    Protected mMasterMageSpellbook As Collection = New Collection()
    Protected mMasterPriestSpellbook As Collection = New Collection()
    Protected mPriestSpellBook As SpellBase()
    Private mScenarioDataPath As String = ""
    Private mRegKey As String = "Software\KClark Software"
#End Region
    Public Overridable ReadOnly Property AlignmentList As String()
        Get
            AlignmentList = {
                enumAlignment.Unaligned.ToString,
                enumAlignment.Good.ToString,
                enumAlignment.Neutral.ToString,
                enumAlignment.Evil.ToString
                }
        End Get
    End Property
    Public ReadOnly Property Characters As CharacterBase()
        Get
            Return mCharacters
        End Get
    End Property
    Public Overridable ReadOnly Property CharacterDataOffset As Int32
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property CharactersMax As Short
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property HonorsList As String()
        Get
            HonorsList = {
                "> Chevron of Trebor",
                "* - Star of Llylgamyn",
                "K - Knight of Gnilda",
                "G - Mark of Gnilda",
                "D - Descendant of Heroes",
                "@ - Heart of Abriel (?)"
                }
        End Get
    End Property
    Public ReadOnly Property Icon As Icon
        Get
            Return mIcon
        End Get
    End Property
    Public Overridable ReadOnly Property MasterItemList As ItemData()
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property MageSpellBook As SpellBase()
        Get
            If mMageSpellBook Is Nothing Then
                With mMasterMageSpellbook
                    'The order of the objects in the array is important. They must correspond to the bit position in the character data structure.
                    mMageSpellBook = {
                        .Item("Unknown"),
                        .Item("HALITO"), .Item("MOGREF"), .Item("KATINO"), .Item("DUMAPIC"),
                        .Item("DILTO"), .Item("SOPIC"),
                        .Item("MAHALITO"), .Item("MOLITO"),
                        .Item("MORLIS"), .Item("DALTO"), .Item("LAHALITO"),
                        .Item("MAMORLIS"), .Item("MAKANITO"), .Item("MADALTO"),
                        .Item("LAKANITO"), .Item("ZILWAN"), .Item("MASOPIC"), .Item("HAMAN"),
                        .Item("MALOR"), .Item("MAHAMAN"), .Item("TILTOWAIT")
                    }
                End With
            End If
            Return mMageSpellBook
        End Get
    End Property
    Public Overridable ReadOnly Property MageSpellList As String()
        Get
            Try
                Dim temp() As String = {}
                ReDim temp(Me.MageSpellBook.Length - 2)
                For i As Short = 0 To Me.MageSpellBook.Length - 2
                    temp(i) = Me.MageSpellBook(i + 1).Name
                Next i
                Return temp
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
                Throw
            End Try
        End Get
    End Property
    Public Overridable ReadOnly Property PriestSpellBook As SpellBase()
        Get
            If mPriestSpellBook Is Nothing Then
                With mMasterPriestSpellbook
                    'The order of the objects in the array is important. They must correspond to the bit position in the character data structure.
                    mPriestSpellBook = {
                        .Item("KALKI"), .Item("DIOS"), .Item("BADIOS"), .Item("MILWA"), .Item("PORFIC"),
                        .Item("MATU"), .Item("CALFO"), .Item("MANIFO"), .Item("MONTINO"),
                        .Item("LOMILWA"), .Item("DIALKO"), .Item("LATUMAPIC"), .Item("BAMATU"),
                        .Item("DIAL"), .Item("BADIAL"), .Item("LATUMOFIS"), .Item("MAPORFIC"),
                        .Item("DIALMA"), .Item("BADIALMA"), .Item("LITOKAN"), .Item("KANDI"), .Item("DI"), .Item("BADI"),
                        .Item("LORTO"), .Item("MADI"), .Item("MABADI"), .Item("LOKTOFEIT"),
                        .Item("MALIKTO"), .Item("KADORTO")
                    }
                End With
            End If
            Return mPriestSpellBook
        End Get
    End Property
    Public Overridable ReadOnly Property PriestSpellList As String()
        Get
            Try
                Dim temp() As String = {}
                ReDim temp(Me.PriestSpellBook.Length - 1)
                For i As Short = 0 To Me.PriestSpellBook.Length - 1
                    temp(i) = Me.PriestSpellBook(i).Name
                Next i
                Return temp
            Catch ex As Exception
                Debug.WriteLine(ex.ToString)
                Throw
            End Try
        End Get
    End Property
    Public Overridable ReadOnly Property ProfessionList As String()
        Get
            ProfessionList = {
                enumProfession.Fighter.ToString,
                enumProfession.Mage.ToString,
                enumProfession.Priest.ToString,
                enumProfession.Thief.ToString,
                enumProfession.Bishop.ToString,
                enumProfession.Samurai.ToString,
                enumProfession.Lord.ToString,
                enumProfession.Ninja.ToString
                }
        End Get
    End Property
    Public Overridable ReadOnly Property RaceList As String()
        Get
            RaceList = {
                enumRace.NoRace.ToString,
                enumRace.Human.ToString,
                enumRace.Elf.ToString,
                enumRace.Dwarf.ToString,
                enumRace.Gnome.ToString,
                enumRace.Hobbit.ToString
                }
        End Get
    End Property
    Public Overridable ReadOnly Property RegDataPath As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property ScenarioDataOffset As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable Property ScenarioDataPath As String
        Get
            If mScenarioDataPath = "" Then mScenarioDataPath = GetRegSetting("Environment", Me.RegDataPath, "")
            If mScenarioDataPath = "" Then Throw New FileNotFoundException("Save Game file not found!")
            If Not File.Exists(mScenarioDataPath) Then Throw New FileNotFoundException("Save Game file not found!")
            Return mScenarioDataPath
        End Get
        Set(value As String)
            Dim binReader As BinaryReader = Nothing
            Try
                If Not File.Exists(mScenarioDataPath) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", mScenarioDataPath))
                binReader = New BinaryReader(File.Open(mScenarioDataPath, FileMode.Open))
                binReader.BaseStream.Position = Me.ScenarioDataOffset
                Dim myScenarioName As String = binReader.ReadString()
                If myScenarioName <> Me.ScenarioName Then Throw New NotSupportedException(String.Format("Save game file specified is not a valid Ultimate Wizardry Archives: {0} save game file.", Me.ScenarioName))
                Me.SaveRegSetting("Environment", Me.RegDataPath, value)
                mScenarioDataPath = value
            Finally
                If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
            End Try
        End Set
    End Property
    Public Overridable ReadOnly Property ScenarioName As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property StatusList As String()
        Get
            StatusList = {
                enumStatus.OK.ToString,
                enumStatus.Afraid.ToString,
                enumStatus.Asleep.ToString,
                enumStatus.Paralyzed.ToString,
                enumStatus.Stoned.ToString,
                enumStatus.Dead.ToString,
                enumStatus.Ashes.ToString,
                "Lost/Deleted"
                }
        End Get
    End Property
#End Region
#Region "Methods"
#Region "Conversion"
    Public Sub DtoI6(ByVal x As Double, ByRef Data() As UInt16)
        Dim r1 As Double
        Dim r2 As Double
        Dim r3 As Double

        r3 = x \ 100000000.0#
        Data(2) = CInt(r3)
        x = x - (r3 * 100000000.0#)
        r2 = x \ 10000.0#
        Data(1) = r2
        x = x - (r2 * 10000.0#)
        r1 = x
        Data(0) = r1
    End Sub
    Public Function I6toD(ByVal Data() As UInt16) As Double
        Dim r1 As Double
        Dim r2 As Double
        Dim r3 As Double

        r1 = Data(0)
        r2 = Data(1) * 10000.0#
        r3 = Data(2) * 100000000.0#

        I6toD = r1 + r2 + r3
    End Function
#End Region
#Region "Registry"
    Public Function GetRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal [Default] As Object) As Object
        Dim Reg As RegistryKey = Nothing
        GetRegSetting = Nothing
        Try
            Reg = Registry.CurrentUser.OpenSubKey(String.Format("{0}\{1}\{2}", mRegKey, Application.ProductName, KeyName)) : If IsNothing(Reg) Then Exit Try
            GetRegSetting = Reg.GetValue(ValueName, [Default])
        Catch ex As System.Exception
        Finally : If Not IsNothing(Reg) Then Reg.Close()
        End Try
    End Function
    Public Sub SaveRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Object)
        Dim Reg As RegistryKey = Nothing
        Dim CurrentValue As Object = Nothing
        Try
            If KeyName Is Nothing Then Throw New ArgumentException("KeyName must be provided.")
            If ValueName Is Nothing Then Throw New ArgumentException("ValueName must be provided.")
            If Value Is Nothing Then Throw New ArgumentException("Value must be provided.")

            KeyName = String.Format("{0}\{1}\{2}", mRegKey, Application.ProductName, KeyName)
            Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            If Reg Is Nothing Then
                'Iterate through the KeyName making sure each sub-key exists (create as necessary)...
                Dim SubKeys() As String = KeyName.Split("\")
                Dim Key As String = SubKeys(0)
                For i As Short = 1 To SubKeys.Length - 1
                    Dim SubKey As String = String.Format("{0}\{1}", Key, SubKeys(i))
                    Reg = Registry.CurrentUser.OpenSubKey(SubKey)
                    If Reg Is Nothing Then
                        Reg = Registry.CurrentUser.OpenSubKey(Key, True)
                        Reg.CreateSubKey(SubKeys(i))
                    End If
                    Reg.Close() : Reg = Nothing
                    Key = SubKey
                Next i
                Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            End If
            CurrentValue = Reg.GetValue(ValueName)
            If CurrentValue Is Nothing OrElse CurrentValue.ToString <> Value.ToString Then Reg.SetValue(ValueName, Value)
        Catch ex As System.Exception
        Finally : If Reg IsNot Nothing Then Reg.Close()
        End Try
    End Sub
#End Region
#Region "Utility"
    Protected Sub Backup()
        Dim fi As FileInfo = New FileInfo(Me.ScenarioDataPath)
        Dim backup As String = fi.Name.Replace(fi.Extension, String.Format(".{0:yyyyMMdd.HHmmssff}{1}", fi.LastWriteTime, fi.Extension))
        Dim backupPath As String = String.Format("{0}\{1}", fi.DirectoryName, backup)
        If Not FileIO.FileSystem.FileExists(backupPath) Then
            FileIO.FileSystem.RenameFile(mScenarioDataPath, backup)
            FileIO.FileSystem.CopyFile(backupPath, mScenarioDataPath, FileIO.UIOption.OnlyErrorDialogs)
        End If
    End Sub
    Public Function IsPrintable(ByVal xByte As Byte) As Boolean
        If xByte < 32 Then Return False
        Select Case xByte
            Case 127, 129, 141, 143, 144, 157 : Return False
            Case Else : Return True
        End Select
    End Function
    Public Function UpCase(uKey As Integer) As Integer
        If uKey > 96 And uKey < 123 Then
            UpCase = uKey - 32
        Else
            UpCase = uKey
        End If
    End Function
    Public Function ValidateByte(ByVal ctl As Control) As Boolean
        ValidateByte = False
        Try
            With ctl
                If .Text = "" Then .Text = "0"
                Dim iLimit As Byte = 99
                If Val(.Text) < 0 Or Val(.Text) > iLimit Then Beep() : .Text = "" : Exit Try
                '.Text = Format(.Text, "00")
            End With
            ValidateByte = True
        Finally
        End Try
    End Function
    Public Function ValidateI2(ByVal ctl As Control) As Boolean
        ValidateI2 = False
        Try
            With ctl
                If .Text = vbNullString Then .Text = "0"
                Dim iLimit As UInt16 = (2 ^ 16) - 1
                If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then Beep() : .Text = "" : Exit Try
                .Text = Format(.Text, "#,##0")
            End With
            ValidateI2 = True
        Finally
        End Try
    End Function
    Public Function ValidateI4(ByVal ctl As Control) As Boolean
        ValidateI4 = False
        Try
            With ctl
                If .Text = vbNullString Then .Text = "0"
                Dim iLimit As UInt32 = (2 ^ 31) - 1
                If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then Beep() : .Text = "" : Exit Try
                .Text = Format(.Text, "#,##0")
            End With
            ValidateI4 = True
        Finally
        End Try
    End Function
#End Region
    Public Function EXPRequiredForNextLevel(ByVal LVL As Integer, ByVal Profession As enumProfession) As Decimal
        EXPRequiredForNextLevel = Nothing
        Select Case Profession
            Case enumProfession.Fighter
                Select Case LVL
                    Case 0 : Return 1000
                    Case 1 : Return 1724
                    Case 2 : Return 2972
                    Case 3 : Return 5124
                    Case 4 : Return 8834
                    Case 5 : Return 15231
                    Case 6 : Return 26250
                    Case 7 : Return 45275
                    Case 8 : Return 78060
                    Case 9 : Return 134586
                    Case 10 : Return 232044
                    Case 11 : Return 400075
                    Case Else : Return 400075 + (289709 * (LVL - 11))
                End Select
            Case enumProfession.Mage
                Select Case LVL
                    Case 0 : Return 1100
                    Case 1 : Return 1896
                    Case 2 : Return 3258
                    Case 3 : Return 5634
                    Case 4 : Return 9713
                    Case 5 : Return 16746
                    Case 6 : Return 28872
                    Case 7 : Return 49779
                    Case 8 : Return 85825
                    Case 9 : Return 147974
                    Case 10 : Return 255127
                    Case 11 : Return 439874
                    Case Else : Return 439874 + (318529 * (LVL - 11))
                End Select
            Case enumProfession.Priest
                Select Case LVL
                    Case 0 : Return 1050
                    Case 1 : Return 1810
                    Case 2 : Return 3120
                    Case 3 : Return 5379
                    Case 4 : Return 9274
                    Case 5 : Return 15989
                    Case 6 : Return 27567
                    Case 7 : Return 47529
                    Case 8 : Return 81946
                    Case 9 : Return 141286
                    Case 10 : Return 243596
                    Case 11 : Return 419993
                    Case Else : Return 419993 + (304132 * (LVL - 11))
                End Select
            Case enumProfession.Thief
                Select Case LVL
                    Case 0 : Return 900
                    Case 1 : Return 1551
                    Case 2 : Return 2674
                    Case 3 : Return 4610
                    Case 4 : Return 7948
                    Case 5 : Return 13703
                    Case 6 : Return 23625
                    Case 7 : Return 40732
                    Case 8 : Return 70227
                    Case 9 : Return 121081
                    Case 10 : Return 208760
                    Case 11 : Return 359931
                    Case Else : Return 359931 + (260639 * (LVL - 11))
                End Select
            Case enumProfession.Bishop
                Select Case LVL
                    Case 0 : Return 1200
                    Case 1 : Return 2105
                    Case 2 : Return 3692
                    Case 3 : Return 5477
                    Case 4 : Return 11353
                    Case 5 : Return 19935
                    Case 6 : Return 34973
                    Case 7 : Return 61356
                    Case 8 : Return 107642
                    Case 9 : Return 188845
                    Case 10 : Return 331307
                    Case 11 : Return 581240
                    Case Else : Return 581240 + (438479 * (LVL - 11))
                End Select
            Case enumProfession.Samurai
                Select Case LVL
                    Case 0 : Return 1250
                    Case 1 : Return 2192
                    Case 2 : Return 3845
                    Case 3 : Return 5745
                    Case 4 : Return 11833
                    Case 5 : Return 20759
                    Case 6 : Return 36419
                    Case 7 : Return 63892
                    Case 8 : Return 112091
                    Case 9 : Return 196650
                    Case 10 : Return 345000
                    Case 11 : Return 605263
                    Case Else : Return 605263 + (456601 * (LVL - 11))
                End Select
            Case enumProfession.Lord
                Select Case LVL
                    Case 0 : Return 1300
                    Case 1 : Return 2280
                    Case 2 : Return 4000
                    Case 3 : Return 7017
                    Case 4 : Return 12310
                    Case 5 : Return 21595
                    Case 6 : Return 37887
                    Case 7 : Return 66458
                    Case 8 : Return 116610
                    Case 9 : Return 204578
                    Case 10 : Return 358908
                    Case 11 : Return 629663
                    Case Else : Return 629663 + (475008 * (LVL - 11))
                End Select
            Case enumProfession.Ninja
                Select Case LVL
                    Case 0 : Return 1450
                    Case 1 : Return 2543
                    Case 2 : Return 4451
                    Case 3 : Return 7826
                    Case 4 : Return 13729
                    Case 5 : Return 24085
                    Case 6 : Return 42254
                    Case 7 : Return 74129
                    Case 8 : Return 130050
                    Case 9 : Return 228157
                    Case 10 : Return 400275
                    Case 11 : Return 702236
                    Case Else : Return 702236 + (529756 * (LVL - 11))
                End Select
        End Select
    End Function
    Public Overridable Function GetCharacter(ByVal Tag As String) As CharacterBase
        For iChar As Short = 0 To mCharacters.Length - 1
            If mCharacters(iChar).Tag = Tag Then Return mCharacters(iChar)
        Next iChar
        Return Nothing
    End Function
    Private Sub LoadMasterSpellBooks()
        'Note that here the order of the objects in each collection is not important
        With mMasterMageSpellbook
            .Add(New SpellBase("Unknown", "Unknown", "<placeholder>", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 0), "Unknown")
            'Level 1
            .Add(New SpellBase("DUMAPIC", "Clarity", "Informs you of the party's exact position from the stairs to the castle", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 1), "DUMAPIC")
            .Add(New SpellBase("HALITO", "Little Fire", "Causes a flame ball the size of a baseball to hit a monster for 1-8 points damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 1), "HALITO")
            .Add(New SpellBase("KATINO", "Bad Air", "Causes most of the monsters in a group to fall asleep", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 1), "KATINO")
            .Add(New SpellBase("MOGREF", "Body Iron", "Reduces the casters armor class by 2 for the encounter", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 1), "MOGREF")

            'Level 2
            'Wizardry [1-4]-Only
            .Add(New SpellBase("DILTO", "Darkness", "Causes one group of monsters to be enveloped in darkness, which reduces their ability to defend against your attacks ", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 2), "DILTO")
            .Add(New SpellBase("SOPIC", "Glass", "Causes the caster to become transparent, thus reducing their armor class by 4", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 2), "SOPIC")
            'Wizardry 5-Only
            .Add(New SpellBase("BOLATU", "Heart of Stone", "Attempts to stone one monster", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 2), "BOLATU")
            .Add(New SpellBase("DESTO", "Unlock", "Gives the caster thief skills of the same level to try and unlock doors", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 2), "DESTO")
            .Add(New SpellBase("MELITO", "Little Sparks", "Causes 1 to 8 points of damage to a monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 2), "MELITO")
            .Add(New SpellBase("PONTI", "Speed", "Reduces a party member's AC by one and makes them quicker in combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Mage, 2), "PONTI")

            'Level 3
            .Add(New SpellBase("MAHALITO", "Big Fire", "Causes a fiery explosion in a monster group, doing 4-24 points damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 3), "MAHALITO")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("MOLITO", "Sparks", "Causes sparks to damage half of the monsters in a group for 3-18 points damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 3), "MOLITO")
            'Wizardry 5-Only
            .Add(New SpellBase("CALIFIC", "Reveal", "Shows secret doors while exploring", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 3), "CALIFIC")
            .Add(New SpellBase("CORTU", "Magic Screen", "Erects a protective barrier from breathing monsters during combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 3), "CORTU")
            .Add(New SpellBase("KANTIOS", "Disruption", "Attempts to confuse a monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 3), "KANTIOS")

            'Level 4
            .Add(New SpellBase("LAHALITO", "Torch", "Does 6-36 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4), "LAHALITO")
            .Add(New SpellBase("MORLIS", "Fear", "Causes a group of monsters to fear the party, twice as powerful as DILTO", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4), "MORLIS")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("DALTO", "Blizzard", "Does 6-36 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4), "DALTO")
            'Wizardry 5-Only
            .Add(New SpellBase("LITOFEIT", "Levitate", "Helps the party avoid traps while exploring", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 4), "LITOFEIT")
            .Add(New SpellBase("ROKDO", "Stun", "Attempts to confuse and stun a group of monsters", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4), "ROKDO")
            .Add(New SpellBase("TZALIK", "Fist of God", "Hits a monster for 24 to 58 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 4), "TZALIK")

            'Level 5
            .Add(New SpellBase("MADALTO", "Frost/Frost King", "Causes 8-64 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 5), "MADALTO")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("MAKANITO", "Deadly Air", "Kills any monsters of less than 8th level (about 35-40 hit points)", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 5), "MAKANITO")
            .Add(New SpellBase("MAMORLIS", "Terror", "Causes all monsters to fear the party", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 5), "MAMORLIS")
            'Wizardry 5-Only
            .Add(New SpellBase("BACORTU", "Fizzle Field", "Erects a spell dampening field around a monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 5), "BACORTU")
            .Add(New SpellBase("PALIOS", "Anti-Magic", "Destroys monster built spell dampening fields", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 5), "PALIOS")
            .Add(New SpellBase("SOCORDI", "Conjure", "Summons an elemental to fight for the party during combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 5), "SOCORDI")
            .Add(New SpellBase("VASKYRE", "Rainbow Rays", "Random damaging effects to a monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 5), "VASKYRE")

            'Level 6
            .Add(New SpellBase("ZILWAN", "Dispel", "[Wizardry 1-4] Will destroy any one undead monster;[Wizardry 5] Causes 500-1000 points of damage to an undead monster", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 6), "ZILWAN")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("HAMAN", "Change/Beg", "Has random effects, and drains the caster one level (See MAHAMAN)", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Variable, SpellBase.enumSpellCategory.Mage, 6), "HAMAN")
            .Add(New SpellBase("LAKANITO", "Suffocation/Vacuum", "Kills all monsters affected by this spell, but some monsters are immune", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 6), "LAKANITO")
            .Add(New SpellBase("MASOPIC", "Big Glass/Crystal", "Reduces the armor class of the entire party by 4", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 6), "MASOPIC")
            'Wizardry 5-Only
            .Add(New SpellBase("LADALTO", "Ice Storm", "Freezes a monster group for 34 to 98 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 6), "LADALTO")
            .Add(New SpellBase("LOKARA", "Earth Feast", "Attempts to eliminate all monsters with varying success", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 6), "LOKARA")
            .Add(New SpellBase("MAMOGREF", "Wall of Force", "Erects an AC -10 field around a party member", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Mage, 6), "MAMOGREF")

            'Level 7
            .Add(New SpellBase("MAHAMAN", "Great Change/Beseech", "Does something random, stronger than Haman. Drains the caster one experience level, and is forgotten when cast. In some versions the caster can choose from a list of three possible effects. In the Wizardry Archives, you cannot choose in Scenario 1, but you can in Scenario 2 (useful for facing the KOD items). These are the possible effects: Silence the Monsters; Make Magic More Effective; DIALKO the Party 3 Times; Heal the Party; Destroy the Monsters; Protect the Party; Teleport the Monsters; Reanimate Corpses", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Variable, SpellBase.enumSpellCategory.Mage, 7), "MAHAMAN")
            .Add(New SpellBase("MALOR", "Teleport", "Teleports the party randomly within the current level when used in melee, but when cast in camp, you can decide exactly where you want to go. If a party teleports into stone it is LOST forever, so the spell is best used in conjunction with DUMAPIC. Some levels of the dungeon (1-10 and 2-6, for example) contain magnetic fields that bounce back incoming teleports.", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 7), "MALOR")
            .Add(New SpellBase("TILTOWAIT", "Ka-Blam!", "The effect of this spell is somewhat like the detonation of a small tactical nuclear weapon. The party is protected from its effects. Unfortunately for the monsters, they are not. The spell causes 10-100 hit points of damage to all monsters.", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 7), "TILTOWAIT")
            'Wizardry 5-Only
            .Add(New SpellBase("ABRIEL", "Divine Wish", "Alas, only the vanished Gatekeeper knows this spell", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Variable, SpellBase.enumSpellCategory.Mage, 7), "ABRIEL")
            .Add(New SpellBase("MAWXIWTZ", "Mad House", "Causes random but usually devastating effects to all monsters", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 7), "MAWXIWTZ")
        End With
        With mMasterPriestSpellbook
            'Level 1
            .Add(New SpellBase("KALKI", "Blessings", "Reduces the armor class of all party members by one during combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 1), "KALKI")
            .Add(New SpellBase("DIOS", "Heal", "Restores from one to eight points of damage to a party member", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 1), "DIOS")
            .Add(New SpellBase("BADIOS", "Harm", "Causes one to eight points of damage to a monster", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 1), "BADIOS")
            .Add(New SpellBase("MILWA", "Light", "Causes a softly glowing light to follow the party, increasing vision and revealing secret doors", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 1), "MILWA")
            .Add(New SpellBase("PORFIC", "Shield", "Lowers the armor class of the caster a little by 4 during combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 1), "PORFIC")

            'Level 2
            .Add(New SpellBase("CALFO", "X-Ray Vision", "Allows the caster to decide what the trap on a chest is 95% of the time", SpellBase.enumSpellType.Looting, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 2), "CALFO")
            .Add(New SpellBase("MANIFO", "Statue", "Causes some of the monsters to become paralyzed temporarily", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2), "MANIFO")
            .Add(New SpellBase("MONTINO", "Still Air", "Causes the air around a group of monsters to stop transmitting sounds, and therefore makes it impossible for them to cast spells", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2), "MONTINO")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("MATU", "Zeal", "Lowers armor class of all party members by two during combat", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 2), "MATU")
            'Wizardry 5-Only
            .Add(New SpellBase("KATU", "Charm", "Attempts to charm a monster", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 2), "KATU")

            'Level 3
            .Add(New SpellBase("LOMILWA", "More Light/Sunbeam", "A more powerful MILWA spell that lasts for the entire expedition, but is terminated upon entering a darkness area", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3), "LOMILWA")
            .Add(New SpellBase("DIALKO", "Softness", "Cures paralysis, and cures the effects of MANIFO and KATINO", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 3), "DIALKO")
            .Add(New SpellBase("LATUMAPIC", "Identify", "Tells you exactly what the monsters really are", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3), "LATUMAPIC")
            .Add(New SpellBase("BAMATU", "Prayer", "Lowers the party's armor class by four in combat [three in Wizardry 5]", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3), "BAMATU")
            'Wizardry 5-Only
            .Add(New SpellBase("HAKANIDO", "Magic Drain", "Attempts to drain a monster of upper magic powers", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 2), "HAKANIDO")

            'Level 4
            .Add(New SpellBase("DIAL", "More Heal/Cure", "Heals 2 to 16 points of damage", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 4), "DIAL")
            .Add(New SpellBase("BADIAL", "More Hurt/Wound", "Causes 2 to 16 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 4), "BADIAL")
            .Add(New SpellBase("LATUMOFIS", "Cure Poison/Cleanse", "Cures poisoning", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 4), "LATUMOFIS")
            .Add(New SpellBase("MAPORFIC", "Big Shield", "Lowers the party's armor class by 2, and lasts for the entire expedition", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 4), "MAPORFIC")
            'Wizardry 5-Only
            .Add(New SpellBase("BARIKO", "Razor Wind", "Causes 6 to 15 points of damage to a monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2), "BARIKO")

            'Level 5
            .Add(New SpellBase("DIALMA", "Great Heal/Big Cure", "Restores 3 to 24 hit points", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 5), "DIALMA")
            .Add(New SpellBase("KANDI", "Locate Soul", "Gives the direction of the person the party is attempting to locate and is relative to the position of the caster", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 5), "KANDI")
            .Add(New SpellBase("DI", "Life", "Causes a dead person to be resurrected, but the character has only 1 hit point and decreased vitality, and it doesn't always work (In which case a dead character is turned to ashes)", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 5), "DI")
            .Add(New SpellBase("BADI", "Death", "Gives a monster a coronary attack, which may or may not cause death", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 5), "BADI")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("BADIALMA", "Great Hurt/Big Wound", "Causes 3 to 24 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 5), "BADIALMA")
            .Add(New SpellBase("LITOKAN", "Flames", "Causes a pillar of flame to strike a group of monsters, doing 3 to 24 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 5), "LITOKAN")
            'Wizardry 5-Only
            .Add(New SpellBase("BAMORDI", "Summoning", "Attempts to summon one group of monsters from the elemental planes to fight for the party", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 2), "BAMORDI")
            .Add(New SpellBase("MOGATO", "Astral Gate", "Attempts to banish a demon monster back from whence it came", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 2), "MOGATO")

            'Level 6
            .Add(New SpellBase("LOKTOFEIT", "Recall", "[Wizardry 1-3] Causes all party members to be transported back to the castle, minus all of their equipment and most of their gold;[Wizardry 5] Party is transported back to the castle with all of their equipment and gold, but the spell is forgotten after casting and must be relearned, and there is a chance the spell will not work ", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 6), "LOKTOFEIT")
            .Add(New SpellBase("MADI", "Healing/Restore", "Causes all hit points to be restored and cures any condition except death", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 6), "MADI")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("LORTO", "Blades", "Causes sharp blades to slice through a group, causing 6 to 36 points of damage", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 6), "LORTO")
            .Add(New SpellBase("MABADI", "Harming/Maiming", "Causes all but 1 to 8 hit points to be removed from a target", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 6), "MABADI")
            'Wizardry 5-Only
            .Add(New SpellBase("KAKAMEN", "Fire Wind", "Causes 18 to 38 points of damage to one monster group", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2), "KAKAMEN")
            .Add(New SpellBase("LABADI", "Life Steal", "Attempts to drain all but 1 to 8 points from a monster, and transfer the life force to heal the caster", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 2), "LABADI")

            'Level 7
            .Add(New SpellBase("KADORTO", "Rebirth", "Restores the dead to life, and restores all hit points, even if the character is ashes, but if the spell fails the character is LOST forever", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 7), "KADORTO")
            'Wizardry [1-4]-Only
            .Add(New SpellBase("MALIKTO", "Word of Death/Wrath", "Causes 12 to 72 hit points of damage to all monsters", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Priest, 7), "MALIKTO")
            'Wizardry 5-Only
            .Add(New SpellBase("BAKADI", "Death Wind", "Attempts to slay one group of monsters", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2), "BAKADI")
            .Add(New SpellBase("IHALON", "Wish", "Grants a special favor to a party member, but is forgotten after being cast", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 2), "IHALON")
            .Add(New SpellBase("MABARIKO", "Meteor Winds", "Causes 18 to 58 points of damage to all monsters", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Priest, 2), "MABARIKO")
        End With
    End Sub
    Public Overridable Sub Read()
        Dim binReader As BinaryReader = Nothing
        Try
            'Wizardry (1-3) supports up to 20 characters...
            'The layout is a little funky in that the Character structure seems
            'not to have lined-up evenly with the disk layout (1024 byte blocks
            'on disk)... So, characters start at offset 0x0001D800 and then are
            'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
            '(completing the 1K disk block), then another 4 blocks... for 5
            'total blocks of 4, making 20 characters.
            binReader = New BinaryReader(File.Open(Me.ScenarioDataPath, FileMode.Open))
            binReader.BaseStream.Position = Me.CharacterDataOffset
            For iBlock As Short = 1 To 5
                Dim iChar As Short = (iBlock * 4) - 4
                Characters(iChar).Read(binReader)
                Characters(iChar + 1).Read(binReader)
                Characters(iChar + 2).Read(binReader)
                Characters(iChar + 3).Read(binReader)
                binReader.BaseStream.Position += 192
            Next iBlock
            'For iChar As Short = 0 To 19
            '    If Characters(iChar).Name <> "" Then
            '        Debug.WriteLine(New String("="c, 132))
            '        Debug.WriteLine(String.Format("{0:00}) {1}", iChar + 1, Characters(iChar).Name))
            '        Debug.WriteLine(New String("-"c, 132))
            '        Debug.WriteLine(Characters(iChar).ToString)
            '    End If
            'Next iChar
        Finally
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
    Public Overridable Sub Save()
        Dim binWriter As BinaryWriter = Nothing
        Try
            Backup()
            binWriter = New BinaryWriter(File.Open(mScenarioDataPath, FileMode.Open, FileAccess.Write, FileShare.None))
            binWriter.BaseStream.Position = Me.CharacterDataOffset
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
    Public Overridable Sub Show()
        mForm = New frmWizardry15Base(Me, mCaption, mIcon, mBoxArt)
        mForm.ShowDialog(mParent)
    End Sub
#End Region
End Class