'WizEditBase.cls
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
    End Sub
#Region "Properties"
#Region "Declarations"
    Private mBoxArt As Image = Nothing
    Private mCaption As String = Nothing
    Private mForm As Form = Nothing
    Private mIcon As Icon = Nothing
    Private mParent As Form = Nothing

    Private mCharacters As CharacterBase()
    Protected mMasterItemList As ItemData()
    Protected mMageSpellBook As SpellBase()
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
                "G - Mark of Gnilda",
                "K - Knight of Gnilda",
                "D - Descendant of Heroes",
                "* - Star of Llylgamyn"
                }
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
                mMageSpellBook = {
                    New SpellBase("Unknown", "Unknown", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 0),
                    New SpellBase("HALITO", "Little Fire", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 1),
                    New SpellBase("MOGREF", "Body Iron", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 1),
                    New SpellBase("KATINO", "Bad Air", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 1),
                    New SpellBase("DUMAPIC", "Clarity", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 1),
                    New SpellBase("DILTO", "Darkness", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 2),
                    New SpellBase("SOPIC", "Glass", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Mage, 2),
                    New SpellBase("MAHALITO", "Big Fire", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 3),
                    New SpellBase("MOLITO", "Sparks", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 3),
                    New SpellBase("MORLIS", "Fear", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4),
                    New SpellBase("DALTO", "Blizzard", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4),
                    New SpellBase("LAHALITO", "Torch", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 4),
                    New SpellBase("MAMORLIS", "Terror", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 5),
                    New SpellBase("MAKANITO", "Deadly Air", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 5),
                    New SpellBase("MADALTO", "Frost King", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 5),
                    New SpellBase("LAKANITO", "Vacuum", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Mage, 6),
                    New SpellBase("ZILWAN", "Dispell", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Mage, 6),
                    New SpellBase("MASOPIC", "Crystal", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 6),
                    New SpellBase("HAMAN", "Beg", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Variable, SpellBase.enumSpellCategory.Mage, 6),
                    New SpellBase("MALOR", "Teleport", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Mage, 7),
                    New SpellBase("MAHAMAN", "Beseech", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Variable, SpellBase.enumSpellCategory.Mage, 7),
                    New SpellBase("TILTOWAIT", "Ka-Blam!", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Mage, 7)
                }
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
                mPriestSpellBook = {
                    New SpellBase("KALKI", "Blessings", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 1),
                    New SpellBase("DIOS", "Heal", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 1),
                    New SpellBase("BADIOS", "Harm", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 1),
                    New SpellBase("MILWA", "Light", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 1),
                    New SpellBase("PORFIC", "Shield", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 1),
                    New SpellBase("MATU", "Zeal", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 2),
                    New SpellBase("CALFO", "X-Ray", SpellBase.enumSpellType.Looting, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 2),
                    New SpellBase("MANIFO", "Statue", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2),
                    New SpellBase("MONTINO", "Still Air", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 2),
                    New SpellBase("LOMILWA", "Sunbeam", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3),
                    New SpellBase("DIALKO", "Softness", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 3),
                    New SpellBase("LATUMAPIC", "Identify", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3),
                    New SpellBase("BAMATU", "Prayer", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 3),
                    New SpellBase("DIAL", "Cure", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 4),
                    New SpellBase("BADIAL", "Wound", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 4),
                    New SpellBase("LATUMOFIS", "Cleanse", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 4),
                    New SpellBase("MAPORFIC", "Big Shield", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 4),
                    New SpellBase("DIALMA", "Big Cure", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("BADIALMA", "Big Wound", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("LITOKAN", "Flames", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("KANDI", "Location", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.Caster, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("DI", "Life", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("BADI", "Death", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 5),
                    New SpellBase("LORTO", "Blades", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneGroup, SpellBase.enumSpellCategory.Priest, 6),
                    New SpellBase("MADI", "Restore", SpellBase.enumSpellType.AnyTime, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 6),
                    New SpellBase("MABADI", "Maiming", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.OneMonster, SpellBase.enumSpellCategory.Priest, 6),
                    New SpellBase("LOKTOFEIT", "Recall", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.EntireParty, SpellBase.enumSpellCategory.Priest, 6),
                    New SpellBase("MALIKTO", "Wrath", SpellBase.enumSpellType.Combat, SpellBase.enumSpellAffects.AllMonsters, SpellBase.enumSpellCategory.Priest, 7),
                    New SpellBase("KADORTO", "Rebirth", SpellBase.enumSpellType.Camp, SpellBase.enumSpellAffects.OnePerson, SpellBase.enumSpellCategory.Priest, 7)
                }
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
    Public Function GetCharacter(ByVal Tag As String) As CharacterBase
        For iChar As Short = 0 To mCharacters.Length - 1
            If mCharacters(iChar).Tag = Tag Then Return mCharacters(iChar)
        Next iChar
        Return Nothing
    End Function
    Public Sub Read()
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
    Public Sub Save()
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
    Public Sub Show()
        mForm = New frmWizardry15Base(Me, mCaption, mIcon, mBoxArt)
        mForm.ShowDialog(mParent)
    End Sub
#End Region
End Class
#Region "Support Class(es)"
Public Class ItemData
    Public Sub New(ByVal Text As String, ByVal ItemCode As Short, Optional ByVal Category As String = "", Optional ByVal Affect As String = "", Optional ByVal Value As Integer = -1, Optional ByVal UserClass As String = "", Optional ByVal Notes As String = "")
        mAffect = Affect
        mCategory = Category
        mItemCode = ItemCode
        mNotes = Notes
        mUserClass = UserClass
        mText = Text
        mValue = Value
        If mCategory = "" AndAlso mText.Contains(":") Then mCategory = mText.Substring(0, mText.IndexOf(":"))
    End Sub
#Region "Properties"
    Private mAffect As String
    Private mCategory As String
    Private mNotes As String
    Private mText As String
    Private mItemCode As Short
    Private mUserClass As String
    Private mValue As Integer
    Public Property Category() As String
        Get
            Return mAffect
        End Get
        Set(value As String)
            mAffect = value
        End Set
    End Property
    Public Property Affect() As String
        Get
            Return mCategory
        End Get
        Set(value As String)
            mCategory = value
        End Set
    End Property
    Public Property ItemCode() As Short
        Get
            Return mItemCode
        End Get
        Set(value As Short)
            mItemCode = value
        End Set
    End Property
    Public Property Notes() As String
        Get
            Return mNotes
        End Get
        Set(value As String)
            mNotes = value
        End Set
    End Property
    Public Property Text() As String
        Get
            Return mText
        End Get
        Set(value As String)
            mText = value
        End Set
    End Property
    Public Property UserClass() As String
        Get
            Return mUserClass
        End Get
        Set(value As String)
            mUserClass = value
        End Set
    End Property
    Public Property Value() As Integer
        Get
            Return mValue
        End Get
        Set(value As Integer)
            mValue = value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Overrides Function ToString() As String
        Return mText
    End Function
#End Region
End Class
#End Region