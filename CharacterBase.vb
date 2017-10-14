'CharacterBase.cls
'   Character Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class CharacterBase
    Public Sub New(ByVal Base As WizEditBase)
        mBase = Base
        mName = ""
        mPassword = ""

        mOut = 0S
        mRace = 0S
        mProfession = 0S
        mAgeInWeeks = 0S
        mStatus = 0S
        mAlignment = 0S

        mStatistics = 0
        mStrength = 0
        mIntelligence = 0
        mPiety = 0
        mVitality = 0
        mAgility = 0
        mLuck = 0

        mGold = 0
        mGoldPacked = {0S, 0S, 0S}
        mItemCount = 0S
        mItemList = {New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase), New ItemBase(mBase)}
        mExperience = 0
        mExperiencePacked = {0S, 0S, 0S}
        mLVL = New PointsBase
        mHP = New PointsBase

        mSpellBooks = {0, 0, 0, 0, 0, 0, 0, 0}
        ReDim mMageSpellBook(mBase.MageSpellBook.Length - 1)
        mMageSpellPoints = {0S, 0S, 0S, 0S, 0S, 0S, 0S}
        ReDim mPriestSpellBook(mBase.PriestSpellBook.Length - 1)
        mPriestSpellPoints = {0S, 0S, 0S, 0S, 0S, 0S, 0S}
        mArmorClass = 0S
        mLocation = 0S
        mDown = 0S
        mHonors = 0S
    End Sub
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
        Priest = 2
        Thief = 3
        Bishop = 4
        Samurai = 5
        Lord = 6
        Ninja = 7
    End Enum
    Public Enum enumRace As Byte
        NoRace = 0
        Human = 1
        Elf = 2
        Dwarf = 3
        Gnome = 4
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
#Region "Properties"
#Region "Declarations"
    Const NamePasswordLengthMax As Short = 15

    Const ItemListMax As Integer = 8
    Const AlignmentMapMax As Integer = 3
    Const RaceMapMax As Integer = 5
    Const ProfessionMapMax As Integer = 7
    Const StatusMapMax As Integer = 7
    Const SpellLevelMax As Integer = 7

    Private mBase As WizEditBase
    Private mName As String
    Private mPassword As String

    Private mOut As UInt16
    Private mRace As UInt16
    Private mProfession As UInt16
    Private mAgeInWeeks As UInt16
    Private mStatus As UInt16
    Private mAlignment As UInt16

    Private mStatistics As UInt32
    Private mStrength As Short
    Private mIntelligence As Short
    Private mPiety As Short
    Private mVitality As Short
    Private mAgility As Short
    Private mLuck As Short

    Private mGold As Long
    Private mGoldPacked(2) As UInt16
    Private mItemCount As UInt16

    Private mItemList(ItemListMax - 1) As ItemBase
    Private mExperience As Long
    Private mExperiencePacked(2) As UInt16
    Private mLVL As PointsBase
    Private mHP As PointsBase
    Private mSpellBooks(7) As Byte
    Private mMageSpellBook() As Boolean
    Private mMageSpellPoints(SpellLevelMax - 1) As UInt16
    Private mPriestSpellBook() As Boolean
    Private mPriestSpellPoints(SpellLevelMax - 1) As UInt16
    Private mArmorClass As Int16
    Private mLocation As UInt16
    Private mDown As UInt16
    Private mHonors As UInt16
#End Region
    Public Property Age As Short
        Get
            Return mAgeInWeeks \ 52
        End Get
        Set(value As Short)
            mAgeInWeeks = value * 52
        End Set
    End Property
    Public Property AgeInWeeks As Short
        Get
            Return mAgeInWeeks
        End Get
        Set(value As Short)
            mAgeInWeeks = value
        End Set
    End Property
    Public Property Alignment As enumAlignment
        Get
            Return mAlignment
        End Get
        Set(value As enumAlignment)
            mAlignment = value
        End Set
    End Property
    Public Property ArmorClass As Short
        Get
            Return mArmorClass
        End Get
        Set(value As Short)
            mArmorClass = value
        End Set
    End Property
    Public Property Experience As Long
        Get
            Return mExperience
        End Get
        Set(value As Long)
            mExperience = value
            mBase.DtoI6(mExperience, mExperiencePacked)
        End Set
    End Property
    Public Property Gold As Long
        Get
            Return mGold
        End Get
        Set(value As Long)
            mGold = value
            mBase.DtoI6(mGold, mGoldPacked)
        End Set
    End Property
    Public ReadOnly Property HitPoints As String
        Get
            Return mHP.ToString
        End Get
    End Property
    Public Property Honors As Short
        Get
            Return mHonors
        End Get
        Set(value As Short)
            mHonors = value
        End Set
    End Property
    Public ReadOnly Property HonorsFull As String
        Get
            If mHonors Or &H1 Then Return "> Chevron of Trebor"
            If mHonors Or &H4000 Then Return "G - Mark of Gnilda"
            If mHonors Or &H800 Then Return "K - Knight of Gnilda"
            If mHonors Or &H2000 Then Return "D - Descendant of Heroes"
            If mHonors Or &H20 Then Return "* - Star of Llylgamyn"
            Return ""
        End Get
    End Property
    Public ReadOnly Property HonorsShort As String
        Get
            If mHonors Or &H1 Then Return ">"
            If mHonors Or &H4000 Then Return "G"
            If mHonors Or &H800 Then Return "K"
            If mHonors Or &H2000 Then Return "D"
            If mHonors Or &H20 Then Return "*"
            Return ""
        End Get
    End Property
    Public ReadOnly Property HP As PointsBase
        Get
            Return mHP
        End Get
    End Property
    Public Property ItemCount As Int16
        Get
            Return mItemCount
        End Get
        Set(value As Int16)
            mItemCount = value
        End Set
    End Property
    Public ReadOnly Property Items As ItemBase()
        Get
            Return mItemList
        End Get
    End Property
    Public ReadOnly Property IsBishop() As Boolean
        Get
            Return CBool(mProfession = enumProfession.Bishop)
        End Get
    End Property
    Public ReadOnly Property IsMage() As Boolean
        Get
            Select Case mProfession
                Case enumProfession.Bishop, enumProfession.Mage, enumProfession.Samurai : Return True
            End Select
            Return False
        End Get
    End Property
    Public ReadOnly Property IsPriest() As Boolean
        Get
            Select Case mProfession
                Case enumProfession.Priest, enumProfession.Bishop, enumProfession.Lord : Return True
            End Select
            Return False
        End Get
    End Property
    Public ReadOnly Property IsLord() As Boolean
        Get
            Return CBool(mProfession = enumProfession.Lord)
        End Get
    End Property
    Public ReadOnly Property IsSamurai() As Boolean
        Get
            Return CBool(mProfession = enumProfession.Samurai)
        End Get
    End Property
    Public ReadOnly Property IsSpellCaster() As Boolean
        Get
            Select Case mProfession
                Case enumProfession.Mage, enumProfession.Priest, enumProfession.Bishop, enumProfession.Lord, enumProfession.Samurai : Return True
            End Select
            Return False
        End Get
    End Property
    Public ReadOnly Property Level As String
        Get
            Return mLVL.ToString
        End Get
    End Property
    Public ReadOnly Property Location As String
        Get
            If mLocation = 0 Then Return "Castle"
            Return String.Format("Facing North; {0:D} East; {1:D} North; {2} Down", (mLocation \ 100), (mLocation Mod 100), mDown)
        End Get
    End Property
    Public Property LocationDown As Int16
        Get
            Return mDown
        End Get
        Set(value As Int16)
            mDown = value
        End Set
    End Property
    Public Property LocationEast As UInt16
        Get
            Return (mLocation \ 100)
        End Get
        Set(value As UInt16)
            Dim north As UInt16 = Me.LocationNorth
            mLocation = (value * 100) + north
        End Set
    End Property
    Public Property LocationNorth As UInt16
        Get
            Return (mLocation Mod 100)
        End Get
        Set(value As UInt16)
            Dim east As UInt16 = Me.LocationEast
            mLocation = (east * 100) + value
        End Set
    End Property
    Public ReadOnly Property LocationFull As String
        Get
            Return String.Format("{0} (from the steps leading to the castle)", Me.Location)
        End Get
    End Property
    Public ReadOnly Property LVL As PointsBase
        Get
            Return mLVL
        End Get
    End Property
    Public ReadOnly Property MageSpellBook As Boolean()
        Get
            Return mMageSpellBook
        End Get
    End Property
    Public Property MageSpellPoints(ByVal Level As Short) As Int16
        Get
            Return mMageSpellPoints(Level)
        End Get
        Set(value As Int16)
            mMageSpellPoints(Level) = value
        End Set
    End Property
    Public Property Name As String
        Get
            Return mName
        End Get
        Set(value As String)
            If value.Length > NamePasswordLengthMax Then Throw New ArgumentException(String.Format("Name length is limited to {0} characters!", NamePasswordLengthMax))
            mName = value
        End Set
    End Property
    Public Property Out As Boolean
        Get
            Return CBool(mOut = 1)
        End Get
        Set(value As Boolean)
            mOut = IIf(value, 1, 0)
        End Set
    End Property
    Public Property Password As String
        Get
            Return mPassword
        End Get
        Set(value As String)
            If value.Length > NamePasswordLengthMax Then Throw New ArgumentException(String.Format("Password length is limited to {0} characters!", NamePasswordLengthMax))
            mPassword = value
        End Set
    End Property
    Public ReadOnly Property PriestSpellBook As Boolean()
        Get
            Return mPriestSpellBook
        End Get
    End Property
    Public Property PriestSpellPoints(ByVal Level As Short) As Int16
        Get
            Return mPriestSpellPoints(Level)
        End Get
        Set(value As Int16)
            mPriestSpellPoints(Level) = value
        End Set
    End Property
    Public Property Profession As enumProfession
        Get
            Return mProfession
        End Get
        Set(value As enumProfession)
            mProfession = value
        End Set
    End Property
    Public Property Race As enumRace
        Get
            Return mRace
        End Get
        Set(value As enumRace)
            mRace = value
        End Set
    End Property
    Public Property StatusCode As enumStatus
        Get
            Return mStatus
        End Get
        Set(value As enumStatus)
            mStatus = value
        End Set
    End Property
    Public ReadOnly Property Status As String
        Get
            Select Case Me.StatusCode
                Case enumStatus.LostDeleted : Return "Lost/Deleted"
                Case Else : Return String.Format("{0}", Me.StatusCode)
            End Select
        End Get
    End Property
    Public ReadOnly Property Tag As String
        Get
            Return String.Format("{0}{1} L {2} {3}-{4} {5}", New Object() {Me.Name.ToUpper, New String(" "c, NamePasswordLengthMax - Me.Name.Length), Me.LVL.Current, Me.Alignment.ToString.Substring(0, 1).ToUpper, Me.Profession.ToString.Substring(0, 3).ToUpper, Me.Race.ToString.ToUpper})
        End Get
    End Property

    Public Property Statistics As UInt32
        Get
            Return mStatistics
        End Get
        Set(value As UInt32)
            mStatistics = value
            mStrength = ((mStatistics \ (2 ^ 0)) And &H1F)
            mIntelligence = ((mStatistics \ (2 ^ 5)) And &H1F)
            mPiety = ((mStatistics \ (2 ^ 10)) And &H1F)
            mVitality = ((mStatistics \ (2 ^ 16)) And &H1F)
            mAgility = ((mStatistics \ (2 ^ 21)) And &H1F)
            mLuck = ((mStatistics \ (2 ^ 26)) And &H1F)
        End Set
    End Property
    Public Property Strength As UInt16
        Get
            Return mStrength
        End Get
        Set(value As UInt16)
            mStrength = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
    Public Property Intelligence As UInt16
        Get
            Return mIntelligence
        End Get
        Set(value As UInt16)
            mIntelligence = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
    Public Property Piety As UInt16
        Get
            Return mPiety
        End Get
        Set(value As UInt16)
            mPiety = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
    Public Property Vitality As UInt16
        Get
            Return mVitality
        End Get
        Set(value As UInt16)
            mVitality = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
    Public Property Agility As UInt16
        Get
            Return mAgility
        End Get
        Set(value As UInt16)
            mAgility = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
    Public Property Luck As UInt16
        Get
            Return mLuck
        End Get
        Set(value As UInt16)
            mLuck = value
            mStatistics = (mStrength * (2 ^ 0)) + (mIntelligence * (2 ^ 5)) + (mPiety * (2 ^ 10)) + (mVitality * (2 ^ 16)) + (mAgility * (2 ^ 21)) + (mLuck * (2 ^ 26))
        End Set
    End Property
#End Region
#Region "Methods"
    Public Sub Read(binReader As BinaryReader)
        'Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", binReader.BaseStream.Position))
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

        Me.Statistics = binReader.ReadInt32()                   '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
        binReader.BaseStream.Position += 4                      '0x1D830

        For i As Short = 0 To 2                                 '0x1D834
            mGoldPacked(i) = binReader.ReadInt16()
        Next i
        mGold = mBase.I6toD(mGoldPacked)

        mItemCount = binReader.ReadInt16()                      '0x1D83A
        For i As Short = 0 To ItemListMax - 1                   '0x1D83C    List of Items (stowing not an option in Wiz01...)
            mItemList(i).Read(binReader)
        Next i

        For i As Short = 0 To 2                                 '0x1D87C
            mExperiencePacked(i) = binReader.ReadInt16()
        Next i
        mExperience = mBase.I6toD(mExperiencePacked)

        mLVL.Read(binReader)                                    '0x1D882
        mHP.Read(binReader)                                     '0x1D886

        For i As Short = 0 To 7                                 '0x1D88A    Need to mask as bits...
            mSpellBooks(i) = binReader.ReadByte()
        Next i
        Dim iSpell As Short = 0
        For i As Short = 0 To mSpellBooks.Length - 1
            For Offset As Short = 0 To 7
                If iSpell > 50 Then Exit For
                If iSpell <= 21 Then
                    mMageSpellBook(iSpell) = CBool((mSpellBooks(i) And 2 ^ Offset) = 2 ^ Offset)
                Else
                    mPriestSpellBook(iSpell - 22) = CBool((mSpellBooks(i) And 2 ^ Offset) = 2 ^ Offset)
                End If
                iSpell += 1
            Next Offset
        Next i

        For i As Short = 0 To SpellLevelMax - 1                 '0x1D892
            mMageSpellPoints(i) = binReader.ReadInt16()
        Next i
        For i As Short = 0 To SpellLevelMax - 1                 '0x1D8A0
            mPriestSpellPoints(i) = binReader.ReadInt16()
        Next i
        binReader.BaseStream.Position += 2                      '0x1D8AE
        mArmorClass = binReader.ReadInt16()                     '0x1D8B0
        binReader.BaseStream.Position += 24                     '0x1D8B2
        mLocation = binReader.ReadInt16()                       '0x1D8CA    Some sort of packed variable...
        mDown = binReader.ReadInt16()                           '0x1D8CC    Seems to be a simple 2-byte Int16...
        mHonors = binReader.ReadInt16()                         '0x1D8CE    Need more testing, but 1 = ">"
        '                                                       '0x1D8D0    Next Character Record...
    End Sub
    Public Sub Save(binWriter As BinaryWriter)
        binWriter.Write(mName)                                  '0x1D800   Pascal Varying Length String Format...
        binWriter.BaseStream.Position += NamePasswordLengthMax - mName.Length
        binWriter.Write(mPassword)                              '0x1D810   Pascal Varying Length String Format...
        binWriter.BaseStream.Position += NamePasswordLengthMax - mPassword.Length

        binWriter.Write(mOut)                                   '0x1D820    00 00 = No; 01 00 = Yes;
        binWriter.Write(mRace)                                  '0x1D822    01 00 = Human
        binWriter.Write(mProfession)                            '0x1D824    06 00 = Lord
        binWriter.Write(mAgeInWeeks)                            '0x1D826    0C 03 = Weeks Alive...
        binWriter.Write(mStatus)                                '0x1D828    00 00 = OK
        binWriter.Write(mAlignment)                             '0x1D82A    02 00 = Neutral

        binWriter.Write(Me.Statistics)                          '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
        binWriter.BaseStream.Position += 4                      '0x1D830

        For i As Short = 0 To 2                                 '0x1D834
            binWriter.Write(mGoldPacked(i))
        Next i

        binWriter.Write(mItemCount)                             '0x1D83A
        For i As Short = 0 To ItemListMax - 1                   '0x1D83C    List of Items (stowing not an option in Wiz01...)
            mItemList(i).Save(binWriter)
        Next i

        For i As Short = 0 To 2                                 '0x1D87C
            binWriter.Write(mExperiencePacked(i))
        Next i

        mLVL.Save(binWriter)                                    '0x1D882
        mHP.Save(binWriter)                                     '0x1D886

        For i As Short = 0 To 7                                 '0x1D88A    Need to mask as bits...
            binWriter.Write(mSpellBooks(i))
        Next i

        For i As Short = 0 To SpellLevelMax - 1                 '0x1D892
            binWriter.Write(mMageSpellPoints(i))
        Next i
        If mName = "NEB" Then Debug.WriteLine(String.Format("NEB Priest Spells @ 0x{0:X00000}", binWriter.BaseStream.Position))
        'TODO: VB6 incarnation reports this as 9/9/9/9/9/9/9 while we seem to find 2/0/0/0/0/0/0/0
        For i As Short = 0 To SpellLevelMax - 1                 '0x1D8A0
            binWriter.Write(mPriestSpellPoints(i))
        Next i
        binWriter.BaseStream.Position += 2                      '0x1D8AE
        binWriter.Write(mArmorClass)                            '0x1D8B0
        binWriter.BaseStream.Position += 24                     '0x1D8B2
        binWriter.Write(mLocation)                              '0x1D8CA    Some sort of packed variable...
        binWriter.Write(mDown)                                  '0x1D8CC    Seems to be a simple 2-byte Int16...
        binWriter.Write(mHonors)                                '0x1D8CE    Need more testing, but 1 = ">"
    End Sub
    Public Overrides Function ToString() As String
        ToString = String.Format("Name:               {1}{2}{0}", vbCrLf, Me.Name, Me.HonorsShort)
        ToString &= String.Format("Password:           {1}{0}", vbCrLf, Me.Password)
        ToString &= String.Format("On Expedition:      {1}{0}", vbCrLf, IIf(Me.Out, "YES", "NO"))
        ToString &= String.Format("Location:           {1}{0}", vbCrLf, Me.LocationFull)
        ToString &= String.Format("Race:               {1}{0}", vbCrLf, Me.Race)
        ToString &= String.Format("Profession:         {1}{0}", vbCrLf, Me.Profession)
        ToString &= String.Format("Age:                {1} ({2} weeks){0}", vbCrLf, Me.Age, Me.AgeInWeeks)
        ToString &= String.Format("Status:             {1}{0}", vbCrLf, Me.Status)
        ToString &= String.Format("Alignment:          {1}{0}", vbCrLf, Me.Alignment)
        ToString &= String.Format("Honors:             {1}{0}", vbCrLf, Me.HonorsFull)
        ToString &= String.Format("Level:              {1}{0}", vbCrLf, Me.Level)
        ToString &= String.Format("Hit Points:         {1}{0}", vbCrLf, Me.HitPoints)
        ToString &= String.Format("Gold Pieces:        {1}{0}", vbCrLf, Me.Gold)
        ToString &= String.Format("Experience Points:  {1}{0}", vbCrLf, Me.Experience)
        ToString &= String.Format("Armor Class:        {1}{0}", vbCrLf, Me.ArmorClass)

        ToString &= String.Format("{0}Basic Statistics...{0}", vbCrLf)
        ToString &= String.Format("   Strength:        {1}{0}", vbCrLf, mStrength)
        ToString &= String.Format("   Intelligence:    {1}{0}", vbCrLf, mIntelligence)
        ToString &= String.Format("   Piety:           {1}{0}", vbCrLf, mPiety)
        ToString &= String.Format("   Vitality:        {1}{0}", vbCrLf, mVitality)
        ToString &= String.Format("   Agility:         {1}{0}", vbCrLf, mAgility)
        ToString &= String.Format("   Luck:            {1}{0}", vbCrLf, mLuck)

        ToString &= String.Format("{0}List of Items (Currently carrying {1} items)...{0}", vbCrLf, Me.Items.Length)
        For iItem As Integer = 0 To Me.Items.Length - 1
            ToString &= String.Format("{1}{2:D}) {3}{0}", New Object() {vbCrLf, vbTab, iItem, Me.Items(iItem).ToString})
        Next iItem

        'SpellBooks...
        ToString &= String.Format("{0}Mage SpellBook...{0}", vbCrLf)
        For iSpell As Short = 1 To mMageSpellBook.Length - 1
            ToString &= String.Format("{1}{2:00}) {3}; Known: {4}{0}", New Object() {vbCrLf, vbTab, iSpell, mBase.MageSpellBook(iSpell).ToString, IIf(mMageSpellBook(iSpell), "Yes", "No")})
        Next iSpell
        ToString &= String.Format("{0}Mage Spell Points:    {1}", vbCrLf, mMageSpellPoints(0))
        For iPoints As Integer = 1 To mMageSpellPoints.Length - 1
            ToString &= String.Format("/{0}", mMageSpellPoints(iPoints))
        Next iPoints
        ToString &= String.Format("{0}Priest SpellBook...{0}", vbCrLf)
        For iSpell As Short = 0 To mPriestSpellBook.Length - 1
            ToString &= String.Format("{1}{2:00}) {3}; Known: {4}{0}", New Object() {vbCrLf, vbTab, iSpell + 1, mBase.PriestSpellBook(iSpell).ToString, IIf(mPriestSpellBook(iSpell), "Yes", "No")})
        Next iSpell
        ToString &= String.Format("{0}Priest Spell Points:  {1}", vbCrLf, mPriestSpellPoints(0))
        For iPoints As Integer = 1 To mPriestSpellPoints.Length - 1
            ToString &= String.Format("/{0}", mPriestSpellPoints(iPoints))
        Next iPoints
    End Function
#End Region
End Class