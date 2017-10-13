﻿'WizEditBase.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/11/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class WizEditBase
    Public Sub New()
        mPath = Nothing
        mFileInfo = Nothing
        mBoxArt = Nothing
        mCaption = Nothing
        mIcon = Nothing
        mParent = Nothing
    End Sub
    Public Sub New(Path As String, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        mPath = Path
        If Path IsNot Nothing AndAlso Path <> "" Then mFileInfo = New FileInfo(Path)

        mBoxArt = BoxArt
        mCaption = Caption
        mIcon = Icon
        mParent = Parent
        ReDim mCharacters(Me.CharactersMax - 1)
        For iChar As Short = 0 To Me.CharactersMax - 1
            mCharacters(iChar) = New Character(Me)
        Next iChar
    End Sub
#Region "#Internal Class(es)"
    Public Class Character
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
            mItemList = {New Item(mBase), New Item(mBase), New Item(mBase), New Item(mBase), New Item(mBase), New Item(mBase), New Item(mBase), New Item(mBase)}
            mExperience = 0
            mExperiencePacked = {0S, 0S, 0S}
            mLVL = New Points
            mHP = New Points

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

        Private mItemList(ItemListMax - 1) As Item
        Private mExperience As Long
        Private mExperiencePacked(2) As UInt16
        Private mLVL As Points
        Private mHP As Points
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
        Public ReadOnly Property HP As Points
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
        Public ReadOnly Property Items As Item()
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
        Public ReadOnly Property LVL As Points
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
    Public Class Item
        Public Sub New(ByVal Base As WizEditBase)
            mBase = Base
            mEquipped = 0
            mCursed = 0
            mIdentified = 0
            mItemCode = 0
        End Sub
        Private mBase As WizEditBase
        Private mEquipped As Int16
        Private mCursed As Int16
        Private mIdentified As Int16
        Private mItemCode As Int16
        Public Property ItemCode As Short
            Get
                Return mItemCode
            End Get
            Set(value As Short)
                mItemCode = value
            End Set
        End Property
        Public Property Cursed As Boolean
            Get
                Return CBool(mCursed <> 0)
            End Get
            Set(value As Boolean)
                mCursed = IIf(value, 1, 0)
            End Set
        End Property
        Public Property Equipped As Boolean
            Get
                Return CBool(mEquipped <> 0)
            End Get
            Set(value As Boolean)
                mEquipped = IIf(value, 1, 0)
            End Set
        End Property
        Public Property Identified As Boolean
            Get
                Return CBool(mIdentified <> 0)
            End Get
            Set(value As Boolean)
                mIdentified = IIf(value, 1, 0)
            End Set
        End Property
        Public Sub Read(binReader As BinaryReader)
            mEquipped = binReader.ReadInt16()
            mCursed = binReader.ReadInt16()
            mIdentified = binReader.ReadInt16()
            mItemCode = binReader.ReadInt16()
        End Sub
        Public Sub Save(binWriter As BinaryWriter)
            binWriter.Write(mEquipped)
            binWriter.Write(mCursed)
            binWriter.Write(mIdentified)
            binWriter.Write(mItemCode)
        End Sub
        Public Overrides Function ToString() As String
            '    Item = vbTab & ItemList(x.ItemCode) & "; Code: " & x.ItemCode & "; Equipped: "
            '    If x.Identified Then Item &= "; Identified"
            '    If x.Equipped Then Item &= "; **EQUIPPED**"
            '    If x.Cursed Then strItem &= "; --CURSED--"

            Return String.Format("{0}{1}{2}", vbTab, IIf(mCursed, "-", IIf(mEquipped, "*", " ")), mBase.MasterItemList(mItemCode))
        End Function
    End Class
    Public Class Points
        Public Sub New()
            mCurrent = 0
            mMaximum = 0
        End Sub
        Private mCurrent As UInt16 = 0
        Private mMaximum As UInt16 = 0
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
        Public Sub Read(binReader As BinaryReader)
            mCurrent = binReader.ReadInt16()
            mMaximum = binReader.ReadInt16()
        End Sub
        Public Sub Save(binWriter As BinaryWriter)
            binWriter.Write(mCurrent)
            binWriter.Write(mMaximum)
        End Sub
        Public Overrides Function ToString() As String
            Return String.Format("{0}/{1}", mCurrent, mMaximum)
        End Function
    End Class
    Public Class Spell
        Public Sub New(ByVal Name As String, ByVal Translation As String, ByVal Type As enumSpellType, ByVal Affects As enumSpellAffects, ByVal Category As enumSpellCategory, ByVal Level As Short)
            mName = Name
            mTranslation = Translation
            mType = Type
            mAffects = Affects
            mCategory = Category
            mLevel = Level
        End Sub
#Region "Properties"
#Region "Enumerations"
        Public Enum enumSpellCategory As Short
            Mage
            Priest
        End Enum
        Public Enum enumSpellType As Short
            AnyTime
            Camp
            Combat
            Looting
        End Enum
        Public Enum enumSpellAffects As Short
            Caster
            OnePerson
            EntireParty
            OneMonster
            OneGroup
            AllMonsters
            Variable
        End Enum
#End Region
#Region "Declarations"
        Private mName As String
        Private mTranslation As String
        Private mType As enumSpellType
        Private mAffects As enumSpellAffects
        Private mCategory As enumSpellCategory
        Private mLevel As Short
#End Region
        Public ReadOnly Property Name As String
            Get
                Return mName
            End Get
        End Property
        Public ReadOnly Property Translation As String
            Get
                Return mTranslation
            End Get
        End Property
        Public ReadOnly Property Type As String
            Get
                Return String.Format("{0}", mType)
            End Get
        End Property
        Public ReadOnly Property Affects As String
            Get
                Select Case mAffects
                    Case enumSpellAffects.AllMonsters : Return "All Monsters"
                    Case enumSpellAffects.Caster : Return "Caster"
                    Case enumSpellAffects.EntireParty : Return "Entire Party"
                    Case enumSpellAffects.OnePerson : Return "1 Person"
                    Case enumSpellAffects.OneMonster : Return "1 Monster"
                    Case enumSpellAffects.OneGroup : Return "1 Group"
                    Case enumSpellAffects.Variable : Return "Variable"
                End Select
                Return "Unknown"
            End Get
        End Property
        Public ReadOnly Property Category As enumSpellCategory
            Get
                Return mCategory
            End Get
        End Property
        Public ReadOnly Property Level As Short
            Get
                Return mLevel
            End Get
        End Property
#End Region
#Region "Methods"
        Public Overrides Function ToString() As String
            Return String.Format("{1}{0}Level {2}{0}{3} (""{4}""){0}Type: {5}{0}Affects: {6}", New Object() {vbTab, Me.Category, Me.Level, Me.Name, Me.Translation, Me.Type, Me.Affects})
        End Function
#End Region
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
#Region "Declarations"
    Private mBoxArt As Image
    Private mCaption As String
    Private mForm As Form
    Private mIcon As Icon
    Private mParent As Form

    Private mCharacters() As Character
    Private mFileInfo As FileInfo = Nothing
    Protected mMasterItemList As ItemData()
    Protected mMageSpellBook As Spell()
    Protected mPriestSpellBook As Spell()
    Private mPath As String = vbNullString
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
    Public ReadOnly Property Characters As Character()
        Get
            Return mCharacters
        End Get
    End Property
    Public Overridable ReadOnly Property CharactersMax As Short
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public ReadOnly Property DirectoryName As String
        Get
            If mFileInfo IsNot Nothing Then Return mFileInfo.DirectoryName
            Return vbNullString
        End Get
    End Property
    Public ReadOnly Property FileName As String
        Get
            If mFileInfo IsNot Nothing Then Return mFileInfo.Name
            Return vbNullString
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
    Public Overridable ReadOnly Property MageSpellBook As Spell()
        Get
            If mMageSpellBook Is Nothing Then
                mMageSpellBook = {
                    New Spell("Unknown", "Unknown", Spell.enumSpellType.Camp, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Mage, 0),
                    New Spell("Halito", "Little Fire", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Mage, 1),
                    New Spell("Mogref", "Body Iron", Spell.enumSpellType.Combat, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Mage, 1),
                    New Spell("Katino", "Bad Air", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 1),
                    New Spell("Dumapic", "Clarity", Spell.enumSpellType.Camp, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Mage, 1),
                    New Spell("Dilto", "Darkness", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 2),
                    New Spell("Sopic", "Glass", Spell.enumSpellType.Combat, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Mage, 2),
                    New Spell("Mahalito", "Big Fire", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 3),
                    New Spell("Molito", "Sparks", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 3),
                    New Spell("Morlis", "Fear", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 4),
                    New Spell("Dalto", "Blizzard", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 4),
                    New Spell("Lahalito", "Torch", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 4),
                    New Spell("Mamorlis", "Terror", Spell.enumSpellType.Combat, Spell.enumSpellAffects.AllMonsters, Spell.enumSpellCategory.Mage, 5),
                    New Spell("Makanito", "Deadly Air", Spell.enumSpellType.Combat, Spell.enumSpellAffects.AllMonsters, Spell.enumSpellCategory.Mage, 5),
                    New Spell("Madalto", "Frost King", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 5),
                    New Spell("Lakanito", "Vacuum", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Mage, 6),
                    New Spell("Zilwan", "Dispell", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Mage, 6),
                    New Spell("Masopic", "Crystal", Spell.enumSpellType.Combat, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Mage, 6),
                    New Spell("Haman", "Beg", Spell.enumSpellType.Combat, Spell.enumSpellAffects.Variable, Spell.enumSpellCategory.Mage, 6),
                    New Spell("Malor", "Teleport", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Mage, 7),
                    New Spell("Mahaman", "Beseech", Spell.enumSpellType.Combat, Spell.enumSpellAffects.Variable, Spell.enumSpellCategory.Mage, 7),
                    New Spell("Tiltowait", "Ka-Blam!", Spell.enumSpellType.Combat, Spell.enumSpellAffects.AllMonsters, Spell.enumSpellCategory.Mage, 7)
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
    Public ReadOnly Property Path As String
        Get
            Return mPath
        End Get
    End Property
    Public Overridable ReadOnly Property PriestSpellBook As Spell()
        Get
            If mPriestSpellBook Is Nothing Then
                mPriestSpellBook = {
                    New Spell("Kalki", "Blessings", Spell.enumSpellType.Combat, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 1),
                    New Spell("Dios", "Heal", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 1),
                    New Spell("Badios", "Harm", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Priest, 1),
                    New Spell("Milwa", "Light", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 1),
                    New Spell("Porfic", "Shield", Spell.enumSpellType.Combat, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Priest, 1),
                    New Spell("Matu", "Zeal", Spell.enumSpellType.Combat, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 2),
                    New Spell("Calfo", "X-Ray", Spell.enumSpellType.Looting, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Priest, 2),
                    New Spell("Manifo", "Statue", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Priest, 2),
                    New Spell("Montino", "Still Air", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Priest, 2),
                    New Spell("Lomilwa", "Sunbeam", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 3),
                    New Spell("Dialko", "Softness", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 3),
                    New Spell("Latumapic", "Identify", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 3),
                    New Spell("Bamatu", "Prayer", Spell.enumSpellType.Combat, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 3),
                    New Spell("Dial", "Cure", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 4),
                    New Spell("Badial", "Wound", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Priest, 4),
                    New Spell("Latumofis", "Cleanse", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 4),
                    New Spell("Maporfic", "Big Shield", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 4),
                    New Spell("Dialma", "Big Cure", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Badialma", "Big Wound", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Litokan", "Flames", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Kandi", "Location", Spell.enumSpellType.Camp, Spell.enumSpellAffects.Caster, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Di", "Life", Spell.enumSpellType.Camp, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Badi", "Death", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Priest, 5),
                    New Spell("Lorto", "Blades", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneGroup, Spell.enumSpellCategory.Priest, 6),
                    New Spell("Madi", "Restore", Spell.enumSpellType.AnyTime, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 6),
                    New Spell("Mabadi", "Maiming", Spell.enumSpellType.Combat, Spell.enumSpellAffects.OneMonster, Spell.enumSpellCategory.Priest, 6),
                    New Spell("Loktofeit", "Recall", Spell.enumSpellType.Combat, Spell.enumSpellAffects.EntireParty, Spell.enumSpellCategory.Priest, 6),
                    New Spell("Malikto", "Wrath", Spell.enumSpellType.Combat, Spell.enumSpellAffects.AllMonsters, Spell.enumSpellCategory.Priest, 7),
                    New Spell("Kadorto", "Rebirth", Spell.enumSpellType.Camp, Spell.enumSpellAffects.OnePerson, Spell.enumSpellCategory.Priest, 7)
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
    Public Overridable ReadOnly Property CharacterDataOffset As Int32
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property RegDataDirectory As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property RegDataFile As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
    End Property
    Public Overridable ReadOnly Property ScenarioDataOffset As String
        Get
            Throw New NotSupportedException("Must Override Property!")
        End Get
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
        If Not File.Exists(mPath) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", mPath))
        Dim fi As FileInfo = New FileInfo(mPath)
        Dim backup As String = fi.Name.Replace(fi.Extension, String.Format(".{0:yyyyMMdd.HHmmssff}{1}", fi.LastWriteTime, fi.Extension))
        Dim backupPath As String = String.Format("{0}\{1}", fi.DirectoryName, backup)
        If Not FileIO.FileSystem.FileExists(backupPath) Then
            FileIO.FileSystem.RenameFile(mPath, backup)
            FileIO.FileSystem.CopyFile(backupPath, mPath, FileIO.UIOption.OnlyErrorDialogs)
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
    Public Function GetCharacter(ByVal Tag As String) As Character
        For iChar As Short = 0 To mCharacters.Length - 1
            If mCharacters(iChar).Tag = Tag Then Return mCharacters(iChar)
        Next iChar
        Return Nothing
    End Function
    Public Sub Read()
        Dim binReader As BinaryReader = Nothing
        Try
            If Not File.Exists(mPath) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", mPath))
            binReader = New BinaryReader(File.Open(mPath, FileMode.Open))
            binReader.BaseStream.Position = Me.ScenarioDataOffset
            Dim myScenarioName As String = binReader.ReadString()
            If myScenarioName <> Me.ScenarioName Then Throw New NotSupportedException(String.Format("Save game file specified is not a valid Ultimate Wizardry Archives: {0} save game file.", Me.ScenarioName))

            'Wizardry (1-5) supports up to 20 characters...
            'The layout is a little funky in that the Character structure seems
            'not to have lined-up evenly with the disk layout (1024 byte blocks
            'on disk)... So, characters start at offset 0x0001D800 and then are
            'stored in blocks of 4 characters (832 bytes), 192 bytes of filler
            '(completing the 1K disk block), then another 4 blocks... for 5
            'total blocks of 4, making 20 characters.
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
            Me.SaveRegSetting("Environment", Me.RegDataDirectory, Me.DirectoryName)
            Me.SaveRegSetting("Environment", Me.RegDataFile, Me.FileName)
        Finally
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
    Public Sub Save()
        Dim binWriter As BinaryWriter = Nothing
        Try
            Backup()
            binWriter = New BinaryWriter(File.Open(mPath, FileMode.Open, FileAccess.Write, FileShare.None))
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
    Public Sub New(Text As String, Data As Object)
        mText = Text
        mData = Data
    End Sub
#Region "Properties"
    Private mText As String = ""
    Private mData As Object = Nothing
    Public Property Text() As String
        Get
            Return mText
        End Get
        Set(value As String)
            mText = value
        End Set
    End Property
    Public Property Data() As String
        Get
            Return mData
        End Get
        Set(value As String)
            mData = value
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