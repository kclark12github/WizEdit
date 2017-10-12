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
        For iChar As Short = 0 To CharactersMax - 1
            mCharacters(iChar) = New Character(Me)
        Next iChar
        Read()
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
            mItemList = {New Item(), New Item(), New Item(), New Item(), New Item(), New Item(), New Item(), New Item()}
            mExperience = 0
            mExperiencePacked = {0S, 0S, 0S}
            mLVL = New Points
            mHP = New Points

            mSpellBooks = {0, 0, 0, 0, 0, 0, 0, 0}
            mMageSpellPoints = {0S, 0S, 0S, 0S, 0S, 0S, 0S}
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
        Private mMageSpellPoints(SpellLevelMax - 1) As UInt16
        Private mPriestSpellPoints(SpellLevelMax - 1) As UInt16
        Private mArmorClass As UInt16
        Private mLocation As UInt16
        Private mDown As UInt16
        Private mHonors As UInt16

        Private MageSpellBook() As Spell = {
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
        Private PriestSpellBook() As Spell = {
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
                Return mHonors.ToString
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
                Return String.Format("Facing North; {0:D} East; {1:D} North; {2} Down", (mLocation \ 100), (mLocation Mod 100), mDown)
            End Get
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
                        MageSpellBook(iSpell).Known = CBool((mSpellBooks(i) And 2 ^ Offset) = 2 ^ Offset)
                    Else
                        PriestSpellBook(iSpell - 22).Known = CBool((mSpellBooks(i) And 2 ^ Offset) = 2 ^ Offset)
                    End If
                    iSpell += 1
                Next Offset
            Next i

            For i As Short = 0 To SpellLevelMax - 1                 '0x1D892
                mMageSpellPoints(i) = binReader.ReadInt16()
            Next i
            If mName = "NEB" Then Debug.WriteLine(String.Format("NEB Priest Spells @ 0x{0:X00000}", binReader.BaseStream.Position))
            'TODO: VB6 incarnation reports this as 9/9/9/9/9/9/9 while we seem to find 2/0/0/0/0/0/0/0
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
            For iSpell As Short = 1 To MageSpellBook.Length - 1
                ToString &= String.Format("{1}{2:00}) {3}{0}", New Object() {vbCrLf, vbTab, iSpell, MageSpellBook(iSpell).ToString})
            Next iSpell
            ToString &= String.Format("{0}Mage Spell Points:    {1}", vbCrLf, mMageSpellPoints(0))
            For iPoints As Integer = 1 To mMageSpellPoints.Length - 1
                ToString &= String.Format("/{0}", mMageSpellPoints(iPoints))
            Next iPoints
            ToString &= String.Format("{0}Priest SpellBook...{0}", vbCrLf)
            For iSpell As Short = 0 To PriestSpellBook.Length - 1
                ToString &= String.Format("{1}{2:00}) {3}{0}", New Object() {vbCrLf, vbTab, iSpell + 1, PriestSpellBook(iSpell).ToString})
            Next iSpell
            ToString &= String.Format("{0}Priest Spell Points:  {1}", vbCrLf, mPriestSpellPoints(0))
            For iPoints As Integer = 1 To mPriestSpellPoints.Length - 1
                ToString &= String.Format("/{0}", mPriestSpellPoints(iPoints))
            Next iPoints
        End Function
#End Region
    End Class
    Public Class Item
        Public Sub New()
            mEquipped = 0
            mCursed = 0
            mIdentified = 0
            mItemCode = 0
        End Sub
        Private mEquipped As UInt16
        Private mCursed As UInt16
        Private mIdentified As UInt16
        Private mItemCode As UInt16
        Private mMasterItemList As String() = {
                "Misc: Broken Item",
                "Weapon: Long Sword",
                "Weapon: Short Sword",
                "Weapon: Anointed Mace",
                "Weapon: Anointed Flail",
                "Weapon: Staff",
                "Weapon: Dagger",
                "Shield: Small Shield",
                "Shield: Large Shield",
                "Armor: Robes",
                "Armor: Leather Armor",
                "Armor: Chain Mail",
                "Armor: Breast Plate",
                "Armor: Plate Mail",
                "Helm: Helm",
                "Magic: Potion Of Dios",
                "Magic: Potion Of Latumofis",
                "Weapon: Long Sword +1",
                "Weapon: Short Sword +1",
                "Weapon: Mace +1",
                "Weapon: Staff of Mogref",
                "Magic: Scroll of Katino",
                "Armor: Leather +1",
                "Armor: Chain Mail +1",
                "Armor: Plate Mail +1",
                "Shield: Shield +1",
                "Armor: Breast Plate +1",
                "Magic: Scroll Of Badios",
                "Magic: Scroll Of Halito",
                "Weapon: Long Sword -1",
                "Weapon: Short Sword -1",
                "Weapon: Mace -1",
                "Weapon: Staff +2",
                "Weapon: Dragon Slayer",
                "Helm: Helm +1",
                "Armor: Leather -1",
                "Armor: Chain -1",
                "Armor: Breast Plate -1",
                "Shield: Shield -1",
                "Magic: Jeweled Amulet",
                "Magic: Scroll of Badios",
                "Magic: Potion of Sopic",
                "Weapon: Long Sword +2",
                "Weapon: Short Sword +2",
                "Weapon: Mace +2",
                "Magic: Scroll Of Lomilwa",
                "Magic: Scroll Of Dilto",
                "Gauntlets: Copper Gloves",
                "Armor: Leather +2",
                "Armor: Chain +2",
                "Armor: Plate Mail +2",
                "Shield: Shield +2",
                "Helm: Helm +2 (E)",
                "Magic: Potion Of Dial",
                "Magic: Ring of Porfic",
                "Weapon: Were Slayer",
                "Weapon: Mage Masher",
                "Weapon: Mace Pro Poison",
                "Weapon: Staff Of Montino",
                "Weapon: Blade Cusinart'",
                "Magic: Amulet Of Manifo",
                "Weapon: Rod Of Flame",
                "Armor: Chain +2 (E)",
                "Armor: Plate +2 (N)",
                "Shield: Shield +3 (E)",
                "Magic: Amulet Of Makanito",
                "Helm: Helm of Malor",
                "Magic: Scroll of Badial",
                "Weapon: Short Sword -2",
                "Weapon: Dagger +2",
                "Weapon: Mace -2",
                "Weapon: Staff -2",
                "Weapon: Dagger Of Speed",
                "Armor: Cursed Robe",
                "Armor: Leather -2",
                "Armor: Chain -2",
                "Armor: Breastplate -2",
                "Shield: Shield -2",
                "Helm: Cursed Helmet",
                "Armor: Breast Plate +2",
                "Gauntlets: Gloves of Silver",
                "Weapon: Evil +3 Sword",
                "Weapon: +3 Evil Short Sword",
                "Weapon: Thieves Dagger",
                "Armor: +3 Breast Plate",
                "Armor: Lord's Garb",
                "Weapon: Muramasa Blade",
                "Weapon: Shiriken",
                "Armor: Chain Pro Fire",
                "Armor: +3 Evil Plate",
                "Shield: +3 Shield",
                "Magic: Ring of Healing",
                "Magic: Ring Pro Undead",
                "Magic: Deadly Ring",
                "Special: Werdna's Amulet",
                "Special: Statuette/Bear",
                "Special: Statuette/Frog",
                "Special: Bronze Key",
                "Special: Silver Key",
                "Special: Gold Key",
                "Special: Blue Ribbon"}
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

            Return String.Format("{0}{1}{2}", vbTab, IIf(mCursed, "-", IIf(mEquipped, "*", " ")), mMasterItemList(mItemCode))
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
        Private mKnown As Boolean = False
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
        Public Property Known As Boolean
            Get
                Return mKnown
            End Get
            Set(value As Boolean)
                mKnown = value
            End Set
        End Property
#End Region
#Region "Methods"
        Public Overrides Function ToString() As String
            Return String.Format("{1}{0}Level {2}{0}{3} (""{4}""){0}Type: {5}{0}Affects: {6}{0}Known: {7}", New Object() {vbTab, Me.Category, Me.Level, Me.Name, Me.Translation, Me.Type, Me.Affects, IIf(Me.Known, "Yes", "No")})
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
    Const ScenarioName As String = "PROVING GROUNDS OF THE MAD OVERLORD!"
    Const ScenarioDataOffset As UInt32 = &H1D400
    Const CharacterDataOffset As UInt32 = &H1D800
    Const CharactersMax As Integer = 20
    Const ItemMapMax As Integer = 100
    Const SpellMapMax As Integer = 50

    Private mCharacters(19) As Character
    Private mItemList(ItemMapMax) As String
    Private mSpells(SpellMapMax) As String

#End Region
    Public ReadOnly Property Characters As Character()
        Get
            Return mCharacters
        End Get
    End Property
#End Region
#Region "Methods"
    Private Sub Backup()
        If Not File.Exists(MyBase.Path) Then Throw New FileNotFoundException(String.Format("{0} does not exist!", MyBase.Path))
        Dim fi As FileInfo = New FileInfo(MyBase.Path)
        Dim backup As String = fi.Name.Replace(fi.Extension, String.Format(".{0:yyyyMMdd.HHmmssff}{1}", fi.LastWriteTime, fi.Extension))
        Dim backupPath As String = String.Format("{0}\{1}", fi.DirectoryName, backup)
        If Not FileIO.FileSystem.FileExists(backupPath) Then
            FileIO.FileSystem.RenameFile(MyBase.Path, backup)
            FileIO.FileSystem.CopyFile(backupPath, MyBase.Path, FileIO.UIOption.OnlyErrorDialogs)
        End If
    End Sub
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
                mCharacters(iChar).Read(binReader)
                mCharacters(iChar + 1).Read(binReader)
                mCharacters(iChar + 2).Read(binReader)
                mCharacters(iChar + 3).Read(binReader)
                binReader.BaseStream.Position += 192
            Next iBlock
            For iChar As Short = 0 To 19
                If mCharacters(iChar).Name <> "" Then
                    Debug.WriteLine(New String("="c, 132))
                    Debug.WriteLine(String.Format("{0:00}) {1}", iChar + 1, mCharacters(iChar).Name))
                    Debug.WriteLine(New String("-"c, 132))
                    Debug.WriteLine(mCharacters(iChar).ToString)
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
                mCharacters(iChar).Save(binWriter)
                mCharacters(iChar + 1).Save(binWriter)
                mCharacters(iChar + 2).Save(binWriter)
                mCharacters(iChar + 3).Save(binWriter)
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
