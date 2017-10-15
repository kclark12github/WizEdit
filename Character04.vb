'Character04.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class Character04
    Inherits CharacterBase
    Public Sub New(ByVal Base As Wizardry04)
        MyBase.New(Base)
    End Sub
    Protected Const SummonedMonsterGroupsMax As Integer = 3
    Protected mSummonedMonsterGroupCode(SummonedMonsterGroupsMax) As UInt16
    Protected mSummonedMonsterGroupCount(SummonedMonsterGroupsMax) As UInt16
    Protected mSummonedMonsterGroupName(SummonedMonsterGroupsMax) As String
    Public Property SummonedMonsterCode(ByVal iGroup As Short) As UInt16
        Get
            Return mSummonedMonsterGroupCode(iGroup)
        End Get
        Set(value As UInt16)
            mSummonedMonsterGroupCode(iGroup) = value
        End Set
    End Property
    Public Property SummonedMonsterCount(ByVal iGroup As Short) As UInt16
        Get
            Return mSummonedMonsterGroupCount(iGroup)
        End Get
        Set(value As UInt16)
            mSummonedMonsterGroupCount(iGroup) = value
        End Set
    End Property
    Public Property SummonedMonsterName(ByVal iGroup As Short) As String
        Get
            Return mSummonedMonsterGroupName(iGroup)
        End Get
        Set(value As String)
            mSummonedMonsterGroupName(iGroup) = value
        End Set
    End Property
    Public Overrides Sub Read(binReader As BinaryReader)
        Dim initialOffset As Long = binReader.BaseStream.Position
        'Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", binReader.BaseStream.Position))
        Try
            mName = binReader.ReadString()                          '0x4AC00   Pascal Varying Length String Format...
            If mName <> "WERDNA" Then Throw New Exception(String.Format("Unexpected data found @ initial Save Game offset 0x{0:X00000} - Expected ""WERDNA"" but found ""{1}""", initialOffset, mName))
            binReader.BaseStream.Position += NamePasswordLengthMax - mName.Length
            mPassword = binReader.ReadString()                      '0x4AC10   Pascal Varying Length String Format...
            binReader.BaseStream.Position += NamePasswordLengthMax - mPassword.Length

            mOut = binReader.ReadUInt16()                            '0x4AC20    00 00 = No; 01 00 = Yes;    'TODO: Confirm
            mRace = binReader.ReadUInt16()                           '0x4AC22    01 00 = Human
            mProfession = binReader.ReadUInt16()                     '0x4AC24    01 00 = Mage
            mAgeInWeeks = binReader.ReadUInt16()                     '0x4AC26    50 14 = 5200 Weeks (100 years) Alive...
            mStatus = binReader.ReadUInt16()                         '0x4AC28    00 00 = OK
            mAlignment = binReader.ReadUInt16()                      '0x4AC2A    03 00 = Evil

            Me.Statistics = binReader.ReadUInt32()                   '0x4AC2C    94 52 94 52 = 20/20/20/20/20/20
            binReader.BaseStream.Position += 4                      '0x4AC30

            For i As Short = 0 To 2                                 '0x4AC34
                mGoldPacked(i) = binReader.ReadUInt16()
            Next i
            mGold = mBase.I6toD(mGoldPacked)

            mItemCount = binReader.ReadUInt16()                      '0x4AC3A
            For i As Short = 0 To ItemListMax - 1                   '0x4AC3C    List of Items (stowing not an option in Wiz01...)
                '	                                                    0x4AC42 ItemCode1[I2];
                '	                                                    0x4AC4A ItemCode2[I2];	
                '	                                                    0x4AC52 ItemCode3[I2];	
                '	                                                    0x4AC5A ItemCode4[I2];	
                '	                                                    0x4AC62 ItemCode5[I2];
                '	                                                    0x4AC6A ItemCode6[I2];	
                '	                                                    0x4AC72 ItemCode7[I2];	
                '	                                                    0x4AC7A ItemCode8[I2];	
                mItemList(i).Read(binReader)
            Next i

            For i As Short = 0 To 2                                 '0x4AC7C (Keystrokes)
                mExperiencePacked(i) = binReader.ReadUInt16()
            Next i
            mExperience = mBase.I6toD(mExperiencePacked)

            mLVL.Read(binReader)                                    '0x4AC82
            mHP.Read(binReader)                                     '0x4AC86

            For i As Short = 0 To 7                                 '0x4AC8A    Need to mask as bits...
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

            For i As Short = 0 To SpellLevelMax - 1                 '0x4AC92-9F
                mMageSpellPoints(i) = binReader.ReadUInt16()
            Next i
            For i As Short = 0 To SpellLevelMax - 1                 '0x4ACA0-AD
                mPriestSpellPoints(i) = binReader.ReadUInt16()
            Next i
            '0x4AC
            binReader.BaseStream.Position += 2                      '0x4ACAE
            mArmorClass = binReader.ReadUInt16()                    '0x4ACB0
            binReader.BaseStream.Position += 24                     '0x4ACB2-C9

            mLocation = binReader.ReadUInt16()                      '0x4ACCA
            mDown = binReader.ReadUInt16()                          '0x4ACCC    Screen says -2 but data says 0E 00
            mHonors = binReader.ReadUInt16()                        '0x4ACCE    Need more testing, but 1 = ">"

            mSummonedMonsterGroupCount(0) = binReader.ReadUInt16()   '0x4ACD0 Group1.Count[I2]
            mSummonedMonsterGroupCount(1) = binReader.ReadUInt16()   '0x4ACD2 Group2.Count[I2]
            mSummonedMonsterGroupCount(2) = binReader.ReadUInt16()   '0x4ACD4 Group3.Count[I2]
            mSummonedMonsterGroupCode(0) = binReader.ReadUInt16()    '0x4ACD6 Group1.Code[I2]
            mSummonedMonsterGroupCode(1) = binReader.ReadUInt16()    '04AACD8 Group2.Code[I2]
            mSummonedMonsterGroupCode(2) = binReader.ReadUInt16()    '0x4ACDA Group3.Code[I2]
            mSummonedMonsterGroupName(0) = binReader.ReadString()    '0x4ACDC Group1.Name ("A DINK") String(15) 0x00
            binReader.BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(0).Length
            mSummonedMonsterGroupName(1) = binReader.ReadString()    '0x4ACEC Group2.Name ("ENTELECHY FUFF") String(15) 0x77
            binReader.BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(1).Length
            mSummonedMonsterGroupName(2) = binReader.ReadString()    '0x4ACFC Group3.Name ("VAMPIRE LORD") String(15) 0x70
            binReader.BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(2).Length

            binReader.BaseStream.Position += 4 + (16 * 15)          '0x4AD0C-FF 
        Catch ex As Exception When ex.message.toupper.Contains("OVERFLOW")
            Debug.WriteLine(String.Format("Read Failed @ 0x{0:X00000}{1}{2}", binReader.BaseStream.Position, vbCrLf, ex.ToString))
            Throw
        End Try
    End Sub
    Public Overrides Sub Save(binWriter As BinaryWriter)
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
End Class