'Character04.vb
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
        'Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", .BaseStream.Position))
        Try
            With binReader
                mName = .ReadString()                          '0x4AC00   Pascal Varying Length String Format...
                If mName <> "WERDNA" Then Throw New Exception(String.Format("Unexpected data found @ initial Save Game offset 0x{0:X00000} - Expected ""WERDNA"" but found ""{1}""", initialOffset, mName))
                .BaseStream.Position += NamePasswordLengthMax - mName.Length
                mPassword = .ReadString()                      '0x4AC10   Pascal Varying Length String Format...
                .BaseStream.Position += NamePasswordLengthMax - mPassword.Length

                mOut = .ReadUInt16()                            '0x4AC20    00 00 = No; 01 00 = Yes;    'TODO: Confirm
                mRace = .ReadUInt16()                           '0x4AC22    01 00 = Human
                mProfession = .ReadUInt16()                     '0x4AC24    01 00 = Mage
                mAgeInWeeks = .ReadUInt16()                     '0x4AC26    50 14 = 5200 Weeks (100 years) Alive...
                mStatus = .ReadUInt16()                         '0x4AC28    00 00 = OK
                mAlignment = .ReadUInt16()                      '0x4AC2A    03 00 = Evil

                Me.Statistics = .ReadUInt32()                   '0x4AC2C    94 52 94 52 = 20/20/20/20/20/20
                .BaseStream.Position += 4                       '0x4AC30

                For i As Short = 0 To 2                         '0x4AC34
                    mGoldPacked(i) = .ReadUInt16()
                Next i
                mGold = mBase.I6toD(mGoldPacked)

                mItemCount = .ReadUInt16()                      '0x4AC3A
                For i As Short = 0 To ItemListMax - 1           '0x4AC3C    List of Items (stowing not an option in Wiz04...)
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

                For i As Short = 0 To 2                         '0x4AC7C (Keystrokes)
                    mExperiencePacked(i) = .ReadUInt16()
                Next i
                mExperience = mBase.I6toD(mExperiencePacked)

                mLVL.Read(binReader)                            '0x4AC82
                mHP.Read(binReader)                             '0x4AC86

                For i As Short = 0 To 7                         '0x4AC8A    Need to mask as bits...
                    mSpellBooks(i) = .ReadByte()
                Next i

                For i As Short = 0 To SpellLevelMax - 1         '0x4AC92-9F
                    mMageSpellPoints(i) = .ReadUInt16()
                Next i
                For i As Short = 0 To SpellLevelMax - 1         '0x4ACA0-AD
                    mPriestSpellPoints(i) = .ReadUInt16()
                Next i
                '0x4AC
                .BaseStream.Position += 2                       '0x4ACAE
                mArmorClass = .ReadUInt16()                     '0x4ACB0
                .BaseStream.Position += 24                      '0x4ACB2-C9

                mLocation = .ReadUInt16()                       '0x4ACCA
                mDown = .ReadUInt16()                           '0x4ACCC    Screen says -2 but data says 0E 00
                mHonors = .ReadUInt16()                         '0x4ACCE    Need more testing, but 1 = ">"

                mSummonedMonsterGroupCount(0) = .ReadUInt16()   '0x4ACD0 Group1.Count[I2]
                mSummonedMonsterGroupCount(1) = .ReadUInt16()   '0x4ACD2 Group2.Count[I2]
                mSummonedMonsterGroupCount(2) = .ReadUInt16()   '0x4ACD4 Group3.Count[I2]
                mSummonedMonsterGroupCode(0) = .ReadUInt16()    '0x4ACD6 Group1.Code[I2] ("A DINK") 0x00
                mSummonedMonsterGroupCode(1) = .ReadUInt16()    '04AACD8 Group2.Code[I2] ("ENTELECHY FUFF") 0x77
                mSummonedMonsterGroupCode(2) = .ReadUInt16()    '0x4ACDA Group3.Code[I2] ("VAMPIRE LORD") 0x70
                mSummonedMonsterGroupName(0) = .ReadString()    '0x4ACDC Group1.Name ("A DINK") String(15) 0x00
                .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(0).Length
                mSummonedMonsterGroupName(1) = .ReadString()    '0x4ACEC Group2.Name ("ENTELECHY FUFF") String(15) 0x77
                .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(1).Length
                mSummonedMonsterGroupName(2) = .ReadString()    '0x4ACFC Group3.Name ("VAMPIRE LORD") String(15) 0x70
                .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(2).Length

                .BaseStream.Position += 4 + (16 * 15)           '0x4AD0C-FF 
            End With
        Catch ex As Exception When ex.Message.ToUpper.Contains("OVERFLOW")
            Debug.WriteLine(String.Format("Read Failed @ 0x{0:X00000}{1}{2}", binReader.BaseStream.Position, vbCrLf, ex.ToString))
            Throw
        End Try
    End Sub
    Public Overrides Sub Save(binWriter As BinaryWriter)
        With binWriter
            .Write(mName)                                  '0x1D800   Pascal Varying Length String Format...
            .BaseStream.Position += NamePasswordLengthMax - mName.Length
            .Write(mPassword)                              '0x1D810   Pascal Varying Length String Format...
            .BaseStream.Position += NamePasswordLengthMax - mPassword.Length

            .Write(mOut)                                   '0x1D820    00 00 = No; 01 00 = Yes;
            .Write(mRace)                                  '0x1D822    01 00 = Human
            .Write(mProfession)                            '0x1D824    06 00 = Lord
            .Write(mAgeInWeeks)                            '0x1D826    0C 03 = Weeks Alive...
            .Write(mStatus)                                '0x1D828    00 00 = OK
            .Write(mAlignment)                             '0x1D82A    02 00 = Neutral

            .Write(Me.Statistics)                          '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
            .BaseStream.Position += 4                      '0x1D830

            For i As Short = 0 To 2                        '0x1D834
                .Write(mGoldPacked(i))
            Next i

            .Write(mItemCount)                             '0x1D83A
            For i As Short = 0 To ItemListMax - 1          '0x1D83C    List of Items (stowing not an option in Wiz01...)
                mItemList(i).Save(binWriter)
            Next i

            For i As Short = 0 To 2                        '0x1D87C
                .Write(mExperiencePacked(i))
            Next i

            mLVL.Save(binWriter)                           '0x1D882
            mHP.Save(binWriter)                            '0x1D886

            For i As Short = 0 To 7                        '0x1D88A    Need to mask as bits...
                .Write(mSpellBooks(i))
            Next i

            For i As Short = 0 To SpellLevelMax - 1        '0x1D892
                .Write(mMageSpellPoints(i))
            Next i

            For i As Short = 0 To SpellLevelMax - 1        '0x1D8A0
                .Write(mPriestSpellPoints(i))
            Next i
            .BaseStream.Position += 2                      '0x1D8AE
            .Write(mArmorClass)                            '0x1D8B0
            .BaseStream.Position += 24                     '0x1D8B2
            .Write(mLocation)                              '0x1D8CA    Some sort of packed variable...
            .Write(mDown)                                  '0x1D8CC    Seems to be a simple 2-byte Int16...
            .Write(mHonors)                                '0x1D8CE    Need more testing, but 1 = ">"

            .Write(mSummonedMonsterGroupCount(0))           '0x4ACD0 Group1.Count[I2]
            .Write(mSummonedMonsterGroupCount(1))           '0x4ACD2 Group2.Count[I2]
            .Write(mSummonedMonsterGroupCount(2))           '0x4ACD4 Group3.Count[I2]
            .Write(mSummonedMonsterGroupCode(0))            '0x4ACD6 Group1.Code[I2] ("A DINK") 0x00
            .Write(mSummonedMonsterGroupCode(1))            '04AACD8 Group2.Code[I2] ("ENTELECHY FUFF") 0x77
            .Write(mSummonedMonsterGroupCode(2))            '0x4ACDA Group3.Code[I2] ("VAMPIRE LORD") 0x70
            .Write(mSummonedMonsterGroupName(0))            '0x4ACDC Group1.Name ("A DINK") String(15) 0x00
            .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(0).Length
            .Write(mSummonedMonsterGroupName(1))            '0x4ACEC Group2.Name ("ENTELECHY FUFF") String(15) 0x77
            .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(1).Length
            .Write(mSummonedMonsterGroupName(2))            '0x4ACFC Group3.Name ("VAMPIRE LORD") String(15) 0x70
            .BaseStream.Position += NamePasswordLengthMax - mSummonedMonsterGroupName(2).Length

            .BaseStream.Position += 4 + (16 * 15)           '0x4AD0C-FF 
        End With
    End Sub
End Class