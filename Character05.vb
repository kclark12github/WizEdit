'Character05.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class Character05
    Inherits CharacterBase
    Protected mMarks As UInt16
    Protected mRIP As UInt16
    Protected mSwim As UInt16
    Public Sub New(ByVal Base As Wizardry05)
        MyBase.New(Base)
    End Sub
    Public Property Marks As UInt16
        Get
            Return mMarks
        End Get
        Set(value As UInt16)
            mMarks = value
        End Set
    End Property
    Public Property RIP As UInt16
        Get
            Return mRIP
        End Get
        Set(value As UInt16)
            mRIP = value
        End Set
    End Property
    Public Property Swim As UInt16
        Get
            Return mSwim
        End Get
        Set(value As UInt16)
            mSwim = value
        End Set
    End Property
    Public Overrides Sub Read(binReader As BinaryReader)
        Dim initialOffset As Long = binReader.BaseStream.Position
        'Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", .BaseStream.Position))
        Try 'Character is 246 bytes
            With binReader                                      'Offset/Sample
                'Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", .BaseStream.Position))
                mName = .ReadString()                           '000/0x4C00   Pascal Varying Length String Format...
                .BaseStream.Position += NamePasswordLengthMax - mName.Length
                mPassword = .ReadString()                       '016/0x4C10   Pascal Varying Length String Format...
                .BaseStream.Position += NamePasswordLengthMax - mPassword.Length

                mOut = .ReadUInt16()                            '032/0x4C20    00 00 = No; 01 00 = Yes;    'TODO: Confirm
                mRace = .ReadUInt16()                           '034/0x4C22    01 00 = Human
                mProfession = .ReadUInt16()                     '036/0x4C24    01 00 = Mage
                mAgeInWeeks = .ReadUInt16()                     '038/0x4C26    50 14 = 5200 Weeks (100 years) Alive...
                mStatus = .ReadUInt16()                         '040/0x4C28    00 00 = OK
                mAlignment = .ReadUInt16()                      '042/0x4C2A    03 00 = Evil

                Me.Statistics = .ReadUInt32()                   '044/0x4C2C    94 52 94 52 = 20/20/20/20/20/20
                .BaseStream.Position += 4                       '048/0x4C30

                For i As Short = 0 To 2                         '052/0x4C34
                    mGoldPacked(i) = .ReadUInt16()
                Next i
                mGold = mBase.I6toD(mGoldPacked)

                mItemCount = .ReadUInt16()                      '058/0x4C3A
                For i As Short = 0 To ItemListMax - 1           '060/0x4C3C    List of Items (stowing not an option in Wiz05...)
                    CType(mItemList(i), Item05).Read(binReader)
                Next i

                For i As Short = 0 To 2                         '092/0x4C5C
                    mExperiencePacked(i) = .ReadUInt16()
                Next i
                mExperience = mBase.I6toD(mExperiencePacked)

                mLVL.Read(binReader)                            '098/0x4C62
                mHP.Read(binReader)                             '102/0x4C66

                For i As Short = 0 To 7                         '106/0x4C6A    Need to mask as bits...
                    mSpellBooks(i) = .ReadByte()
                Next i
                For i As Short = 0 To SpellLevelMax - 1         '114/0x4C72
                    mMageSpellPoints(i) = .ReadByte()
                Next i
                .BaseStream.Position += 1 '(alignment)
                For i As Short = 0 To SpellLevelMax - 1         '122/0x4C7A
                    mPriestSpellPoints(i) = .ReadByte()
                Next i
                .BaseStream.Position += 1 '(alignment)
                .BaseStream.Position += 2                       '130/0x4C82 - Unknown (but something)
                mArmorClass = .ReadUInt16()                     '132/0x4C84
                .BaseStream.Position += 28                      '134/0x4C86-0x4CA1 - Unknown (but lots of somethings)
                mMarks = .ReadUInt16()                          '162/0x4CA2
                .BaseStream.Position += 8                       '164/0x4C86-0x4CA1 - Unknown (zeros)
                mHonors = .ReadUInt16()                         '172/0x4CAC    
                .BaseStream.Position += 72                      '174/0x4CAE
                '246

                'mLocation = .ReadUInt16()                       '0x4CCA
                'mDown = .ReadUInt16()                           '0x4CCC    Screen says -2 but data says 0E 00
            End With
        Catch ex As Exception When ex.Message.ToUpper.Contains("OVERFLOW")
            Debug.WriteLine(String.Format("Read Failed @ 0x{0:X00000}{1}{2}", binReader.BaseStream.Position, vbCrLf, ex.ToString))
            Throw
        End Try
    End Sub
    Public Overrides Sub Save(binWriter As BinaryWriter)
        'With binWriter
        '    .Write(mName)                                  '0x1D800   Pascal Varying Length String Format...
        '    .BaseStream.Position += NamePasswordLengthMax - mName.Length
        '    .Write(mPassword)                              '0x1D810   Pascal Varying Length String Format...
        '    .BaseStream.Position += NamePasswordLengthMax - mPassword.Length

        '    .Write(mOut)                                   '0x1D820    00 00 = No; 01 00 = Yes;
        '    .Write(mRace)                                  '0x1D822    01 00 = Human
        '    .Write(mProfession)                            '0x1D824    06 00 = Lord
        '    .Write(mAgeInWeeks)                            '0x1D826    0C 03 = Weeks Alive...
        '    .Write(mStatus)                                '0x1D828    00 00 = OK
        '    .Write(mAlignment)                             '0x1D82A    02 00 = Neutral

        '    .Write(Me.Statistics)                          '0x1D82C    94 52 94 52 = 20/20/20/20/20/20
        '    .BaseStream.Position += 4                      '0x1D830

        '    For i As Short = 0 To 2                        '0x1D834
        '        .Write(mGoldPacked(i))
        '    Next i

        '    .Write(mItemCount)                             '0x1D83A
        '    For i As Short = 0 To ItemListMax - 1          '0x1D83C    List of Items (stowing not an option in Wiz01...)
        '        CType(mItemList(i), Item05).Save(binWriter)
        '    Next i

        '    For i As Short = 0 To 2                        '0x1D87C
        '        .Write(mExperiencePacked(i))
        '    Next i

        '    mLVL.Save(binWriter)                           '0x1D882
        '    mHP.Save(binWriter)                            '0x1D886

        '    For i As Short = 0 To 7                        '0x1D88A    Need to mask as bits...
        '        .Write(mSpellBooks(i))
        '    Next i

        '    For i As Short = 0 To SpellLevelMax - 1        '0x1D892
        '        .Write(mMageSpellPoints(i))
        '    Next i

        '    For i As Short = 0 To SpellLevelMax - 1        '0x1D8A0
        '        .Write(mPriestSpellPoints(i))
        '    Next i
        '    .BaseStream.Position += 2                      '0x1D8AE
        '    .Write(mArmorClass)                            '0x1D8B0
        '    .BaseStream.Position += 24                     '0x1D8B2
        '    .Write(mLocation)                              '0x1D8CA    Some sort of packed variable...
        '    .Write(mDown)                                  '0x1D8CC    Seems to be a simple 2-byte Int16...
        '    .Write(mHonors)                                '0x1D8CE    Need more testing, but 1 = ">"


        '    .BaseStream.Position += 4 + (16 * 15)           '0x4AD0C-FF 
        'End With
    End Sub
End Class