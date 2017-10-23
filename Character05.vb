'Character05.vb
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
    Protected myMageSpellPoints(SpellLevelMax - 1) As Byte
    Protected myPriestSpellPoints(SpellLevelMax - 1) As Byte
    Public Sub New(ByVal Base As Wizardry05)
        MyBase.New(Base)
        ReDim mItemList(ItemListMax - 1)
        For iItem As Short = 0 To ItemListMax - 1
            mItemList(iItem) = New Item05(Base)
        Next iItem
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
    Public Overloads Property MageSpellPoints(ByVal Level As Short) As Byte
        Get
            Return myMageSpellPoints(Level)
        End Get
        Set(value As Byte)
            mMageSpellPoints(Level) = value
        End Set
    End Property
    Public Overloads Property PriestSpellPoints(ByVal Level As Short) As Byte
        Get
            Return mPriestSpellPoints(Level)
        End Get
        Set(value As Byte)
            mPriestSpellPoints(Level) = value
        End Set
    End Property
    Public Overrides Sub Read(binReader As BinaryReader)
        Dim initialOffset As Long = binReader.BaseStream.Position
        Try 'Character is 246 bytes
            With binReader                                      'Offset/Sample
                mName = .ReadString()                           '000/0x4C00   Pascal Varying Length String Format...
                .BaseStream.Position += NamePasswordLengthMax - mName.Length
                mPassword = .ReadString()                       '016/0x4C10   Pascal Varying Length String Format...
                .BaseStream.Position += NamePasswordLengthMax - mPassword.Length
                '???mOut = .ReadUInt16()                        '    00 00 = No; 01 00 = Yes;    'TODO: Confirm
                mRace = .ReadUInt16()                           '032/0x4C20     01 00 = Human
                mProfession = .ReadUInt16()                     '034/0x4C22     01 00 = Mage
                mAlignment = .ReadUInt16()                      '036/0x4C24     03 00 = Evil
                .BaseStream.Position += 2                       '038/0x4C26
                mStatus = .ReadUInt16()                         '040/0x4C28     00 00 = OK
                mAgeInWeeks = .ReadUInt16()                     '042/0x4C2A     50 14 = 5200 Weeks (100 years) Alive...
                Me.Statistics = .ReadUInt32()                   '044/0x4C2C     94 52 94 52 = 20/20/20/20/20/20
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
                    myMageSpellPoints(i) = .ReadByte()
                Next i
                .BaseStream.Position += 1 '(alignment)
                For i As Short = 0 To SpellLevelMax - 1         '122/0x4C7A
                    myPriestSpellPoints(i) = .ReadByte()
                Next i
                .BaseStream.Position += 1 '(alignment)
                .BaseStream.Position += 2                       '130/0x4C82 - Unknown (but something)
                mArmorClass = .ReadUInt16()                     '132/0x4C84
                .BaseStream.Position += 28                      '134/0x4C86 - Unknown (but lots of somethings)
                mMarks = .ReadUInt16()                          '162/0x4CA2
                .BaseStream.Position += 4                       '164/0x4C86 - Unknown (zeros)
                mRIP = .ReadUInt16()                            '168/0x4CA8
                .BaseStream.Position += 2                       '170/0x4CAA - Unknown (zeros)
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
        Dim initialOffset As Long = binWriter.BaseStream.Position
        Dim Offset As Long = 0
        Debug.WriteLine(String.Format("Character Data @ 0x{0:X00000}", binWriter.BaseStream.Position))
        With binWriter                                      'Offset/Sample
            .Write(mName)                                   '000/0x4C00   Pascal Varying Length String Format...
            .BaseStream.Position += NamePasswordLengthMax - mName.Length
            Offset += 16 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mPassword)                               '016/0x4C10   Pascal Varying Length String Format...
            .BaseStream.Position += NamePasswordLengthMax - mPassword.Length
            Offset += 16 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            '???mOut = .ReadUInt16()                        '    00 00 = No; 01 00 = Yes;    'TODO: Confirm
            .Write(mRace)                                   '032/0x4C20     01 00 = Human
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mProfession)                             '034/0x4C22     01 00 = Mage
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mAlignment)                              '036/0x4C24     03 00 = Evil
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 2                       '038/0x4C26
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mStatus)                                 '040/0x4C28     00 00 = OK
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mAgeInWeeks)                             '042/0x4C2A     50 14 = 5200 Weeks (100 years) Alive...
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(Me.Statistics)                           '044/0x4C2C     94 52 94 52 = 20/20/20/20/20/20
            Offset += 4 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 4                       '048/0x4C30
            Offset += 4 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To 2                         '052/0x4C34
                .Write(mGoldPacked(i))
            Next i
            Offset += 6 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mItemCount)                              '058/0x4C3A
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To ItemListMax - 1           '060/0x4C3C    List of Items (stowing not an option in Wiz05...)
                CType(mItemList(i), Item05).Save(binWriter)
            Next i
            Offset += (4 * 8) : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To 2                         '092/0x4C5C
                .Write(mExperiencePacked(i))
            Next i
            Offset += 6 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            mLVL.Save(binWriter)                            '098/0x4C62
            Offset += 4 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            mHP.Save(binWriter)                             '102/0x4C66
            Offset += 4 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To 7                         '106/0x4C6A    Need to mask as bits...
                .Write(mSpellBooks(i))
            Next i
            Offset += 8 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To SpellLevelMax - 1         '114/0x4C72
                .Write(myMageSpellPoints(i))
            Next i
            .BaseStream.Position += 1 '(alignment)
            Offset += 8 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            For i As Short = 0 To SpellLevelMax - 1         '122/0x4C7A
                .Write(myPriestSpellPoints(i))
            Next i
            .BaseStream.Position += 1 '(alignment)
            Offset += 8 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 2                       '130/0x4C82 - Unknown (but something)
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mArmorClass)                             '132/0x4C84
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 28                      '134/0x4C86 - Unknown (but lots of somethings)
            Offset += 28 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mMarks)                                  '162/0x4CA2
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 4                       '164/0x4C86 - Unknown (zeros)
            Offset += 4 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mRIP)                                    '168/0x4CA8
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 2                       '170/0x4CAA - Unknown (zeros)
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .Write(mHonors)                                 '172/0x4CAC    
            Offset += 2 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
            .BaseStream.Position += 72                      '174/0x4CAE
            Offset += 72 : Debug.Assert(.BaseStream.Position = initialOffset + Offset)
        End With
        '1) 0x4C00
        '2) 0x4CF6
        '3) 0x4DEC
        '4) 0x4EE2
    End Sub
End Class