'Item05.vb
'   Item Class for WizEdit/Wizardry05...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/21/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class Item05
    Inherits ItemBase
    Private mItemStatus As UInt16
    Public Sub New(ByVal Base As WizEditBase)
        MyBase.New(Base)
        mItemStatus = 0
    End Sub
    Public Overrides Sub Read(binReader As BinaryReader)
        '	0x4C3C ItemCode1[I2];
        '	0x4C3C ItemStatus1[I2];  2^0 = Equipped; 2^1 = Identified;
        '       0x00=0000 0000 = Unidentified
        '       0x01=0000 0001 = Equipped & Unidentified (Weapon)
        '       0x02=0000 0010 = Identified
        '       0x03=0000 0011 = Equipped & Identified (Weapon)
        '       0x04=0000 0100 = Unidentified
        '       0x05=0000 0101 = Equipped & Unidentified
        '       0x06=0000 0110 = Identified
        '       0x07=0000 0111 = Equipped & Identified (Weapon)
        '       0x08=0000 1000
        '       0x09=0000 1001
        '       0x0A=0000 1010
        '       0x0B=0000 1011 = Equipped & Identified (Armor)
        '	0x4C3E Cursed1[I2];
        '	0x4C40 Identified1[I2];
        With binReader
            mItemCode = .ReadUInt16()
            mItemStatus = .ReadUInt16
        End With
        'TODO: Unpack ItemStatus into Equipped/Identified/Cursed properties...
    End Sub
    Public Overrides Sub Save(binWriter As BinaryWriter)
        'TODO: Repack Equipped/Identified/Cursed properties back into ItemStatus...
        With binWriter
            .Write(mItemCode)
            .Write(mItemStatus)
        End With
    End Sub
End Class