'ItemBase.cls
'   Item Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class ItemBase
    Public Sub New(ByVal Base As WizEditBase)
        mBase = Base
        mEquipped = 0
        mCursed = 0
        mIdentified = 0
        mItemCode = 0
    End Sub
    Private mBase As WizEditBase
    Private mEquipped As UInt16
    Private mCursed As UInt16
    Private mIdentified As UInt16
    Private mItemCode As UInt16
    Public Property ItemCode As UInt16
        Get
            Return mItemCode
        End Get
        Set(value As UInt16)
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
        With binReader
            mEquipped = .ReadUInt16()
            mCursed = .ReadUInt16()
            mIdentified = .ReadUInt16()
            mItemCode = .ReadUInt16()
        End With
    End Sub
    Public Sub Save(binWriter As BinaryWriter)
        With binWriter
            .Write(mEquipped)
            .Write(mCursed)
            .Write(mIdentified)
            .Write(mItemCode)
        End With
    End Sub
    Public Overrides Function ToString() As String
        '    Item = vbTab & ItemList(x.ItemCode) & "; Code: " & x.ItemCode & "; Equipped: "
        '    If x.Identified Then Item &= "; Identified"
        '    If x.Equipped Then Item &= "; **EQUIPPED**"
        '    If x.Cursed Then strItem &= "; --CURSED--"
        ToString = String.Format("{0}{1}{2}", vbTab, IIf(mCursed, "-", IIf(mEquipped, "*", " ")), mBase.MasterItemList(mItemCode))
    End Function
End Class