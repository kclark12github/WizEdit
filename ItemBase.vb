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
        ToString = String.Format("{0}{1}{2}", vbTab, IIf(mCursed, "-", IIf(mEquipped, "*", " ")), mBase.MasterItemList(mItemCode))
    End Function
End Class