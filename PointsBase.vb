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
Public Class PointsBase
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