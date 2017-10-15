'MonsterGroupData.cls
'   Item Metadata Class...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/16/19    Ken Clark       Created;
'=================================================================================================================================
Public Class MonsterGroupData
    Public Sub New(ByVal Text As String, ByVal Code As UInt16, ByVal Level As UInt16, Optional ByVal Notes As String = "")
        mCode = Code
        mLevel = Level
        mNotes = Notes
        mText = Text
    End Sub
#Region "Properties"
    Private mCode As UInt16
    Private mLevel As UInt16
    Private mNotes As String
    Private mText As String
    Public Property Code() As UShort
        Get
            Return mCode
        End Get
        Set(value As UShort)
            mCode = value
        End Set
    End Property
    Public Property Level() As UShort
        Get
            Return mLevel
        End Get
        Set(value As UShort)
            mLevel = value
        End Set
    End Property
    Public Property Notes() As String
        Get
            Return mNotes
        End Get
        Set(value As String)
            mNotes = value
        End Set
    End Property
    Public Property Text() As String
        Get
            Return mText
        End Get
        Set(value As String)
            mText = value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Overrides Function ToString() As String
        Return mText
    End Function
#End Region
End Class