'ItemData.cls
'   Item Metadata Class...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/15/19    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class ItemData
    Public Sub New(ByVal Text As String, ByVal ItemCode As Short, Optional ByVal Category As String = "", Optional ByVal Effect As String = "", Optional ByVal Value As Integer = -1, Optional ByVal UserClass As String = "", Optional ByVal Notes As String = "")
        mCategory = Category
        mEffect = Effect
        mItemCode = ItemCode
        mNotes = Notes
        mUserClass = UserClass
        mText = Text
        mValue = Value
        If mCategory = "" AndAlso mText.Contains(":") Then mCategory = mText.Substring(0, mText.IndexOf(":"))
    End Sub
#Region "Properties"
    Private mCategory As String
    Private mEffect As String
    Private mNotes As String
    Private mText As String
    Private mItemCode As Short
    Private mUserClass As String
    Private mValue As Integer
    Public Property Category() As String
        Get
            Return mCategory
        End Get
        Set(value As String)
            mCategory = value
        End Set
    End Property
    Public Property Effect() As String
        Get
            Return mEffect
        End Get
        Set(value As String)
            mEffect = value
        End Set
    End Property
    Public Property ItemCode() As Short
        Get
            Return mItemCode
        End Get
        Set(value As Short)
            mItemCode = value
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
    Public Property UserClass() As String
        Get
            Return mUserClass
        End Get
        Set(value As String)
            mUserClass = value
        End Set
    End Property
    Public ReadOnly Property Tag As String
        Get
            Tag = String.Format("{0}; ", mText)
            If mCategory <> "" Then Tag &= String.Format("{0}; ", mCategory)
            If mUserClass <> "" Then Tag &= String.Format("User Class(es): {0}; ", mUserClass)
            If mValue <> 0 Then Tag &= String.Format("Value: {0}; ", mValue)
            If mEffect <> "" Then Tag &= String.Format("Effect: {0}; ", mEffect)
            If mNotes <> "" Then Tag &= String.Format("{0}", mNotes)
        End Get
    End Property
    Public Property Value() As Integer
        Get
            Return mValue
        End Get
        Set(value As Integer)
            mValue = value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Overrides Function ToString() As String
        Return mText
    End Function
#End Region
End Class