'WizEditBase.cls
'   Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/11/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class WizEditBase
    Public Sub New()
        mPath = Nothing
        mFileInfo = Nothing
        mBoxArt = Nothing
        mCaption = Nothing
        mIcon = Nothing
        mParent = Nothing
    End Sub
    Public Sub New(Path As String, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        mPath = Path
        If Path IsNot Nothing AndAlso Path <> "" Then mFileInfo = New FileInfo(Path)
        mBoxArt = BoxArt
        mCaption = Caption
        mIcon = Icon
        mParent = Parent
    End Sub
#Region "Properties"
#Region "Declarations"
    Private mBoxArt As Image
    Private mCaption As String
    Private mForm As Form
    Private mIcon As Icon
    Private mParent As Form

    Private mFileInfo As FileInfo = Nothing
    Private mPath As String = vbNullString
    Private mRegKey As String = "Software\KClark Software"
#End Region
    Public ReadOnly Property DirectoryName As String
        Get
            If mFileInfo IsNot Nothing Then Return mFileInfo.DirectoryName
            Return vbNullString
        End Get
    End Property
    Public ReadOnly Property FileName As String
        Get
            If mFileInfo IsNot Nothing Then Return mFileInfo.Name
            Return vbNullString
        End Get
    End Property
    Public ReadOnly Property Path As String
        Get
            Return mPath
        End Get
    End Property
#End Region
#Region "Methods"
#Region "Conversion"
    Public Sub DtoI6(ByVal x As Double, ByRef Data() As UInt16)
        Dim r1 As Double
        Dim r2 As Double
        Dim r3 As Double

        r3 = x \ 100000000.0#
        Data(2) = CInt(r3)
        x = x - (r3 * 100000000.0#)
        r2 = x \ 10000.0#
        Data(1) = r2
        x = x - (r2 * 10000.0#)
        r1 = x
        Data(0) = r1
    End Sub
    Public Function I6toD(ByVal Data() As UInt16) As Double
        Dim r1 As Double
        Dim r2 As Double
        Dim r3 As Double

        r1 = Data(0)
        r2 = Data(1) * 10000.0#
        r3 = Data(2) * 100000000.0#

        I6toD = r1 + r2 + r3
    End Function
#End Region
#Region "Registry"
    Public Function GetRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal [Default] As Object) As Object
        Dim Reg As RegistryKey = Nothing
        GetRegSetting = Nothing
        Try
            Reg = Registry.CurrentUser.OpenSubKey(String.Format("{0}\{1}\{2}", mRegKey, Application.ProductName, KeyName)) : If IsNothing(Reg) Then Exit Try
            GetRegSetting = Reg.GetValue(ValueName, [Default])
        Catch ex As System.Exception
        Finally : If Not IsNothing(Reg) Then Reg.Close()
        End Try
    End Function
    Public Sub SaveRegSetting(ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Object)
        Dim Reg As RegistryKey = Nothing
        Dim CurrentValue As Object = Nothing
        Try
            If KeyName Is Nothing Then Throw New ArgumentException("KeyName must be provided.")
            If ValueName Is Nothing Then Throw New ArgumentException("ValueName must be provided.")
            If Value Is Nothing Then Throw New ArgumentException("Value must be provided.")

            KeyName = String.Format("{0}\{1}\{2}", mRegKey, Application.ProductName, KeyName)
            Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            If Reg Is Nothing Then
                'Iterate through the KeyName making sure each sub-key exists (create as necessary)...
                Dim SubKeys() As String = KeyName.Split("\")
                Dim Key As String = SubKeys(0)
                For i As Short = 1 To SubKeys.Length - 1
                    Dim SubKey As String = String.Format("{0}\{1}", Key, SubKeys(i))
                    Reg = Registry.CurrentUser.OpenSubKey(SubKey)
                    If Reg Is Nothing Then
                        Reg = Registry.CurrentUser.OpenSubKey(Key, True)
                        Reg.CreateSubKey(SubKeys(i))
                    End If
                    Reg.Close() : Reg = Nothing
                    Key = SubKey
                Next i
                Reg = Registry.CurrentUser.OpenSubKey(KeyName, True)
            End If
            CurrentValue = Reg.GetValue(ValueName)
            If CurrentValue Is Nothing OrElse CurrentValue.ToString <> Value.ToString Then Reg.SetValue(ValueName, Value)
        Catch ex As System.Exception
        Finally : If Reg IsNot Nothing Then Reg.Close()
        End Try
    End Sub
#End Region
#Region "Utility"
    Public Function IsPrintable(ByVal xByte As Byte) As Boolean
        If xByte < 32 Then Return False
        Select Case xByte
            Case 127, 129, 141, 143, 144, 157 : Return False
            Case Else : Return True
        End Select
    End Function
    Public Function UpCase(uKey As Integer) As Integer
        If uKey > 96 And uKey < 123 Then
            UpCase = uKey - 32
        Else
            UpCase = uKey
        End If
    End Function
    Public Function ValidateByte(ByVal ctl As Control) As Boolean
        ValidateByte = False
        Try
            With ctl
                If .Text = "" Then .Text = "0"
                Dim iLimit As Byte = 99
                If Val(.Text) < 0 Or Val(.Text) > iLimit Then Beep() : .Text = "" : Exit Try
                '.Text = Format(.Text, "00")
            End With
            ValidateByte = True
        Finally
        End Try
    End Function
    Public Function ValidateI2(ByVal ctl As Control) As Boolean
        ValidateI2 = False
        Try
            With ctl
                If .Text = vbNullString Then .Text = "0"
                Dim iLimit As UInt16 = (2 ^ 16) - 1
                If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then Beep() : .Text = "" : Exit Try
                .Text = Format(.Text, "#,##0")
            End With
            ValidateI2 = True
        Finally
        End Try
    End Function
    Public Function ValidateI4(ByVal ctl As Control) As Boolean
        ValidateI4 = False
        Try
            With ctl
                If .Text = vbNullString Then .Text = "0"
                Dim iLimit As UInt32 = (2 ^ 31) - 1
                If Val(.Text) < 0 Or CLng(Val(.Text)) > iLimit Then Beep() : .Text = "" : Exit Try
                .Text = Format(.Text, "#,##0")
            End With
            ValidateI4 = True
        Finally
        End Try
    End Function
#End Region
    Public Sub Show()
        mForm = New frmWizardry15Base(Me, mCaption, mIcon, mBoxArt)
        mForm.ShowDialog(mParent)
    End Sub
#End Region
End Class
