'SpellBase.vb
'   Spell Base Class for WizEdit...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/14/17    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On
Public Class SpellBase
    Public Sub New(ByVal Name As String, ByVal Translation As String, ByVal Type As enumSpellType, ByVal Affects As enumSpellAffects, ByVal Category As enumSpellCategory, ByVal Level As Short)
        mName = Name
        mTranslation = Translation
        mType = Type
        mAffects = Affects
        mCategory = Category
        mLevel = Level
    End Sub
#Region "Properties"
#Region "Enumerations"
    Public Enum enumSpellCategory As Short
        Mage
        Priest
    End Enum
    Public Enum enumSpellType As Short
        AnyTime
        Camp
        Combat
        Looting
    End Enum
    Public Enum enumSpellAffects As Short
        Caster
        OnePerson
        EntireParty
        OneMonster
        OneGroup
        AllMonsters
        Variable
    End Enum
#End Region
#Region "Declarations"
    Private mName As String
    Private mTranslation As String
    Private mType As enumSpellType
    Private mAffects As enumSpellAffects
    Private mCategory As enumSpellCategory
    Private mLevel As Short
#End Region
    Public ReadOnly Property Name As String
        Get
            Return mName
        End Get
    End Property
    Public ReadOnly Property Translation As String
        Get
            Return mTranslation
        End Get
    End Property
    Public ReadOnly Property Type As String
        Get
            Return String.Format("{0}", mType)
        End Get
    End Property
    Public ReadOnly Property Affects As String
        Get
            Select Case mAffects
                Case enumSpellAffects.AllMonsters : Return "All Monsters"
                Case enumSpellAffects.Caster : Return "Caster"
                Case enumSpellAffects.EntireParty : Return "Entire Party"
                Case enumSpellAffects.OnePerson : Return "1 Person"
                Case enumSpellAffects.OneMonster : Return "1 Monster"
                Case enumSpellAffects.OneGroup : Return "1 Group"
                Case enumSpellAffects.Variable : Return "Variable"
            End Select
            Return "Unknown"
        End Get
    End Property
    Public ReadOnly Property Category As enumSpellCategory
        Get
            Return mCategory
        End Get
    End Property
    Public ReadOnly Property Level As Short
        Get
            Return mLevel
        End Get
    End Property
#End Region
#Region "Methods"
    Public Overrides Function ToString() As String
        Return String.Format("{1}{0}Level {2}{0}{3} (""{4}""){0}Type: {5}{0}Affects: {6}", New Object() {vbTab, Me.Category, Me.Level, Me.Name, Me.Translation, Me.Type, Me.Affects})
    End Function
#End Region
End Class