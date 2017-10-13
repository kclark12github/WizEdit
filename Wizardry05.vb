'Wizardry05.cls
'   Main Class for Heart of the Maelstrom...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/19    Ken Clark       Migrated to VS2017;
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class Wizardry05
    Inherits WizEditBase
    Public Sub New(FileName As String, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(FileName, Caption, Icon, BoxArt, Parent)
        Read()
    End Sub
    Public Overrides ReadOnly Property CharacterDataOffset As Int32
        Get
            Return &H25C00
        End Get
    End Property
    Public Overrides ReadOnly Property CharactersMax As Short
        Get
            Return 20
        End Get
    End Property
    Public Overrides ReadOnly Property MasterItemList As ItemData()
        Get
            If mMasterItemList Is Nothing Then
                mMasterItemList = {
                    New ItemData("Misc: Broken Item", 1000),
                    New ItemData("Special: Orb of Earithin", 1001),
                    New ItemData("Special: Neutral Crystal", 1002),
                    New ItemData("Special: Crystal of Evil", 1003),
                    New ItemData("Special: Crystal of Good", 1004),
                    New ItemData("Special: Ship in a Bottle", 1005),
                    New ItemData("Weapon: Staff of Earth", 1006),
                    New ItemData("Magic: Amulet of Air", 1007),
                    New ItemData("Misc: Holy Water", 1008),
                    New ItemData("Magic: Rod of Fire", 1009),
                    New ItemData("Special: Gold Medallion", 1010),
                    New ItemData("Special: Orb of Mhuuzfes", 1011),
                    New ItemData("Weapon: Butterfly Knife", 1012),
                    New ItemData("Weapon: Short Sword", 1013),
                    New ItemData("Weapon: Broad Sword", 1014),
                    New ItemData("Weapon: Mace", 1015),
                    New ItemData("Weapon: Staff of Earth", 1016),
                    New ItemData("Weapon: Hand Axe", 1017),
                    New ItemData("Weapon: Battle Axe", 1018),
                    New ItemData("Weapon: Dagger", 1019),
                    New ItemData("Weapon: Flail", 1020),
                    New ItemData("Shield: Round Shield", 1021),
                    New ItemData("Shield: Heater Shield", 1022),
                    New ItemData("Armor: Mage's Robes", 1023),
                    New ItemData("Armor: Cuirass", 1024),
                    New ItemData("Armor: Hauberk", 1025),
                    New ItemData("Armor: Breast Plate", 1026),
                    New ItemData("Armor: Plate Armor", 1027),
                    New ItemData("Helm: Sallet", 1028),
                    New ItemData("Magic: Potion of Dios", 1029),
                    New ItemData("Magic: Latumofis Oil", 1030),
                    New ItemData("Weapon: Short Sword +1", 1031),
                    New ItemData("Weapon: Broad Sword +1", 1032),
                    New ItemData("Weapon: Mace +1", 1033),
                    New ItemData("Weapon: Battle Axe +1", 1034),
                    New ItemData("Weapon: Nunchuka", 1035),
                    New ItemData("Weapon: Dagger +1", 1036),
                    New ItemData("Magic: Scroll of Katino", 1037),
                    New ItemData("Armor: Cuirass +1", 1038),
                    New ItemData("Armor: Hauberk +1", 1039),
                    New ItemData("Armor: Breast Plate +1", 1040),
                    New ItemData("Armor: Plate Armor +1", 1041),
                    New ItemData("Shield: Heater +1", 1042),
                    New ItemData("Helm: Bascinet", 1043),
                    New ItemData("Gauntlets: Gloves of Iron", 1044),
                    New ItemData("Magic: Scroll of Badios", 1045),
                    New ItemData("Magic: Potion of Halito", 1046),
                    New ItemData("Weapon: Short Sword -1", 1047),
                    New ItemData("Weapon: Broadsword -1", 1048),
                    New ItemData("Weapon: Mace -1", 1049),
                    New ItemData("Weapon: Dagger -1", 1050),
                    New ItemData("Weapon: Battle Axe -1", 1051),
                    New ItemData("Weapon: Margaux's Flail", 1052),
                    New ItemData("Special: Bag of Gems", 1053),
                    New ItemData("Weapon: Wizard's Staff", 1054),
                    New ItemData("Weapon: Flametongue", 1055),
                    New ItemData("Shield: Round Shield -1", 1056),
                    New ItemData("Armor: Cuirass -1", 1057),
                    New ItemData("Armor: Hauberk -1", 1058),
                    New ItemData("Armor: Breast Plate -1", 1059),
                    New ItemData("Armor: Plate Armor -1", 1060),
                    New ItemData("Helm: Sallet -1", 1061),
                    New ItemData("Magic: Potion of Sopic", 1062),
                    New ItemData("Special: Gold Ring", 1063),
                    New ItemData("Special: Salamander Ring", 1064),
                    New ItemData("Special: Serpent's Tooth", 1065),
                    New ItemData("Weapon: Short Sword +2", 1066),
                    New ItemData("Weapon: Broad Sword +2", 1067),
                    New ItemData("Weapon: Battle Axe +2", 1068),
                    New ItemData("Weapon: Ivory Blade (G)", 1069),
                    New ItemData("Weapon: Ebony Blade (E)", 1070),
                    New ItemData("Weapon: Amber Blade (N)", 1071),
                    New ItemData("Weapon: Mace +2", 1072),
                    New ItemData("Gauntlets: Gloves of Mithril", 1073),
                    New ItemData("Magic: Amulet of Dailko", 1074),
                    New ItemData("Armor: Cuirass +2", 1075),
                    New ItemData("Shield: Heater +2", 1076),
                    New ItemData("Armor: Displacer Robes", 1077),
                    New ItemData("Armor: Hauberk +2", 1078),
                    New ItemData("Armor: Breast Plate +2", 1079),
                    New ItemData("Armor: Plate Armor +2", 1080),
                    New ItemData("Armor: Armet", 1081),
                    New ItemData("Armor: Wargan Robes", 1082),
                    New ItemData("Weapon: Giant's Club", 1083),
                    New ItemData("Weapon: Blade Cuisinart'", 1084),
                    New ItemData("Weapon: Shepherd Crook", 1085),
                    New ItemData("Weapon: Unholy Axe", 1086),
                    New ItemData("Weapon: Rod of Death", 1087),
                    New ItemData("Special: Gem of Exorcism", 1088),
                    New ItemData("Special: Bag of Emeralds", 1089),
                    New ItemData("Special: Bag of Garnets", 1090),
                    New ItemData("Special: Blue Pearl", 1091),
                    New ItemData("Special: Ruby Slippers", 1092),
                    New ItemData("Weapon: Necrology Rod", 1093),
                    New ItemData("Misc: Book of Life", 1094),
                    New ItemData("Misc: Book of Death", 1095),
                    New ItemData("Special: Dragon's Tooth", 1096),
                    New ItemData("Special: Trollkin Ring", 1097),
                    New ItemData("Special: Rabbit's Foot", 1098),
                    New ItemData("Special: Thief's Pick", 1099),
                    New ItemData("Misc: Book of Demons", 1100),
                    New ItemData("Weapon: Butterfly Knife", 1101),
                    New ItemData("Special: Gold Tiara", 1102),
                    New ItemData("Gauntlets: Mantis Gloves", 1103)
                }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataDirectory As String
        Get
            Return "UWAPath05"
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataFile As String
        Get
            Return "Wiz05DataFile"
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioDataOffset As String
        Get
            Return &H25800
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioName As String
        Get
            Return "HEART OF THE MAELSTROM"
        End Get
    End Property
End Class