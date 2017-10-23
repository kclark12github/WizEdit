'Wizardry03.vb
'   Main Class for The Legacy of Llylgamyn...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/19    Ken Clark       Migrated to VS2017;
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class Wizardry03
    Inherits WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(Caption, Icon, BoxArt, Parent)
    End Sub
    Public Overrides ReadOnly Property CharacterDataOffset As Int32
        Get
            Return &H1DE00
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
                New ItemData("Broken Item", 1000, "Misc", "", 0, "*", "None (The item you get when you use a scroll or invoke a special power of an object and the object breaks)"),
                New ItemData("Orb of Earithin", 1001, "Special", "", 0, "*", "Wins the Game"),
                New ItemData("Neutral Crystal", 1002, "Special", "", 0, "*", "Used to obtain Orb"),
                New ItemData("Crystal of Evil", 1003, "Special", "", 0, "*", "Invoke: Ashes, unless with Crystal of Good"),
                New ItemData("Crystal of Good", 1004, "Special", "", 0, "*", "Invoke: Ashes, unless with Crystal of Evil"),
                New ItemData("Ship in a Bottle", 1005, "Special", "", 0, "*", "Allows passage to level 4 & 5 stairs from level 1"),
                New ItemData("Staff of Earth", 1006, "Weapon", " Damage", 25000, "*", "Casts MANIFO"),
                New ItemData("Amulet of Air", 1007, "Magic", "", 25000, "*", "Casts DALTO"),
                New ItemData("Holy Water", 1008, "Misc", "", 25000, "*", "Casts DIAL"),
                New ItemData("Rod of Fire", 1009, "Magic", "", 25000, "*", "Casts MAHALITO"),
                New ItemData("Gold Medallion", 1010, "Special", "", 0, "*", "Can be traded for Holy Water"),
                New ItemData("Orb of Mhuuzfes", 1011, "Special", "", 0, "*", "False Orb; Piety -1; AC (20)"),
                New ItemData("Butterfly Knife", 1012, "Weapon", "17-35 Damage", 75000, "TBN", "Invoke: Class to Ninja"),
                New ItemData("Short Sword", 1013, "Weapon", "1-6 Damage", 30, "FTSLN", ""),
                New ItemData("Broad Sword", 1014, "Weapon", "1-8 Damage", 50, "FSLN", ""),
                New ItemData("Mace", 1015, "Weapon", "2-6 Damage", 60, "FPBSLN", ""),
                New ItemData("Staff of Earth", 1016, "Weapon", " Damage", 25000, "*", "Casts MANIFO"),
                New ItemData("Hand Axe", 1017, "Weapon", "1-4 Damage", 30, "FTSN", ""),
                New ItemData("Battle Axe", 1018, "Weapon", "3-8 Damage", 140, "FSN", ""),
                New ItemData("Dagger", 1019, "Weapon", "1-4 Damage", 10, "All", "AC +1"),
                New ItemData("Flail", 1020, "Weapon", "1-7 Damage", 300, "FPSLN", ""),
                New ItemData("Round Shield", 1021, "Shield", "AC 1", 40, "FPTBSLN", ""),
                New ItemData("Heater Shield", 1022, "Shield", "AC 2", 80, "FPSLN", ""),
                New ItemData("Mage's Robes", 1023, "Armor", "AC 1", 30, "All", ""),
                New ItemData("Cuirass", 1024, "Armor", "AC 2", 100, "FPTBSLN", ""),
                New ItemData("Hauberk", 1025, "Armor", "AC 3", 200, "FPBSLN", ""),
                New ItemData("Breast Plate", 1026, "Armor", "AC 4", 400, "FPSLN", ""),
                New ItemData("Plate Armor", 1027, "Armor", "AC 5", 1500, "FSLN", ""),
                New ItemData("Sallet", 1028, "Helmet", "AC 1", 200, "FSL", ""),
                New ItemData("Potion of DIOS", 1029, "Magic", "", 1000, "All", ""),
                New ItemData("LATUMOFIS Oil", 1030, "Magic", "", 600, "All", ""),
                New ItemData("Short Sword +1", 1031, "Weapon", "1-8 Damage", 10000, "FTSLN", ""),
                New ItemData("Broad Sword +1", 1032, "Weapon", "1-10 Damage", 10000, "FSLN", ""),
                New ItemData("Mace +1", 1033, "Weapon", "3-9 Damage", 10000, "FPBSLN", ""),
                New ItemData("Battle Axe +1", 1034, "Weapon", "3-10 Damage", 12500, "FSN", ""),
                New ItemData("Nunchakas", 1035, "Weapon", "1-10 Damage", 15000, "FPSN", ""),
                New ItemData("Dagger +1", 1036, "Weapon", "1-6 Damage", 10000, "All", ""),
                New ItemData("Scroll of KATINO", 1037, "Magic", "", 1000, "All", ""),
                New ItemData("Cuirass +1", 1038, "Armor", "AC 3", 3000, "FPTBSLN", ""),
                New ItemData("Hauberk +1", 1039, "Armor", "AC 4", 3500, "FPBSLN", ""),
                New ItemData("Breast Plate +1", 1040, "Armor", "AC 5", 4000, "FPSLN", ""),
                New ItemData("Plate Armor +1", 1041, "Armor", "AC 6", 5000, "FSLN", ""),
                New ItemData("Heater +1", 1042, "Shield", "AC 3", 2500, "FPSLN", ""),
                New ItemData("Bascinet", 1043, "Helmet", "AC 2", 1000, "FSL", ""),
                New ItemData("Gloves of Iron", 1044, "Gauntlets", "AC 1", 2500, "FPSL", ""),
                New ItemData("Scroll of BADIOS", 1045, "Magic", "", 1000, "All", ""),
                New ItemData("Potion of HALITO", 1046, "Magic", "", 1000, "All", ""),
                New ItemData("Short Sword -1", 1047, "Weapon", "1-3 Damage", 1000, "FSLN", "Cursed; AC -1"),
                New ItemData("Broadsword -1", 1048, "Weapon", "1-4 Damage", 1000, "FSLN", "Cursed; AC -1"),
                New ItemData("Mace -1", 1049, "Weapon", "1-3 Damage", 8000, "FPBSLN", "Cursed; AC -1"),
                New ItemData("Dagger -1", 1050, "Weapon", "1-3 Damage", 500, "All", "Cursed; AC -2"),
                New ItemData("Battle Axe -1", 1051, "Weapon", "1-4 Damage", 1000, "FSN", "Cursed; AC -1"),
                New ItemData("Margaux's Flail", 1052, "Weapon", "1-3 Damage", 1000, "FPTBSLN", ""),
                New ItemData("Bag of Gems", 1053, "Special", "", 100, "All", ""),
                New ItemData("Wizard's Staff", 1054, "Weapon", "2-8 Damage", 6000, "MBS", "Casts MOGREF"),
                New ItemData("Flametongue", 1055, "Weapon", "2-11 Damage", 15000, "FSL", "Casts HALITO"),
                New ItemData("Round Shield -1", 1056, "Shield", "AC ", 4000, "FPTBSLN", "Cursed;"),
                New ItemData("Cuirass -1", 1057, "Armor", "AC ", 2000, "FPTBSLN", "Cursed;"),
                New ItemData("Hauberk -1", 1058, "Armor", "AC ", 1000, "FPBSLN", "Cursed;"),
                New ItemData("Breast Plate -1", 1059, "Armor", "AC 2", 8000, "FPSLN", "Cursed;"),
                New ItemData("Plate Armor -1", 1060, "Armor", "AC ", 4000, "FSLN", "Cursed;"),
                New ItemData("Sallet -1", 1061, "Helmet", "AC -1", 1000, "FSL", "Cursed;"),
                New ItemData("Potion of SOPIC", 1062, "Magic", "", 2500, "All", ""),
                New ItemData("Gold Ring", 1063, "Special", "", 10000, "All", ""),
                New ItemData("Salamander Ring", 1064, "Special", "", 15000, "All", ""),
                New ItemData("Serpent's Tooth", 1065, "Special", "AC 1", 15000, "MPTB", ""),
                New ItemData("Short Sword +2", 1066, "Weapon", "2-10 Damage", 20000, "FTSLN", ""),
                New ItemData("Broad Sword +2", 1067, "Weapon", "2-12 Damage", 20000, "FSLN", ""),
                New ItemData("Battle Axe +2", 1068, "Weapon", "4-12 Damage", 20000, "FSN", ""),
                New ItemData("Ivory Blade (G)", 1069, "Weapon", "1-8 Damage", 15000, "FMTSL", ""),
                New ItemData("Ebony Blade (E)", 1070, "Weapon", "1-8 Damage", 15000, "FMTSN", ""),
                New ItemData("Amber Blade (N)", 1071, "Weapon", "1-8 Damage", 15000, "FMT", ""),
                New ItemData("Mace +2", 1072, "Weapon", "2-10 Damage", 20000, "FPBSLN", ""),
                New ItemData("Gloves of Mythril", 1073, "Gauntlets", "AC 2", 6000, "FSL", ""),
                New ItemData("Amulet of DIALKO", 1074, "Magic", "", 8000, "All", "Casts DIALKO"),
                New ItemData("Cuirass +2", 1075, "Armor", "AC 4", 6000, "FPTBSLN", ""),
                New ItemData("Heater +2", 1076, "Shield", "AC 4", 6000, "FPSLN", ""),
                New ItemData("Displacer Robes", 1077, "Armor", "AC 3", 12000, "All", ""),
                New ItemData("Hauberk +2", 1078, "Armor", "AC 5", 8000, "FPBSLN", ""),
                New ItemData("Breast Plate +2", 1079, "Armor", "AC 6", 10000, "FPSLN", ""),
                New ItemData("Plate Armor +2", 1080, "Armor", "AC 7", 14000, "FSLN", ""),
                New ItemData("Armet", 1081, "Armor", "AC 3", 8000, "FSL", ""),
                New ItemData("Wargan Robes", 1082, "Armor", "AC -2", 4000, "All", "Cursed;"),
                New ItemData("Giant's Club", 1083, "Weapon", "4-10 Damage", 20000, "FPSLN", ""),
                New ItemData("Blade Cusinart'", 1084, "Weapon", "2-7 Damage", 15000, "FSLN", ""),
                New ItemData("Shepherd Crook", 1085, "Weapon", "1-4 Damage", 20, "All", ""),
                New ItemData("Unholy Axe", 1086, "Weapon", "3-12 Damage", 22500, "FSN", ""),
                New ItemData("Rod of Death", 1087, "Weapon", " Damage", 17500, "All", "Casts MAKANITO"),
                New ItemData("Gem of Exorcism", 1088, "Special", "", 12000, "All", ""),
                New ItemData("Bag of Emeralds", 1089, "Special", "", 2000, "All", "Age -1"),
                New ItemData("Bag of Garnets", 1090, "Special", "", 1000, "All", "Strength -1"),
                New ItemData("Blue Pearl", 1091, "Special", "", 8000, "All", ""),
                New ItemData("Ruby Slippers", 1092, "Special", "", 16000, "All", "Casts LOKTOFEIT"),
                New ItemData("Necrology Rod", 1093, "Weapon", " Damage", 20000, "All", ""),
                New ItemData("Book of Life", 1094, "Misc", "", 25000, "All", "Casts DI"),
                New ItemData("Book of Death", 1095, "Misc", "", 0, "All", "Casts MABADI"),
                New ItemData("Dragon's Tooth", 1096, "Special", "AC 2", 30000, "All", ""),
                New ItemData("Trollkin Ring", 1097, "Special", "", 40000, "All", ""),
                New ItemData("Rabbit's Foot", 1098, "Special", "", 10000, "All", "Invoke: Luck +1"),
                New ItemData("Thief's Pick", 1099, "Special", "", 5000, "TN", "Invoke: Agility +1"),
                New ItemData("Book of Demons", 1100, "Misc", "", 50000, "All", "Invoke: Piety +1"),
                New ItemData("Butterfly Knife", 1101, "Weapon", "17-35 Damage", 75000, "TBN", "Invoke: Class to Ninja"),
                New ItemData("Gold Tiara", 1102, "Special", "2", 100000, "All", ""),
                New ItemData("Mantis Gloves", 1103, "Gauntlets", "AC 3", 15000, "FSLP", "")
               }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataPath As String
        Get
            Return "Scenario03DataPath"
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioDataOffset As String
        Get
            Return &H1DA00
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioName As String
        Get
            Return "THE LEGACY OF LLYLGAMYN"
        End Get
    End Property
End Class