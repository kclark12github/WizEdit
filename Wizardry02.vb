'Wizardry02.cls
'   Main Class for Knight of Diamonds (KOD)...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/19    Ken Clark       Migrated to VS2017;
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class Wizardry02
    Inherits WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(Caption, Icon, BoxArt, Parent)
    End Sub
    Public Overrides ReadOnly Property CharacterDataOffset As Int32
        Get
            Return &H1D200
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
                    New ItemData("Broken Item", 0, "Misc", " ", 0, "*", "None (The item you get when you use a scroll or invoke a special power of an object and the object breaks) "),
                    New ItemData("Long Sword", 1, "Weapon", "1-8 Damage", 25, "FSLN", ""),
                    New ItemData("Short Sword", 2, "Weapon", "1-6 Damage", 15, "FTSLN", ""),
                    New ItemData("Anointed Mace", 3, "Weapon", "2-6 Damage", 30, "FPBSLN", ""),
                    New ItemData("Anointed Flail", 4, "Weapon", "1-7 Damage", 150, "FPSLN", ""),
                    New ItemData("Staff", 5, "Weapon", "1-5 Damage", 10, " *", ""),
                    New ItemData("Dagger", 6, "Weapon", "1-4 Damage", 5, "FMTSLN", ""),
                    New ItemData("Small Shield", 7, "Shield", "AC 2", 20, "FPTBSLN", ""),
                    New ItemData("Large Shield", 8, "Shield", "AC 3", 40, "FPSLN", ""),
                    New ItemData("Robes", 9, "Armor", "AC 1", 15, "*", ""),
                    New ItemData("Leather Armor", 10, "Armor", "AC 2", 50, "FPTBSLN", ""),
                    New ItemData("Chain Mail", 11, "Armor", "AC 3", 90, "FPSLN", ""),
                    New ItemData("Breast Plate", 12, "Armor", "AC 4", 200, "FPSLN", ""),
                    New ItemData("Plate Mail", 13, "Armor", "AC 5", 750, "FSLN", ""),
                    New ItemData("Helm", 14, "Helmet", "AC 1", 100, "FSLN", ""),
                    New ItemData("Potion of DIOS", 15, "Magic", "", 500, "*", "Casts DIOS"),
                    New ItemData("Potion of LATUMOFIS", 16, "Magic", "", 300, "*", "Casts LATUMOFIS"),
                    New ItemData("Long Sword +1", 17, "Weapon", "2-11 Damage", 10000, "FSLN", ""),
                    New ItemData("Short Sword +1", 18, "Weapon", "2-6 Damage", 15000, "FTSLN", ""),
                    New ItemData("Mace +1", 19, "Weapon", "3-9 Damage", 12500, "FPBSLN", ""),
                    New ItemData("Staff of MOGREF", 20, "Weapon", "1-6 Damage", 3000, "MB", "Casts MOGREF"),
                    New ItemData("Scroll of KATINO", 21, "Magic", "", 500, "*", "Casts KATINO"),
                    New ItemData("Leather +1", 22, "Armor", "AC 3", 1500, "FPTBSLN", ""),
                    New ItemData("Chain Mail +1", 23, "Armor", "AC 4", 1500, "FPSLN", ""),
                    New ItemData("Plate Mail +1", 24, "Armor", "AC 6", 1500, "FSLN", ""),
                    New ItemData("Shield +1", 25, "Shield", "AC 4", 1500, "FPTSLN", ""),
                    New ItemData("Breast Plate +1", 26, "Armor", "AC 5", 1500, "FPSLN", ""),
                    New ItemData("Scroll of BADIOS", 27, "Magic", "", 500, "*", "Casts BADIOS"),
                    New ItemData("Scroll of HALITO", 28, "Magic", "", 500, "*", "Casts HALITO"),
                    New ItemData("Long Sword -1", 29, "Weapon", "1-8 Damage", 1000, "FSLN", "Cursed;"),
                    New ItemData("Short Sword -1", 30, "Weapon", "1-6 Damage", 1000, "FTSLN", "Cursed;"),
                    New ItemData("Mace -1", 31, "Weapon", "2-6 Damage", 1000, "FPBSLN", "Cursed;"),
                    New ItemData("Staff +2", 32, "Weapon", "3-6 Damage", 2500, "All", ""),
                    New ItemData("Dragon Slayer", 33, "Weapon", "2-9 Damage", 10000, "FSLN", "Prot and Vs: Dragon"),
                    New ItemData("Helm +1", 34, "Helmet", "AC 2", 3000, "FSLN", ""),
                    New ItemData("Leather -1", 35, "Armor", "AC 1", 1500, "FPTBSL", "Cursed;"),
                    New ItemData("Chain -1", 36, "Armor", "AC 2", 1500, "FPSLN", "Cursed;"),
                    New ItemData("Breast Plate -1", 37, "Armor", "AC 3", 1500, "FPSLN", "Cursed;"),
                    New ItemData("Shield -1", 38, "Shield", "AC -1", 1500, "FPTSL", "Cursed;"),
                    New ItemData("Jeweled Amulet", 39, "Magic", "", 5000, "*", "Casts DUMAPIC"),
                    New ItemData("Scroll of BADIOS", 40, "Magic", "", 500, "*", "Casts BADIOS"),
                    New ItemData("Potion of SOPIC", 41, "Magic", "", 1500, "*", "Casts SOPIC"),
                    New ItemData("Long Sword +2", 42, "Weapon", "3-10 Damage", 4000, "FSLN", ""),
                    New ItemData("Short Sword +2", 43, "Weapon", "3-12 Damage", 4000, "FTSLN", ""),
                    New ItemData("Mace +2", 44, "Weapon", "3-8 Damage", 4000, "FPBSLN", ""),
                    New ItemData("Scroll of LOMILWA", 45, "Magic", "", 2500, "*", "Casts LOMILWA"),
                    New ItemData("Scroll of DILTO", 46, "Magic", "", 2500, "*", "Casts DILTO"),
                    New ItemData("Copper Gloves", 47, "Gauntlets", "AC 1", 6000, "FSLN", ""),
                    New ItemData("Leather +2", 48, "Armor", "AC 4", 6000, "FPTBSLN", ""),
                    New ItemData("Chain +2", 49, "Armor", "AC 5", 6000, "FPSLN", ""),
                    New ItemData("Plate Mail +2", 50, "Armor", "AC 7", 6000, "FPSLN", ""),
                    New ItemData("Shield +2", 51, "Shield", "AC 5", 7000, "FPTSLN", ""),
                    New ItemData("Helm +2 (E)", 52, "Helmet", "AC 3", 8000, "FSLN", "Alig: Evil; Casts BADIOS"),
                    New ItemData("Potion of DIAL", 53, "Magic", "", 5000, "*", "Casts DIAL"),
                    New ItemData("Ring of PORFIC", 54, "Magic", "", 10000, "*", "Casts PORFIC"),
                    New ItemData("Were Slayer", 55, "Weapon", "1-8 Damage", 10000, "FSLN", "Prot and Vs: Were"),
                    New ItemData("Mage Masher", 56, "Weapon", "2-7 Damage", 10000, "FTSLN", "Prot: Mage"),
                    New ItemData("Mace Pro Poison", 57, "Weapon", "2-11 Damage", 10000, "FPBSLN", "Prot: Insect; Res: Poison"),
                    New ItemData("Staff of MONTINO", 58, "Weapon", "10-12 Damage", 15000, "All", "Casts MONTINO"),
                    New ItemData("Blade Cusinart'", 59, "Weapon", "2-7 Damage", 15000, "FSLN", ""),
                    New ItemData("Amulet of MANIFO", 60, "Magic", "", 15000, "P", "Casts MANIFO"),
                    New ItemData("Rod of Flame", 61, "Magic", "", 25000, "MBS", "Prot: Fire; Casts MAHALITO"),
                    New ItemData("Chain +2 (E)", 62, "Armor", "AC 5", 8000, "FPSLN", "Alignment: Evil"),
                    New ItemData("Plate +2 (N)", 63, "Armor", "AC 7", 8000, "FPSLN", "Alignment: Neutral"),
                    New ItemData("Shield +3 (E)", 64, "Shield", "AC 5", 25000, "FPTSLN", "Alignment: Evil"),
                    New ItemData("Amulet of MAKANITO", 65, "Magic", "", 20000, "*", "Casts MAKANITO"),
                    New ItemData("Helm of MALOR", 66, "Helmet", "AC 2", 25000, "*", "Casts MALOR"),
                    New ItemData("Scroll of BADIAL", 67, "Magic", "", 8000, "*", "Casts BADIAL"),
                    New ItemData("Short Sword -2", 68, "Weapon", "1-6 Damage", 8000, "FTSLN", ""),
                    New ItemData("Dagger +2", 69, "Weapon", "3-6 Damage", 8000, "FMTSLN", ""),
                    New ItemData("Mace -2", 70, "Weapon", "1-8 Damage", 2500, "FPBSLN", "Cursed;"),
                    New ItemData("Staff -2", 71, "Weapon", "1-4 Damage", 8000, "*", "Cursed;"),
                    New ItemData("Dagger of Speed", 72, "Weapon", "1-4 Damage", 30000, "MN", "AC: 3"),
                    New ItemData("Cursed Robe", 73, "Armor", "AC -2", 8000, "*", "Cursed;"),
                    New ItemData("Leather -2", 74, "Armor", "AC ", 8000, "FPTBSLN", "Cursed;"),
                    New ItemData("Chain -2", 75, "Armor", "AC 1", 8000, "FPSLN", "Cursed;"),
                    New ItemData("Breast Plate -2", 76, "Armor", "AC 2", 8000, "FPSLN", "Cursed;"),
                    New ItemData("Shield -2", 77, "Shield", "AC ", 8000, "FPTSLN", "Cursed;"),
                    New ItemData("Cursed Helmet", 78, "Helmet", "AC -2", 50000, "FSLN", ""),
                    New ItemData("Breast Plate +2", 79, "Armor", "AC 6", 10000, "FPSLN", ""),
                    New ItemData("Gloves of Silver", 80, "Gauntlets", "AC 3", 60000, "FSLN", ""),
                    New ItemData("Evil +3 Sword", 81, "Weapon", "4-13 Damage", 50000, "FSLN", ""),
                    New ItemData("Evil Short Sword +3", 82, "Weapon", " Damage",, "", ""),
                    New ItemData("Thieves Dagger", 83, "Weapon", "11-16 Damage", 50000, "TN", "Invoke: Class to Ninja"),
                    New ItemData("Breast Plate +3", 84, "Armor", "AC 7", 100000, "FPSLN", ""),
                    New ItemData("Lord's Garb", 85, "Armor", "AC 10", 1000000, "L", "Prot: Mythical, Dragon; Regeneration (1); Vs: Were, Demon, Undead"),
                    New ItemData("Murasama Blade", 86, "Weapon", "10-50 Damage", 1000000, "S", "Invoking: St+1"),
                    New ItemData("Shuriken", 87, "Weapon", "1-6 Damage", 50000, "N", "Alig: Evil; Res: Poison, LvlDrain; Invoking: Hp+1."),
                    New ItemData("Chain Pro Fire", 88, "Armor", "AC 6", 150000, "FPSLN", ""),
                    New ItemData("Evil Plate +3", 89, "Armor", "AC 9", 150000, "FPSLN", "Alignment: Evil"),
                    New ItemData("Shield +3", 90, "Shield", "AC 6", 250000, "FPTSLN", ""),
                    New ItemData("Ring of Healing", 91, "Magic", "", 300000, "*", "Regeneration(1)"),
                    New ItemData("Ring Pro Undead", 92, "Magic", "", 500000, "*", "Prot: Undead"),
                    New ItemData("Deadly Ring", 93, "Magic", "", 500000, "*", "Cursed; Regeneration(1)"),
                    New ItemData("Rod of Raising", 94, "Weapon", "3-24 Damage", 0, "*", "Casts KADORTO"),
                    New ItemData("Amulet of Cover", 95, "Magic", "AC 3", 120000, "*", ""),
                    New ItemData("Robe +3", 96, "Armor", "AC 4", 180000, "M", ""),
                    New ItemData("Winter Mittens", 97, "Gauntlets", "AC 3", 138344, "FSLN", ""),
                    New ItemData("Necklace Pro Magic", 98, "Magic", "", 0, "*", "Prot: Magic"),
                    New ItemData("Staff of Light", 99, "Weapon", "4-18 Damage", 60000, "*", "Casts LOMILWA"),
                    New ItemData("Long Sword +5", 100, "Weapon", "11-18 Damage", 70000, "FTSLN", ""),
                    New ItemData("Sword of Swinging", 101, "Weapon", "1-8 Damage", 0, "FTSLN", ""),
                    New ItemData("Priest Puncher", 102, "Weapon", "2-16 Damage", 70000, "FTSLN", ""),
                    New ItemData("Priest Mace", 103, "Weapon", "2-16 Damage", 75000, "PB", ""),
                    New ItemData("Short Sword of Swinging", 104, "Weapon", "2-6 Damage", 74675, "FTSLN", ""),
                    New ItemData("Ring Pro Fire", 105, "Magic", "", 250000, "*", "Prot: Fire"),
                    New ItemData("Cursed +1 Plate", 106, "Armor", "AC 6", 0, "FPSLN", "Cursed;"),
                    New ItemData("Plate +5", 107, "Armor", "AC 10", 275345, "FPSLN", ""),
                    New ItemData("Staff of Curing", 108, "Weapon", "4-11 Damage", 100000, "P", ""),
                    New ItemData("Ring of Regen", 109, "Magic", "", 100000, "*", "Regeneration (1)"),
                    New ItemData("Metamorph Ring", 110, "Magic", "", 0, "*", "Invoke: Change to Advanced Class"),
                    New ItemData("Stone (Granite) Stone", 111, "Misc", "", 0, "*", "Casts MONTINO"),
                    New ItemData("Dreamer's Stone", 112, "Misc", "", 0, "*", "Casts KATINO"),
                    New ItemData("Damien Stone", 113, "Misc", "", 0, "*", "Invoke: Try it and see..."),
                    New ItemData("Great Mage Wand", 114, "Magic", "", 0, "*", "Invoke: 9 Spells in all Levels"),
                    New ItemData("Coin of Power", 115, "Misc", "", 0, "*", "Invoke: Change to Advanced Class"),
                    New ItemData("Stone of Youth", 116, "Misc", "", 0, "*", "Invoke: Age -1"),
                    New ItemData("Mind Stone", 117, "Misc", "", 0, "*", "Invoke: I.Q. +1"),
                    New ItemData("Stone of Piety", 118, "Misc", "", 0, "*", "Invoke: Piety +1"),
                    New ItemData("Blarney Stone", 119, "Misc", "", 0, "*", "Invoke: Luck +1"),
                    New ItemData("Amulet of Skill", 120, "Magic", "", 0, "*", "Exp +50000"),
                    New ItemData("Amulet of Skill", 121, "Magic", "", 0, "*", "Exp +50000"),
                    New ItemData("Great Mage Wand", 122, "Magic", "", 0, "*", "Invoke: 9 Spells in all Levels"),
                    New ItemData("Coin of Power", 123, "Magic", "", 0, "*", "Invoke: Change to Advanced Class"),
                    New ItemData("Staff of Gnilda", 124, "Special", "AC 21", 0, "", ""),
                    New ItemData("Hrathnir", 125, "Special", "12-30 Damage", 0, "FSL", "Casts LORTO"),
                    New ItemData("KOD's Helm", 126, "Special", "AC 4", 0, "", "Casts MADALTO"),
                    New ItemData("KOD's Shield", 127, "Special", "AC 6", 0, "", "Casts DIALMA"),
                    New ItemData("KOD's Gauntlets", 128, "Special", "AC 4", 0, "", "Casts TILTOWAIT"),
                    New ItemData("KOD's Armor", 129, "Special", "AC 14", 0, "", "Casts MATU")
                }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataPath As String
        Get
            Return "Scenario02DataPath"
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioDataOffset As String
        Get
            Return &H1CE00
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioName As String
        Get
            Return "THE KNIGHT OF DIAMONDS"
        End Get
    End Property
    Public Overrides ReadOnly Property HonorsList As String()
        Get
            HonorsList = {
                "> Chevron of Trebor",
                "K - Knight of Gnilda"
       }
        End Get
    End Property
End Class