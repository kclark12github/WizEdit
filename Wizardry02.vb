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
    Public Sub New(FileName As String, ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(FileName, Caption, Icon, BoxArt, Parent)
        Read()
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
                    New ItemData("Misc: Broken Item", 0),
                    New ItemData("Weapon: Long Sword", 1),
                    New ItemData("Weapon: Short Sword", 2),
                    New ItemData("Weapon: Anointed Mace", 3),
                    New ItemData("Weapon: Anointed Flail", 4),
                    New ItemData("Weapon: Staff", 5),
                    New ItemData("Weapon: Dagger", 6),
                    New ItemData("Shield: Small Shield", 7),
                    New ItemData("Shield: Large Shield", 8),
                    New ItemData("Armor: Robes", 9),
                    New ItemData("Armor: Leather Armor", 10),
                    New ItemData("Armor: Chain Mail", 11),
                    New ItemData("Armor: Breast Plate", 12),
                    New ItemData("Armor: Plate Mail", 13),
                    New ItemData("Helm: Helm", 14),
                    New ItemData("Magic: Potion Of Dios", 15),
                    New ItemData("Magic: Potion Of Latumofis", 16),
                    New ItemData("Weapon: Long Sword +1", 17),
                    New ItemData("Weapon: Short Sword +1", 18),
                    New ItemData("Weapon: Mace +1", 19),
                    New ItemData("Weapon: Staff of Mogref", 20),
                    New ItemData("Magic: Scroll of Katino", 21),
                    New ItemData("Armor: Leather +1", 22),
                    New ItemData("Armor: Chain Mail +1", 23),
                    New ItemData("Armor: Plate Mail +1", 24),
                    New ItemData("Shield: Shield +1", 25),
                    New ItemData("Armor: Breast Plate +1", 26),
                    New ItemData("Magic: Scroll Of Badios", 27),
                    New ItemData("Magic: Scroll Of Halito", 28),
                    New ItemData("Weapon: Long Sword -1", 29),
                    New ItemData("Weapon: Short Sword -1", 30),
                    New ItemData("Weapon: Mace -1", 31),
                    New ItemData("Weapon: Staff +2", 32),
                    New ItemData("Weapon: Dragon Slayer", 33),
                    New ItemData("Helm: Helm +1", 34),
                    New ItemData("Armor: Leather -1", 35),
                    New ItemData("Armor: Chain -1", 36),
                    New ItemData("Armor: Breast Plate -1", 37),
                    New ItemData("Shield: Shield -1", 38),
                    New ItemData("Magic: Jeweled Amulet", 39),
                    New ItemData("Magic: Scroll of Badios", 40),
                    New ItemData("Magic: Potion of Sopic", 41),
                    New ItemData("Weapon: Long Sword +2", 42),
                    New ItemData("Weapon: Short Sword +2", 43),
                    New ItemData("Weapon: Mace +2", 44),
                    New ItemData("Magic: Scroll Of Lomilwa", 45),
                    New ItemData("Magic: Scroll Of Dilto", 46),
                    New ItemData("Gauntlets: Copper Gloves", 47),
                    New ItemData("Armor: Leather +2", 48),
                    New ItemData("Armor: Chain +2", 49),
                    New ItemData("Armor: Plate Mail +2", 50),
                    New ItemData("Shield: Shield +2", 51),
                    New ItemData("Helm: Helm +2 (E)", 52),
                    New ItemData("Magic: Potion Of Dial", 53),
                    New ItemData("Magic: Ring of Porfic", 54),
                    New ItemData("Weapon: Were Slayer", 55),
                    New ItemData("Weapon: Mage Masher", 56),
                    New ItemData("Weapon: Mace Pro Poison", 57),
                    New ItemData("Weapon: Staff Of Montino", 58),
                    New ItemData("Weapon: Blade Cusinart'", 59),
                    New ItemData("Magic: Amulet Of Manifo", 60),
                    New ItemData("Weapon: Rod Of Flame", 61),
                    New ItemData("Armor: Chain +2 (E)", 62),
                    New ItemData("Armor: Plate +2 (N)", 63),
                    New ItemData("Shield: Shield +3 (E)", 64),
                    New ItemData("Magic: Amulet Of Makanito", 65),
                    New ItemData("Helm: Helm of Malor", 66),
                    New ItemData("Magic: Scroll of Badial", 67),
                    New ItemData("Weapon: Short Sword -2", 68),
                    New ItemData("Weapon: Dagger +2", 69),
                    New ItemData("Weapon: Mace -2", 70),
                    New ItemData("Weapon: Staff -2", 71),
                    New ItemData("Weapon: Dagger Of Speed", 72),
                    New ItemData("Armor: Cursed Robe", 73),
                    New ItemData("Armor: Leather -2", 74),
                    New ItemData("Armor: Chain -2", 75),
                    New ItemData("Armor: Breastplate -2", 76),
                    New ItemData("Shield: Shield -2", 77),
                    New ItemData("Helm: Cursed Helmet", 78),
                    New ItemData("Armor: Breast Plate +2", 79),
                    New ItemData("Gauntlets: Gloves of Silver", 80),
                    New ItemData("Weapon: Evil +3 Sword", 81),
                    New ItemData("Weapon: +3 Evil Short Sword", 82),
                    New ItemData("Weapon: Thieves Dagger", 83),
                    New ItemData("Armor: +3 Breast Plate", 84),
                    New ItemData("Armor: Lord's Garb", 85),
                    New ItemData("Weapon: Muramasa Blade", 86),
                    New ItemData("Weapon: Shiriken", 87),
                    New ItemData("Armor: Chain Pro Fire", 88),
                    New ItemData("Armor: +3 Evil Plate", 89),
                    New ItemData("Shield: +3 Shield", 90),
                    New ItemData("Magic: Ring of Healing", 91),
                    New ItemData("Magic: Ring Pro Undead", 92),
                    New ItemData("Magic: Deadly Ring", 93),
                    New ItemData("Magic: Rod of Raising", 94),
                    New ItemData("Magic: Amulet of Cover", 95),
                    New ItemData("Armor: +3 Robe", 96),
                    New ItemData("Gauntlets: Winter Mittens", 97),
                    New ItemData("Magic: Necklace Pro Magic", 98),
                    New ItemData("Weapon: Staff of Light", 99),
                    New ItemData("Weapon: +5 Long Sword", 100),
                    New ItemData("Weapon: Sword of Swinging", 101),
                    New ItemData("Weapon: Priest Puncher", 102),
                    New ItemData("Weapon: Priest Mace", 103),
                    New ItemData("Weapon: Short Sword of Swinging", 104),
                    New ItemData("Magic: Ring Pro Fire", 105),
                    New ItemData("Armor: Cursed +1 Plate", 106),
                    New ItemData("Armor: +5 Plate", 107),
                    New ItemData("Weapon: Staff of Curing", 108),
                    New ItemData("Magic: Ring of Regen", 109),
                    New ItemData("Magic: Metamorph Ring", 110),
                    New ItemData("Misc: Stone Stone", 111),
                    New ItemData("Misc: Dreamer's Stone", 112),
                    New ItemData("Misc: Damien Stone", 113),
                    New ItemData("Weapon: Great Mage Wand", 114),
                    New ItemData("Misc: Coin of Power", 115),
                    New ItemData("Misc: Stone of Youth", 116),
                    New ItemData("Misc: Mind Stone", 117),
                    New ItemData("Misc: Stone of Piety", 118),
                    New ItemData("Misc: Blarney Stone", 119),
                    New ItemData("Magic: Amulet of Skill", 120),
                    New ItemData("Magic: Amulet of Skill", 121),
                    New ItemData("Weapon: Great Mage Wand", 122),
                    New ItemData("Magic: Coin of Power", 123),
                    New ItemData("Weapon: Staff of Gnilda", 124),
                    New ItemData("Special: Hrathnir", 125),
                    New ItemData("Special: KOD's Helm", 126),
                    New ItemData("Special: KOD's Shield", 127),
                    New ItemData("Special: KOD's Gauntlets", 128),
                    New ItemData("Special: KOD's Armor", 129)
                }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataDirectory As String
        Get
            Return "UWAPath02"
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataFile As String
        Get
            Return "Wiz02DataFile"
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