'Wizardry04.cls
'   Main Class for The Return of Werdna...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/19    Ken Clark       Migrated to VS2017;
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
'WERDNA IV
'0x4AC00	Name
'0x4AC10 Password?
'0x4AC20	Out[I2]
'0x4AC22 Race[I2]
'0x4AC24 Profession[I2]
'0x4AC26 Age[I2]
'0x4AC28 Status[I2]
'0x4AC2A Alignment[I2]
'0x4AC2C Statistics[I4];
'0x4AC20 Unknown[I4];
'0x4AC24 Gold[I6];
'0x4AC3A ItemCount[I2];

'	0x4AC3C Equipped1[I2];
'	0x4AC3E Cursed1[I2];
'	0x4AC40 Identified1[I2];
'	0x4AC42 ItemCode1[I2];

'	0x4AC44 Equipped2[I2];
'	0x4AC46 Cursed2[I2];
'	0x4AC48 Identified2[I2];
'	0x4AC4A ItemCode2[I2];	

'	0x4AC4C Equipped3[I2];
'	0x4AC4E Cursed3[I2];
'	0x4AC50 Identified3[I2];
'	0x4AC52 ItemCode3[I2];	

'	0x4AC54 Equipped4[I2];
'	0x4AC56 Cursed4[I2];
'	0x4AC58 Identified4[I2];
'	0x4AC5A ItemCode4[I2];	

'	0x4AC5C Equipped5[I2];
'	0x4AC5E Cursed5[I2];
'	0x4AC60 Identified5[I2];
'	0x4AC62 ItemCode5[I2];

'	0x4AC64 Equipped6[I2];
'	0x4AC66 Cursed6[I2];
'	0x4AC68 Identified6[I2];
'	0x4AC6A ItemCode6[I2];	

'	0x4AC6C Equipped7[I2];
'	0x4AC6E Cursed7[I2];
'	0x4AC70 Identified7[I2];
'	0x4AC72 ItemCode7[I2];	

'	0x4AC74 Equipped8[I2];
'	0x4AC76 Cursed8[I2];
'	0x4AC78 Identified8[I2];
'	0x4AC7A ItemCode8[I2];	
'0x4AC7C Keys[I6]; [I12]?
'0x4AC82 Level.Current[I2];.Max[I2];
'0x4AC86 HP.Current[I2];.Max[I2];
'0x4AC8A SpellBooks[8B];
'0x4AC92 SpellPoints[I2]x14
'0x4ACAE Unknown[[I2]x14 (Coincidence?) - Spells Cast maybe?
'0x4ACCA Location[I2];
'0x4ACCE	Down[I2] - Screen says -2 but data says 0E 00
'0x4ACD0 Group1.Count[I2]
'0x4ACD2 Group2.Count[I2]
'0x4ACD4 Group3.Count[I2]
'0x4ACD6 Group1.Code[I2]
'04AACD8 Group2.Code[I2]
'0x4ACDA Group3.Code[I2]
'0x4ACDC Group1.Name ("A DINK") String(15) 0x00
'0x4ACEC Group2.Name ("ENTELECHY FUFF") String(15) 0x77
'0x4ACFC Group3.Name ("VAMPIRE LORD") String(15) 0x70
'0x4ACFC 
Option Explicit On

Public Class Wizardry04
    Inherits WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(Caption, Icon, BoxArt, Parent)
    End Sub
    Public Overrides ReadOnly Property CharacterDataOffset As Int32
        Get
            Return &H4AC00
        End Get
    End Property
    Public Overrides ReadOnly Property CharactersMax As Short
        Get
            Return 8
        End Get
    End Property
    Public Overrides ReadOnly Property MasterItemList As ItemData()
        Get
            If mMasterItemList Is Nothing Then
                mMasterItemList = {
                    New ItemData("Misc: Broken Item", 0),
                    New ItemData("Bloodstone", 1),
                    New ItemData("Lander's Turq.", 2),
                    New ItemData("Amber Dragon", 3),
                    New ItemData("HHG of Aunty Ock", 4),
                    New ItemData("Winged Boots", 5),
                    New ItemData("Dreampainter Ka", 6),
                    New ItemData("East Wind Sword", 7),
                    New ItemData("West Wind Sword", 8),
                    New ItemData("Dragon's Claw", 9),
                    New ItemData("Hopalong Carrot", 10),
                    New ItemData("Cleansing Oil", 11),
                    New ItemData("Witching Rod", 12),
                    New ItemData("Aromatic Ball", 13),
                    New ItemData("Void Transducer", 14),
                    New ItemData("Kris of Truth", 15),
                    New ItemData("Inn Key", 16),
                    New ItemData("Crystal Rose", 17),
                    New ItemData("Dab of Puce", 18),
                    New ItemData("Pennoncaeux", 19),
                    New ItemData("Maintenance Cap", 20),
                    New ItemData("Long Sword", 21),
                    New ItemData("Short Sword", 22),
                    New ItemData("Anointed Mace", 23),
                    New ItemData("Anointed Flail", 24),
                    New ItemData("Staff", 25),
                    New ItemData("Dagger", 26),
                    New ItemData("Small Shield", 27),
                    New ItemData("Large Shield", 28),
                    New ItemData("Robes", 29),
                    New ItemData("Leather Armor", 30),
                    New ItemData("Chain Mail", 31),
                    New ItemData("Breast Plate", 32),
                    New ItemData("Plate Mail", 33),
                    New ItemData("Helm", 34),
                    New ItemData("Potion of Dios", 35),
                    New ItemData("Potion of Porfic", 36),
                    New ItemData("Long Sword+1", 37),
                    New ItemData("Short Sword+1", 38),
                    New ItemData("Mace+1", 39),
                    New ItemData("Staff of Mogref", 40),
                    New ItemData("Scroll of Katino", 41),
                    New ItemData("Leather+1", 42),
                    New ItemData("Chain Mail+1", 43),
                    New ItemData("Plate Mail+1", 44),
                    New ItemData("Shield+1", 45),
                    New ItemData("St. K.A.'s Foot", 46),
                    New ItemData("Scroll of Badios", 47),
                    New ItemData("Scroll of Halito", 48),
                    New ItemData("Staff+2", 49),
                    New ItemData("Dragon Slayer", 50),
                    New ItemData("Helm+1", 51),
                    New ItemData("Jeweled Amulet", 52),
                    New ItemData("Scroll of Badial", 53),
                    New ItemData("Potion of Sopic", 54),
                    New ItemData("Long Sword+2", 55),
                    New ItemData("Good Hope Cape", 56),
                    New ItemData("Magician's Hat", 57),
                    New ItemData("Novice's Cap", 58),
                    New ItemData("Scroll of Dilto", 59),
                    New ItemData("Copper Gloves", 60),
                    New ItemData("Initiate Turban", 61),
                    New ItemData("Wizard Skullcap", 62),
                    New ItemData("Plate Mail+2", 63),
                    New ItemData("Shield+2", 64),
                    New ItemData("Mordorcharge", 65),
                    New ItemData("Potion of Dial", 66),
                    New ItemData("Ring of Porfic", 67),
                    New ItemData("Were Slayer", 68),
                    New ItemData("Mage Masher", 69),
                    New ItemData("Mace of Curing", 70),
                    New ItemData("Staff of Montino", 71),
                    New ItemData("Blade Cusinart'", 72),
                    New ItemData("Amulet of Badialma", 73),
                    New ItemData("Rod of Flame", 74),
                    New ItemData("Cape of Hide", 75),
                    New ItemData("Cape of Jackel", 76),
                    New ItemData("Cape of Hide", 77),
                    New ItemData("Amulet of Makanito", 78),
                    New ItemData("Diadem of Malor", 79),
                    New ItemData("Scroll of Badial", 80),
                    New ItemData("Dagger+2", 81),
                    New ItemData("Dagger of Speed", 82),
                    New ItemData("Lich's Robes", 83),
                    New ItemData("Skull's Cap", 84),
                    New ItemData("Potion of Masopic", 85),
                    New ItemData("Silver Gloves", 86),
                    New ItemData("GetOut of JailFree", 87),
                    New ItemData("Gold Pyrite", 88),
                    New ItemData("Oxygen Mask", 89),
                    New ItemData("Chronicles of H", 90),
                    New ItemData("Lord's Garb", 91),
                    New ItemData("Murasama Blade", 92),
                    New ItemData("Shuriken", 93),
                    New ItemData("Chain Pro Ice", 94),
                    New ItemData("**ERR**", 95),
                    New ItemData("**ERR**", 96),
                    New ItemData("Ring of Healing", 97),
                    New ItemData("Ring of Dispelling", 98),
                    New ItemData("Ring of Death", 99),
                    New ItemData("Adept Baldness", 100),
                    New ItemData("Arabic Diary", 101),
                    New ItemData("Demonic Chimes", 102),
                    New ItemData("Black Candle", 103),
                    New ItemData("Black Box", 104),
                    New ItemData("St. Trebor Rump", 105),
                    New ItemData("Bish's Tongue", 106),
                    New ItemData("St. Rimbo Digit", 107),
                    New ItemData("Arrow of Truth", 108),
                    New ItemData("Orb of Dreams", 109),
                    New ItemData("Rallying Horn", 110),
                    New ItemData("Signet Ring", 111),
                    New ItemData("Mythril Glove", 112),
                    New ItemData("Holy Limp Wrist", 113),
                    New ItemData("Twilight Cloak", 114),
                    New ItemData("Shadow Cloak", 115),
                    New ItemData("Cone of Silence", 116),
                    New ItemData("Darkness Cloak", 117),
                    New ItemData("Night Cloak", 118),
                    New ItemData("Entropy Cloak", 119)
                }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataPath As String
        Get
            Return "Scenario04DataPath"
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioDataOffset As String
        Get
            Return &H4BC00
        End Get
    End Property
    Public Overrides ReadOnly Property ScenarioName As String
        Get
            Return "THE RETURN OF WERDNA"
        End Get
    End Property
End Class