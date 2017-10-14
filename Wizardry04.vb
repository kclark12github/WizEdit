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
                    New ItemData("Special: Bloodstone", 1),
                    New ItemData("Special: Lander's Turq.", 2),
                    New ItemData("Special: Amber Dragon", 3),
                    New ItemData("Weapon: Holy Hand Grenade of Aunty Ock", 4),
                    New ItemData("Special: Winged Boots", 5),
                    New ItemData("Magic: Dreampainter Ka (casts MADI)", 6),
                    New ItemData("Weapon: East Wind Sword", 7),
                    New ItemData("Weapon: West Wind Sword", 8),
                    New ItemData("Weapon: Dragon's Claw", 9),
                    New ItemData("Weapon: Hopalong Carrot", 10),
                    New ItemData("Special: Cleansing Oil", 11),
                    New ItemData("Magic: Witching Rod (casts KANDI)", 12),
                    New ItemData("Special: Aromatic Ball", 13),
                    New ItemData("Special: Void Transducer", 14),
                    New ItemData("Weapon: Kris of Truth", 15),
                    New ItemData("Special: Inn Key", 16),
                    New ItemData("Special: Crystal Rose", 17),
                    New ItemData("Special: Dab of Puce", 18),
                    New ItemData("Special: Pennoncaeux", 19),
                    New ItemData("Helmet: Maintenance Cap (casts DIAL)", 20),
                    New ItemData("Weapon: Long Sword", 21),
                    New ItemData("Weapon: Short Sword", 22),
                    New ItemData("Weapon: Anointed Mace", 23),
                    New ItemData("Weapon: Anointed Flail", 24),
                    New ItemData("Weapon: Staff", 25),
                    New ItemData("Weapon: Dagger", 26),
                    New ItemData("Shield: Small Shield", 27),
                    New ItemData("Shield: Large Shield", 28),
                    New ItemData("Armor: Robes", 29),
                    New ItemData("Armor: Leather Armor", 30),
                    New ItemData("Armor: Chain Mail", 31),
                    New ItemData("Armor: Breast Plate", 32),
                    New ItemData("Armor: Plate Mail", 33),
                    New ItemData("Helmet: Helm", 34),
                    New ItemData("Magic: Potion of DIOS", 35),
                    New ItemData("Magic: Potion of PORFIC", 36),
                    New ItemData("Weapon: Long Sword+1", 37),
                    New ItemData("Weapon: Short Sword+1", 38),
                    New ItemData("Weapon: Mace+1", 39),
                    New ItemData("Weapon: Staff of MOGREF", 40),
                    New ItemData("Magic: Scroll of KATINO", 41),
                    New ItemData("Armor: Leather+1", 42),
                    New ItemData("Armor: Chain Mail+1", 43),
                    New ItemData("Armor: Plate Mail+1", 44),
                    New ItemData("Shield: Shield+1", 45),
                    New ItemData("Magic: St. K.A.'s Foot (casts MALIKTO)", 46),
                    New ItemData("Magic: Scroll of BADIOS", 47),
                    New ItemData("Magic: Scroll of HALITO", 48),
                    New ItemData("Weapon: Staff+2", 49),
                    New ItemData("Weapon: Dragon Slayer", 50),
                    New ItemData("Helmet: Helm+1", 51),
                    New ItemData("Magic: Jeweled Amulet (casts DUMAPIC)", 52),
                    New ItemData("Magic: Scroll of BADIAL", 53),
                    New ItemData("Magic: Potion of SOPIC", 54),
                    New ItemData("Weapon: Long Sword+2", 55),
                    New ItemData("Shield: Good Hope Cape", 56),
                    New ItemData("Helmet: Magician's Hat (casts SOPIC)", 57),
                    New ItemData("Helmet: Novice's Cap (casts KATINO)", 58),
                    New ItemData("Magic: Scroll of DILTO", 59),
                    New ItemData("Gauntlets: Copper Gloves", 60),
                    New ItemData("Helmet: Initiate Turban (casts HALITO)", 61),
                    New ItemData("Helmet: Wizard Skullcap (casts MASOPIC)", 62),
                    New ItemData("Armor: Plate Mail+2", 63),
                    New ItemData("Shield: Shield+2", 64),
                    New ItemData("Special: Mordorcharge (Charge Oracle)", 65),
                    New ItemData("Magic: Potion of DIAL", 66),
                    New ItemData("Magic: Ring of PORFIC", 67),
                    New ItemData("Weapon: Were Slayer", 68),
                    New ItemData("Weapon: Mage Masher", 69),
                    New ItemData("Weapon: Mace of Curing", 70),
                    New ItemData("Weapon: Staff of MONTINO", 71),
                    New ItemData("Weapon: Blade Cusinart'", 72),
                    New ItemData("Magic: Amulet of BADIALMA", 73),
                    New ItemData("Magic: Rod of Flame (casts MAHALITO)", 74),
                    New ItemData("Shield: Cape of Hide", 75),
                    New ItemData("Shield: Cape of Jackel", 76),
                    New ItemData("Shield: Cape of Hide", 77),
                    New ItemData("Magic: Amulet of MAKANITO", 78),
                    New ItemData("Helmet: Diadem of MALOR", 79),
                    New ItemData("Magic: Scroll of BADIAL", 80),
                    New ItemData("Weapon: Dagger+2", 81),
                    New ItemData("Weapon: Dagger of Speed", 82),
                    New ItemData("Shield: Lich's Robes", 83),
                    New ItemData("Helmet: Skull's Cap", 84),
                    New ItemData("Magic: Potion of MASOPIC", 85),
                    New ItemData("Gauntlets: Silver Gloves", 86),
                    New ItemData("Special: GetOut of JailFree (Prison Release)", 87),
                    New ItemData("Misc: Gold Pyrite", 88),
                    New ItemData("Armor: Oxygen Mask (casts MAKANITO)", 89),
                    New ItemData("Misc: Chronicles of H", 90),
                    New ItemData("Armor: Lord's Garb", 91),
                    New ItemData("Weapon: Murasama Blade", 92),
                    New ItemData("Weapon: Shuriken", 93),
                    New ItemData("Armor: Chain Pro Ice", 94),
                    New ItemData("**ERR**", 95),
                    New ItemData("**ERR**", 96),
                    New ItemData("Magic: Ring of Healing", 97),
                    New ItemData("Magic: Ring of Dispelling", 98),
                    New ItemData("Magic: Ring of Death", 99),
                    New ItemData("Helmet: Adept Baldness", 100),
                    New ItemData("Magic: Arabic Diary (casts BADI)", 101),
                    New ItemData("Magic: Demonic Chimes (casts MAMORLIS)", 102),
                    New ItemData("Special: Black Candle", 103),
                    New ItemData("Special: Black Box", 104),
                    New ItemData("Special: St. Trebor Rump", 105),
                    New ItemData("Special: Bish's Tongue", 106),
                    New ItemData("Gauntlets: St. Rinbo Digit (casts TILTOWAIT)", 107),
                    New ItemData("Special: Arrow of Truth", 108),
                    New ItemData("Special: Orb of Dreams", 109),
                    New ItemData("Special: Rallying Horn", 110),
                    New ItemData("Special: Signet Ring", 111),
                    New ItemData("Gauntlets: Mythril Glove (Holds Amulet)", 112),
                    New ItemData("Special: Holy Limp Wrist", 113),
                    New ItemData("Shield: Twilight Cloak", 114),
                    New ItemData("Shield: Shadow Cloak", 115),
                    New ItemData("Helmet: Cone of Silence (casts MONTINO)", 116),
                    New ItemData("Shield: Darkness Cloak", 117),
                    New ItemData("Shield: Night Cloak", 118),
                    New ItemData("Shield: Entropy Cloak", 119)
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