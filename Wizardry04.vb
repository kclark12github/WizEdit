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
                    New ItemData("Broken Item", 0, "Misc", "", 0, "", ""),
                    New ItemData("Bloodstone", 1, "Special", "", 0, "Y", ""),
                    New ItemData("Lander's Turq.", 2, "Special", "", 0, "Y", ""),
                    New ItemData("Amber Dragon", 3, "Special", "", 0, "Y", ""),
                    New ItemData("Holy Hand Grenade of Aunty Ock", 4, "Weapon", "0-0 Damage", 0, "Y", "Pull the pin at LVL1 15E 15N"),
                    New ItemData("Winged Boots", 5, "Special", "", 0, "Y", ""),
                    New ItemData("Dreampainter Ka", 6, "Magic", "", 0, "Y", "Casts MADI"),
                    New ItemData("East Wind Sword", 7, "Weapon", "5-20 Damage", 0, "Y", ""),
                    New ItemData("West Wind Sword", 8, "Weapon", "1-1 Damage", 0, "Y", ""),
                    New ItemData("Dragon's Claw", 9, "Weapon", "6-20 Damage", 0, "Y", ""),
                    New ItemData("Hopalong Carrot", 10, "Weapon", "0-0 Damage", 0, "Y", ""),
                    New ItemData("Cleansing Oil", 11, "Special", "", 50000, "Y", ""),
                    New ItemData("Witching Rod", 12, "Magic", "", 0, "Y", "Casts KANDI"),
                    New ItemData("Aromatic Ball", 13, "Special", "", 0, "Y", ""),
                    New ItemData("Void Transducer", 14, "Special", "", 0, "Y", ""),
                    New ItemData("Kris of Truth", 15, "Weapon", "0-0 Damage", 0, "Y", ""),
                    New ItemData("Inn Key", 16, "Special", "", 0, "Y", ""),
                    New ItemData("Crystal Rose", 17, "Special", "", 0, "Y", ""),
                    New ItemData("Dab of Puce", 18, "Special", "", 0, "N", ""),
                    New ItemData("Pennonceaux", 19, "Special", "", 0, "N", ""),
                    New ItemData("Maintenance Cap", 20, "Helmet", "0", 0, "Y", "Casts DIAL"),
                    New ItemData("Long Sword", 21, "Weapon", "1-8 Damage", 25, "N", ""),
                    New ItemData("Short Sword", 22, "Weapon", "1-6 Damage", 15, "N", ""),
                    New ItemData("Anointed Mace", 23, "Weapon", "2-6 Damage", 30, "N", ""),
                    New ItemData("Anointed Flail", 24, "Weapon", "1-7 Damage", 150, "N", ""),
                    New ItemData("Staff", 25, "Weapon", "1-5 Damage", 10, "Y", ""),
                    New ItemData("Dagger", 26, "Weapon", "1-4 Damage", 5, "Y", ""),
                    New ItemData("Small Shield", 27, "Shield", "AC 2", 20, "N", ""),
                    New ItemData("Large Shield", 28, "Shield", "AC 3", 40, "N", ""),
                    New ItemData("Robes", 29, "Armor", "AC 1", 15, "Y", ""),
                    New ItemData("Leather Armor", 30, "Armor", "AC 2", 50, "N", ""),
                    New ItemData("Chain Mail", 31, "Armor", "AC 3", 90, "N", ""),
                    New ItemData("Breast Plate", 32, "Armor", "AC 4", 200, "N", ""),
                    New ItemData("Plate Mail", 33, "Armor", "AC 5", 750, "N", ""),
                    New ItemData("Helm", 34, "Helmet", "1", 100, "N", ""),
                    New ItemData("Potion of DIOS", 35, "Magic", "", 500, "Y", ""),
                    New ItemData("Potion of PORFIC", 36, "Magic", "", 300, "Y", ""),
                    New ItemData("Long Sword +1", 37, "Weapon", "2-9 Damage", 10000, "N", ""),
                    New ItemData("Short Sword +1", 38, "Weapon", "2-7 Damage", 15000, "N", ""),
                    New ItemData("Mace +1", 39, "Weapon", "3-9 Damage", 12500, "N", ""),
                    New ItemData("Staff of MOGREF", 40, "Weapon", "1-6 Damage", 3000, "Y", "Casts MOGREF"),
                    New ItemData("Scroll of KATINO", 41, "Magic", "", 500, "Y", ""),
                    New ItemData("Leather +1", 42, "Armor", "AC 3", 1500, "N", ""),
                    New ItemData("Chain Mail +1", 43, "Armor", "AC 4", 1500, "N", ""),
                    New ItemData("Plate Mail +1", 44, "Armor", "AC 6", 1500, "N", ""),
                    New ItemData("Shield +1", 45, "Shield", "AC 4", 1500, "N", ""),
                    New ItemData("St. K.A.'s Foot", 46, "Magic", "", 0, "Y", "Casts MALIKTO"),
                    New ItemData("Scroll of BADIOS", 47, "Magic", "", 500, "Y", ""),
                    New ItemData("Scroll of HALITO", 48, "Magic", "", 500, "Y", ""),
                    New ItemData("Staff +2", 49, "Weapon", "3-6 Damage", 2500, "Y", ""),
                    New ItemData("Dragon Slayer", 50, "Weapon", "2-11 Damage", 10000, "N", ""),
                    New ItemData("Helm +1", 51, "Helmet", "2", 3000, "N", ""),
                    New ItemData("Jeweled Amulet", 52, "Magic", "", 5000, "Y", "Casts DUMAPIC"),
                    New ItemData("Scroll of BADIAL", 53, "Magic", "", 500, "Y", ""),
                    New ItemData("Potion of SOPIC", 54, "Magic", "", 1500, "Y", ""),
                    New ItemData("Long Sword +2", 55, "Weapon", "3-12 Damage", 4000, "N", ""),
                    New ItemData("Good Hope Cape", 56, "Shield", "AC 3", 0, "Y", ""),
                    New ItemData("Magician's Hat", 57, "Helmet", "2", 0, "Y", "Casts SOPIC"),
                    New ItemData("Novice's Cap", 58, "Helmet", "2", 0, "Y", "Casts KATINO"),
                    New ItemData("Scroll of DILTO", 59, "Magic", "", 2500, "Y", ""),
                    New ItemData("Copper Gloves", 60, "Gauntlets", "AC 1", 6000, "N", ""),
                    New ItemData("Initiate Turban", 61, "Helmet", "1", 0, "Y", "Casts HALITO"),
                    New ItemData("Wizard Skullcap", 62, "Helmet", "3", 0, "Y", "Casts MASOPIC"),
                    New ItemData("Plate Mail +2", 63, "Armor", "AC 7", 6000, "N", ""),
                    New ItemData("Shield +2", 64, "Shield", "AC 5", 7000, "N", ""),
                    New ItemData("Mordorcharge", 65, "Special", "", 0, "Y", "Charge Oracle"),
                    New ItemData("Potion of DIAL", 66, "Magic", "", 5000, "Y", ""),
                    New ItemData("Ring of PORFIC", 67, "Magic", "", 10000, "Y", "Casts PORFIC"),
                    New ItemData("Were Slayer", 68, "Weapon", "2-11 Damage", 10000, "N", ""),
                    New ItemData("Mage Masher", 69, "Weapon", "2-7 Damage", 10000, "Y", ""),
                    New ItemData("Mace of Curing", 70, "Weapon", "1-8 Damage", 10000, "N", ""),
                    New ItemData("Staff of MONTINO", 71, "Weapon", "2-6 Damage", 15000, "Y", "Casts MONTINO"),
                    New ItemData("Blade Cusinart'", 72, "Weapon", "10-12 Damage", 15000, "Y", ""),
                    New ItemData("Amulet of BADIALMA", 73, "Magic", "", 15000, "N", "Casts BADIALMA"),
                    New ItemData("Rod of Flame", 74, "Magic", "", 25000, "Y", "Casts MAHALITO"),
                    New ItemData("Cape of Hide", 75, "Shield", "AC 2", 0, "Y", ""),
                    New ItemData("Cape of Jackal", 76, "Shield", "AC 4", 0, "Y", ""),
                    New ItemData("Cape of Hide", 77, "Shield", "AC ", 0, "", ""),
                    New ItemData("Amulet of MAKANITO", 78, "Magic", "", 20000, "Y", "Casts MAKANITO"),
                    New ItemData("Diadem of MALOR", 79, "Helmet", "2", 25000, "Y", "Casts MALOR"),
                    New ItemData("Scroll of BADIAL", 80, "Magic", "", 0, "", ""),
                    New ItemData("Dagger +2", 81, "Weapon", "3-6 Damage", 8000, "Y", ""),
                    New ItemData("Dagger of Speed", 82, "Weapon", "1-4 Damage", 30000, "Y", ""),
                    New ItemData("Lich's Robes", 83, "Shield", "AC 4", 8000, "Y", ""),
                    New ItemData("Skull's Cap", 84, "Helmet", "2", 50000, "Y", ""),
                    New ItemData("Potion of MASOPIC", 85, "Magic", "", 10000, "Y", ""),
                    New ItemData("Silver Gloves", 86, "Gauntlets", "AC 3", 60000, "N", ""),
                    New ItemData("GetOut of JailFree", 87, "Special", "", 0, "Y", "Prison release"),
                    New ItemData("Gold Pyrite", 88, "Misc", "", 0, "Y", ""),
                    New ItemData("Oxygen Mask", 89, "Armor", "AC 0", 0, "Y", "Casts MAKANITO"),
                    New ItemData("Chronicles of H", 90, "Misc", "", 0, "Y", ""),
                    New ItemData("Lord's Garb", 91, "Armor", "AC 10", 1000000, "Y", ""),
                    New ItemData("Muramasa Blade", 92, "Weapon", "10-50 Damage", 1000000, "N", ""),
                    New ItemData("Shuriken", 93, "Weapon", "11-15 Damage", 50000, "N", ""),
                    New ItemData("Chain Pro Ice", 94, "Armor", "AC 6", 150000, "Y", ""),
                    New ItemData("**ERR**", 95, "", "", 0, "", ""),
                    New ItemData("**ERR**", 96, "", "", 0, "", ""),
                    New ItemData("Ring of Healing", 97, "Magic", "", 300000, "Y", ""),
                    New ItemData("Ring of Dispelling", 98, "Magic", "", 500000, "Y", ""),
                    New ItemData("Ring of Death", 99, "Magic", "", 500000, "Y", ""),
                    New ItemData("Adept Baldness", 100, "Helmet", "4", 0, "Y", ""),
                    New ItemData("Arabic Diary", 101, "Magic", "", 0, "Y", "Casts BADI"),
                    New ItemData("Demonic Chimes", 102, "Magic", "", 0, "Y", "Casts MAMORLIS"),
                    New ItemData("Black Candle", 103, "Special", "", 0, "Y", "Casts LOMILWA"),
                    New ItemData("Black Box", 104, "Special", "", 0, "Y", "Stores 20 Items"),
                    New ItemData("St. Trebor Rump", 105, "Special", "", 100000, "Y", ""),
                    New ItemData("Bish's Tongue", 106, "Special", "", 0, "Y", "Casts LORTO"),
                    New ItemData("St. Rimbo Digit", 107, "Gauntlets", "AC 0", 0, "Y", "Casts TILTOWAIT"),
                    New ItemData("Arrow of Truth", 108, "Special", "", 0, "Y", ""),
                    New ItemData("Orb of Dreams", 109, "Special", "", 0, "Y", ""),
                    New ItemData("Rallying Horn", 110, "Special", "", 1000000, "N", ""),
                    New ItemData("Signet Ring", 111, "Special", "", 0, "N", ""),
                    New ItemData("Mythril Glove", 112, "Gauntlets", "AC 0", 0, "Y", "Holds AMULET"),
                    New ItemData("Holy Limp Wrist", 113, "Special", "", 0, "Y", "Casts DIALKO"),
                    New ItemData("Twilight Cloak", 114, "Shield", "AC 1", 0, "Y", ""),
                    New ItemData("Shadow Cloak", 115, "Shield", "AC 1", 0, "Y", ""),
                    New ItemData("Cone of Silence", 116, "Helmet", "1", 0, "Y", "Casts MONTINO"),
                    New ItemData("Darkness Cloak", 117, "Shield", "AC 1", 0, "Y", ""),
                    New ItemData("Night Cloak", 118, "Shield", "AC 2", 0, "Y", ""),
                    New ItemData("Entropy Cloak", 119, "Shield", "AC 4", 0, "Y", "")
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