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
'MAELSTROM V
'0x4C00	Name
'0x4C10 Password
'0x4C20	? Out[I2]
'0x4C22 ? Race[I2]
'0x4C24 ? Profession[I2]
'0x4C26 ? Age[I2]
'0x4C28 ? Status[I2]
'0x4C2A ? Alignment[I2]
'0x4C2C ? Statistics[I4];
'0x4C20 ? Unknown[I4];
'0x4C24 ? Gold[I6];

'0x4C3A ItemCount[I2];
'	0x4C3C ItemCode1[I2];
'	0x4C3C ItemStatus1[I2];  2^0 = Equipped; 2^1 = Identified;
'       0x00=0000 0000 = Unidentified
'       0x01=0000 0001 = Equipped & Unidentified (Weapon)
'       0x02=0000 0010 = Identified
'       0x03=0000 0011 = Equipped & Identified (Weapon)
'       0x04=0000 0100 = Unidentified
'       0x05=0000 0101 = Equipped & Unidentified
'       0x06=0000 0110 = Identified
'       0x07=0000 0111 = Equipped & Identified (Weapon)
'       0x08=0000 1000
'       0x09=0000 1001
'       0x0A=0000 1010
'       0x0B=0000 1011 = Equipped & Identified (Armor)
'	0x4C3E Cursed1[I2];
'	0x4C40 Identified1[I2];

'	0x4C44 Equipped2[I2];
'	0x4C46 Cursed2[I2];
'	0x4C48 Identified2[I2];
'	0x4C4A ItemCode2[I2];	

'	0x4C4C Equipped3[I2];
'	0x4C4E Cursed3[I2];
'	0x4C50 Identified3[I2];
'	0x4C52 ItemCode3[I2];	

'	0x4C54 Equipped4[I2];
'	0x4C56 Cursed4[I2];
'	0x4C58 Identified4[I2];
'	0x4C5A ItemCode4[I2];	

'	0x4C5C Equipped5[I2];
'	0x4C5E Cursed5[I2];
'	0x4C60 Identified5[I2];
'	0x4C62 ItemCode5[I2];

'	0x4C64 Equipped6[I2];
'	0x4C66 Cursed6[I2];
'	0x4C68 Identified6[I2];
'	0x4C6A ItemCode6[I2];	

'	0x4C6C Equipped7[I2];
'	0x4C6E Cursed7[I2];
'	0x4C70 Identified7[I2];
'	0x4C72 ItemCode7[I2];	

'	0x4C74 Equipped8[I2];
'	0x4C76 Cursed8[I2];
'	0x4C78 Identified8[I2];
'	0x4C7A ItemCode8[I2];	
'TODO: vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'0x4C7C Keys[I6]; [I12]?
'0x4C82 Level.Current[I2];.Max[I2];
'0x4C86 HP.Current[I2];.Max[I2];
'0x4C8A SpellBooks[8B];
'0x4C92 SpellPoints[I2]x14
'0x4CAE Unknown[[I2]x14 (Coincidence?) - Spells Cast maybe?
'0x4CCA Location[I2];
'0x4CCE	Down[I2] - Screen says -2 but data says 0E 00
'0x4CD0 Group1.Count[I2]
'0x4CD2 Group2.Count[I2]
'0x4CD4 Group3.Count[I2]
'0x4CD6 Group1.Code[I2]
'04ACD8 Group2.Code[I2]
'0x4CDA Group3.Code[I2]
'0x4CDC Group1.Name ("A DINK") String(15) 0x00
'0x4CEC Group2.Name ("ENTELECHY FUFF") String(15) 0x77
'0x4CFC Group3.Name ("VAMPIRE LORD") String(15) 0x70
'0x4CFC 

Option Explicit On

Public Class Wizardry05
    Inherits WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(Caption, Icon, BoxArt, Parent)
    End Sub
    Public Overrides ReadOnly Property CharacterDataOffset As Int32
        Get
            Return &H4C00
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
                    New ItemData("Broken Item", 0, "Misc", "", 0, "*", "Worthless"),
                    New ItemData("Torch", 1, "Misc", "", 10, "*", "Casts MILWA"),
                    New ItemData("Lantern", 2, "Magic", "", 75, "*", "Casts LOMILWA"),
                    New ItemData("Rubber Duck", 3, "Misc", "", 0, "*", "Equip for Perfect Swimming"),
                    New ItemData("Dagger", 4, "Weapon", "2-4 Damage", 25, "FMTSLN", "Close Range;"),
                    New ItemData("Staff", 5, "Weapon", "1-5 Damage", 30, "*", "Close Range;"),
                    New ItemData("Short Sword", 6, "Weapon", "1-6 Damage", 35, "FTSLN", "Close Range;"),
                    New ItemData("Long Sword", 7, "Weapon", "2-7 Damage", 45, "FSLN", "Close Range;"),
                    New ItemData("Mace", 8, "Weapon", "3-6 Damage", 100, "FPBSLN", "Close Range;"),
                    New ItemData("Battle Axe", 9, "Weapon", "4-8 Damage", 180, "FSLN", "Close Range;"),
                    New ItemData("Pike", 10, "Weapon", "2-6 Damage", 250, "FSLN", "Short Range;"),
                    New ItemData("War Hammer", 11, "Weapon", "4-9 Damage", 400, "FL", "Short Range;"),
                    New ItemData("Holy Basher", 12, "Weapon", "3-7 Damage", 400, "PB", "Short Range;"),
                    New ItemData("Long Bow", 13, "Weapon", "2-7 Damage", 325, "FSLN", "Long Range;"),
                    New ItemData("Thieve's Bow", 14, "Weapon", "2-6 Damage", 600, "FTSLN", "Medium Range;"),
                    New ItemData("Robes", 15, "Armor", "AC 1", 20, "*", ""),
                    New ItemData("Leather Armor", 16, "Armor", "AC 2", 95, "FPTBSLN", ""),
                    New ItemData("Chain Mail", 17, "Armor", "AC 3", 145, "FPBSLN", ""),
                    New ItemData("Scale Mail", 18, "Armor", "AC 4", 400, "FSLN", ""),
                    New ItemData("Plate Mail", 19, "Armor", "AC 5", 750, "FL", ""),
                    New ItemData("Target Shield", 20, "Shield", "AC 1", 65, "FTSLN", ""),
                    New ItemData("Heater Shield", 21, "Shield", "AC 2", 125, "FL", ""),
                    New ItemData("Leather Sallet", 22, "Helmet", "1", 250, "FPBSLN", ""),
                    New ItemData("Leather Gloves", 23, "Gauntlets", "AC 1", 500, "FPTBSLN", ""),
                    New ItemData("Short Sword +1", 24, "Weapon", "7-12 Damage", 1500, "FSLN", "Close Range;"),
                    New ItemData("Long Sword +1", 25, "Weapon", "7-12 Damage", 1500, "FSLN", "Close Range;"),
                    New ItemData("Blackblade", 26, "Weapon", "6-12 Damage", 1500, "FSLN", "Close Range;"),
                    New ItemData("Katana", 27, "Weapon", "7-13 Damage", 1750, "SN", "Close Range;"),
                    New ItemData("Battle Axe +1", 28, "Weapon", "8-14 Damage", 1750, "FSLN", "Close Range;"),
                    New ItemData("Morningstar", 29, "Weapon", "4-10 Damage", 2000, "FPBSLN", "Short Range;"),
                    New ItemData("Runed Flail", 30, "Weapon", " Damage", 2000, "FPBSLN", "Cursed; AC -2"),
                    New ItemData("Halberd", 31, "Weapon", "7-13 Damage", 2500, "FSLN", "Short Range;"),
                    New ItemData("Lt. Crossbow", 32, "Weapon", "5-10 Damage", 2500, "FTSLN", "Long Range;"),
                    New ItemData("Leather +1", 33, "Armor", "AC 3", 1500, "FPTBSLN", ""),
                    New ItemData("Chain Mail +1", 34, "Armor", "AC 4", 1750, "FPBSLN", ""),
                    New ItemData("Scale Mail +1", 35, "Armor", "AC 5", 2000, "FSLN", ""),
                    New ItemData("Plate Mail +1", 36, "Armor", "AC 6", 2500, "FL", ""),
                    New ItemData("Silver Mail", 37, "Armor", "AC 4", 2500, "FPSLN", "Cursed; Invoke: Heal"),
                    New ItemData("Target +1", 38, "Shield", "AC 2", 1500, "FTSLN", ""),
                    New ItemData("Heater +1", 39, "Shield", "AC 3", 2000, "FL", ""),
                    New ItemData("Crested Shield", 40, "Shield", "AC -3", 2000, "FL", "Cursed;"),
                    New ItemData("Brass Sallet", 41, "Helmet", "2", 1500, "FPBSLN", ""),
                    New ItemData("Iron Gloves", 42, "Gauntlets", "AC 2", 2500, "FSLN", ""),
                    New ItemData("Bracers", 43, "Gauntlets", "AC ", 2500, "*", "[FSLN] AC: 1"),
                    New ItemData("Long Sword +2", 44, "Weapon", "7-17 Damage", 5000, "FSLN", "Close Range;"),
                    New ItemData("Robinsword", 45, "Weapon", "5-15 Damage", 7000, "FTSLN", "Close Range;"),
                    New ItemData("Sword of Fire", 46, "Weapon", "8-22 Damage", 10000, "FSLN", "Close Range; Casts MAHALITO"),
                    New ItemData("Master Katana", 47, "Weapon", "7-19 Damage", 13500, "S", "Close Range;"),
                    New ItemData("Soulstealer", 48, "Weapon", " Damage", 13500, "FSLN", "Cursed; Invoke: Vitality -1, Age +1"),
                    New ItemData("Battle Axe +2", 49, "Weapon", "10-20 Damage", 8500, "FSLN", "Close Range;"),
                    New ItemData("Axe of Death", 50, "Weapon", " Damage", 8500, "FSLN", "Cursed;"),
                    New ItemData("Sacred Basher", 51, "Weapon", "8-14 Damage", 7500, "PB", "Short Range;"),
                    New ItemData("Faust Halberd", 52, "Weapon", "8-20 Damage", 10000, "FSLN", "Short Range; Invoke: Vitality -1"),
                    New ItemData("Silver Hammer", 53, "Weapon", "10-20 Damage", 10000, "FSLN", "Short Range; Invoke: Strength +1, Luck -1"),
                    New ItemData("Mage's Yew Bow", 54, "Weapon", "6-12 Damage", 12000, "M", "Long Range; Invoke: Vitality +1"),
                    New ItemData("Hv. Crossbow", 55, "Weapon", "8-15 Damage", 12000, "FTSLN", "Long Range;"),
                    New ItemData("Leather +2", 56, "Armor", "AC 4", 4000, "FPTBSLN", ""),
                    New ItemData("Chain Mail +2", 57, "Armor", "AC 5", 6000, "FPBSLN", ""),
                    New ItemData("Scale Mail +2", 58, "Armor", "AC 6", 8000, "FSLN", ""),
                    New ItemData("Plate Mail +2", 59, "Armor", "AC 7", 10000, "FL", ""),
                    New ItemData("Scarlet Robes", 60, "Armor", "AC -2", 4500, "M", "Cursed;"),
                    New ItemData("Emerald Robes", 61, "Armor", "AC 4", 4500, "*", ""),
                    New ItemData("Heater +2", 62, "Shield", "AC 4", 5000, "FL", ""),
                    New ItemData("Bacinet", 63, "Helmet", "3", 3500, "FSLN", ""),
                    New ItemData("Cone of Fire", 64, "Helmet", "-4", 3000, "*", "Cursed; Invoke: Ashes"),
                    New ItemData("Silver Gloves", 65, "Gauntlets", "AC 3", 7500, "FSLN", ""),
                    New ItemData("Bracers +1", 66, "Gauntlets", "AC ", 10000, "*", "[FSLN] AC:2"),
                    New ItemData("Long Sword +3", 67, "Weapon", "12-22 Damage", 20000, "FSLN", "Close Range; Formerly Blade Cusinart'"),
                    New ItemData("Plate Mail +3", 68, "Armor", "AC 8", 25000, "FSLN", ""),
                    New ItemData("Shield Pro Magic", 69, "Shield", "AC 3", 20000, "FSLN", ""),
                    New ItemData("Jeweled Armet", 70, "Gauntlets", "AC 4", 12500, "FL", ""),
                    New ItemData("Wizard's Cap", 71, "Helmet", "1", 8000, "M", ""),
                    New ItemData("Gloves of Myrdall", 72, "Gauntlets", "AC 4", 40000, "FSLN", ""),
                    New ItemData("Cloak of Capricorn", 73, "Armor", "AC ", 9000, "*", "AC: 2"),
                    New ItemData("Sylvan Bow", 74, "Weapon", "14-26 Damage", 100000, "FTSLN", "Long Range; Invoke: Agility +1"),
                    New ItemData("Muramasa Katana", 75, "Weapon", "15-30 Damage", 150000, "SN", "Close Range; Invoke: Vitality +1"),
                    New ItemData("Odinsword", 76, "Weapon", "15-35 Damage", 250000, "L", "Close Range; Invoke: Vitality +1"),
                    New ItemData("Gold Plate +5", 77, "Armor", "AC 10", 250000, "FSLN", ""),
                    New ItemData("Ring of Frozz", 78, "Magic", "", 5000, "*", "[M] AC: 2; Invoke: 9s All Spell Pts."),
                    New ItemData("Ring of Skulls", 79, "Magic", "", 5000, "*", "Invoke: Piety -1, Age -1"),
                    New ItemData("Ring of MADI", 80, "Magic", "", 15000, "*", "Invoke: Casts MADI"),
                    New ItemData("Ring of Jade", 81, "Magic", "AC -2", 10000, "*", "Cursed; Invoke: Age +1"),
                    New ItemData("Ring of Solitude", 82, "Magic", "", 20000, "*", "Invoke: Luck +1"),
                    New ItemData("Ankh of Wonder", 83, "Magic", "", 12000, "*", "AC: 1; Invoke: Casts IHALON"),
                    New ItemData("Ankh of Power", 84, "Magic", "", 12000, "*", "AC: 1; Invoke: Strength +1"),
                    New ItemData("Ankh of Life", 85, "Magic", "", 12000, "*", "AC: 1; Casts MADI"),
                    New ItemData("Ankh of Intellect", 86, "Magic", "", 12000, "*", "AC: 1; Invoke: I.Q. +1"),
                    New ItemData("Ankh of Sanctity", 87, "Magic", "", 12000, "*", "AC: 1; Invoke: Piety +1"),
                    New ItemData("Ankh of Youth", 88, "Magic", "", 12000, "*", "AC: 1; Invoke: Age -1"),
                    New ItemData("Staff of Summoning", 89, "Weapon", "4-9 Damage", 7750, "MPB", "Close Range; Casts BAMORDI"),
                    New ItemData("Staff of Death", 90, "Weapon", " Damage", 7750, "MPB", "Cursed; Invoke: Casts BADI"),
                    New ItemData("Scroll of KATINO", 91, "Magic", "", 250, "*", ""),
                    New ItemData("Scroll of Stoning", 92, "Magic", "", 750, "*", "Casts BOLATU"),
                    New ItemData("Scroll of Fire", 93, "Magic", "", 1250, "*", "Casts MAHALITO"),
                    New ItemData("Scroll of Conjuring", 94, "Magic", "", 3500, "*", "Casts SOCORDI"),
                    New ItemData("Potion of DIOS", 95, "Magic", "", 100, "*", ""),
                    New ItemData("Potion of Charming", 96, "Magic", "", 350, "*", "Casts KATU"),
                    New ItemData("Potion of LATUMOFIS", 97, "Magic", "", 250, "*", ""),
                    New ItemData("Potion of DIALKO", 98, "Magic", "", 400, "*", "Casts DIALKO"),
                    New ItemData("Potion of Wounding", 99, "Magic", "", 500, "*", "Casts BADIAL"),
                    New ItemData("Potion of MADI", 100, "Magic", "", 2500, "*", "Casts MADI"),
                    New ItemData("King of Diamonds", 101, "Special", "", 0, "*", "Used for Endgame Level 8 Access"),
                    New ItemData("Queen of Hearts", 102, "Special", "", 0, "*", "Used for Endgame Level 8 Access"),
                    New ItemData("Jack of Spades", 103, "Special", "", 0, "*", "Used for Endgame Level 8 Access"),
                    New ItemData("Ace of Clubs", 104, "Special", "", 0, "*", "Used for Endgame Level 8 Access"),
                    New ItemData("Munke Wand", 105, "Weapon", "2-8 Damage", 0, "MB", "Close Range; Use at LVL7 1W 13S"),
                    New ItemData("Lightning Rod", 106, "Weapon", "8-20 Damage", 0, "MB", "Close Range; Use at LVL7 13E 4S"),
                    New ItemData("Lark in a Cage", 107, "Special", "", 0, "*", "Use at LVL7 13W 0N"),
                    New ItemData("Staff of Water", 108, "Weapon", "6-12 Damage", 0, "*", "Short Range; Use during Endgame"),
                    New ItemData("Staff of Fire", 109, "Weapon", "6-12 Damage", 0, "*", "Short Range; Use during Endgame"),
                    New ItemData("Staff of Air", 110, "Weapon", "6-12 Damage", 0, "*", "Short Range; Use during Endgame"),
                    New ItemData("Staff of Earth", 111, "Weapon", "6-12 Damage", 0, "*", "Short Range; Use during Endgame"),
                    New ItemData("Potion of Demon-Out", 112, "Magic", "", 4500, "*", "Casts MOGATO"),
                    New ItemData("Gold Medallion", 113, "Special", "", 50000, "*", "AC: 2, Use at LVL6 14E 5S"),
                    New ItemData("Ice Key", 114, "Special", "", 0, "*", "Use at LVL6 5E 25S"),
                    New ItemData("Ticket Stubs", 115, "Special", "", 0, "*", "Use at LVL5 7E 1S"),
                    New ItemData("Tickets", 116, "Special", "", 0, "*", "Use at LVL5 7E 1S"),
                    New ItemData("Skeleton Key", 117, "Special", "", 0, "*", "Use at LVL4 4E 24S"),
                    New ItemData("Pocketwatch", 118, "Special", "", 0, "*", "Use at LVL4 2E 21S"),
                    New ItemData("Battery", 119, "Special", "", 0, "*", "Use at LVL3 6E 10S"),
                    New ItemData("Petrified Demon", 120, "Special", "AC 2", 0, "*", "Cursed; Invoke: Vitality -1, +HP; Use at LVL4 6E 17S"),
                    New ItemData("Gold Key", 121, "Special", "", 0, "*", "Use at LVL4 15W 21S"),
                    New ItemData("Blue Candle", 122, "Special", "", 3000, "*", "Use at LVL3 12E 25S"),
                    New ItemData("Jeweled Scepter", 123, "Special", "", 0, "*", "Use at LVL2 12E 5N"),
                    New ItemData("Potion of Spirit-Away", 124, "Magic", "", 500, "*", "Casts MORLIS; Use at LVL2 4E 0N"),
                    New ItemData("Hacksaw", 125, "Special", "", 0, "*", "Use at LVL2 2E 15S"),
                    New ItemData("Bottle of Rum", 126, "Special", "", 0, "*", "Use at LVL2 7W 3N"),
                    New ItemData("Silver Key", 127, "Special", "", 0, "*", "Use at LVL1 5E 27N"),
                    New ItemData("Bag of Tokens", 128, "Special", "", 0, "*", "Use at LVL1 12E 4N"),
                    New ItemData("Brass Key", 129, "Special", "", 0, "*", "Use at LVL1 6E 3N"),
                    New ItemData("Orb of Llylgamyn", 130, "Special", "", 0, "*", "Use at LVL1 8E 17N and Endgame"),
                    New ItemData("Heart of Abriel", 131, "Special", "", 0, "*", "Wins the Game"),
                    New ItemData("Holy Talisman", 132, "Magic", "", 25000, "*", "Casts DUMAPIC; Invoke: Piety -1"),
                    New ItemData("Amulet of Rainbows", 133, "Magic", "", 10000, "*", "Casts VASKYRE"),
                    New ItemData("Amulet of Screens", 134, "Magic", "", 10000, "*", "Casts CORTU"),
                    New ItemData("Amulet of Flames", 135, "Magic", "", 10000, "*", "Casts LAHALITO")
                }
            End If
            Return mMasterItemList
        End Get
    End Property
    Public Overrides ReadOnly Property RegDataPath As String
        Get
            Return "Scenario05DataPath"
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