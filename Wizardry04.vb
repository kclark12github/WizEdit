'Wizardry04.vb
'   Main Class for The Return of Werdna...
'   Copyright © 2017, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   10/09/19    Ken Clark       Migrated to VS2017;
'   09/02/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit On

Public Class Wizardry04
    Inherits WizEditBase
    Public Sub New(ByVal Caption As String, ByVal Icon As Icon, ByVal BoxArt As Image, ByVal Parent As Form)
        MyBase.New(Caption, Icon, BoxArt, Parent)
        ReDim mCharacters(Me.CharactersMax - 1)
        For iChar As Short = 0 To Me.CharactersMax - 1
            mCharacters(iChar) = New Character04(Me)
        Next iChar
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
    Protected mMasterMonsterGroupList As MonsterGroupData()
    Public ReadOnly Property MasterMonsterGroupList As MonsterGroupData()
        Get
            If mMasterMonsterGroupList Is Nothing Then
                mMasterMonsterGroupList = {
                    New MonsterGroupData("A Dink", 0, 10),
                    New MonsterGroupData("Fuzzballs", 1, 10),
                    New MonsterGroupData("Creeping Coins", 2, 10),
                    New MonsterGroupData("Bubbly Slimes", 3, 10),
                    New MonsterGroupData("Orcs", 4, 10),
                    New MonsterGroupData("Lvl 1 Mages", 5, 10),
                    New MonsterGroupData("Gas Clouds", 6, 10),
                    New MonsterGroupData("Skeletons", 7, 10),
                    New MonsterGroupData("Garian Raiders", 8, 10),
                    New MonsterGroupData("Lvl 1 Priests", 9, 10),
                    New MonsterGroupData("Zombies", 10, 10),
                    New MonsterGroupData("Kobolds", 11, 10),
                    New MonsterGroupData("Anacondas", 12, 9),
                    New MonsterGroupData("Ashers", 13, 9),
                    New MonsterGroupData("Crawling Kelps", 14, 9),
                    New MonsterGroupData("Creeping Cruds", 15, 9),
                    New MonsterGroupData("Dusters", 16, 9),
                    New MonsterGroupData("Huge Spiders", 17, 9),
                    New MonsterGroupData("Lvl 3 Priests", 18, 9),
                    New MonsterGroupData("Mummies", 19, 9),
                    New MonsterGroupData("No-See-Um Swarm", 20, 9),
                    New MonsterGroupData("Poltergeists", 21, 9),
                    New MonsterGroupData("Rogues", 22, 9),
                    New MonsterGroupData("Witches", 23, 9),
                    New MonsterGroupData("Banshees", 24, 8),
                    New MonsterGroupData("Bugbears", 25, 8),
                    New MonsterGroupData("Dragon Flies", 26, 8),
                    New MonsterGroupData("Gaze Hounds", 27, 8),
                    New MonsterGroupData("Harpies", 28, 8),
                    New MonsterGroupData("Looters", 29, 8),
                    New MonsterGroupData("Lvl 5 Priests", 30, 8),
                    New MonsterGroupData("Ronins", 31, 8),
                    New MonsterGroupData("Rotting Corpse", 32, 8),
                    New MonsterGroupData("Shades", 33, 8),
                    New MonsterGroupData("Spirits", 34, 8),
                    New MonsterGroupData("Wererats", 35, 8),
                    New MonsterGroupData("Blink Dogs", 36, 7),
                    New MonsterGroupData("Bushwackers", 37, 7),
                    New MonsterGroupData("Cockatrices", 38, 7),
                    New MonsterGroupData("Giant Slugs", 39, 7),
                    New MonsterGroupData("Giant Toads", 40, 7),
                    New MonsterGroupData("Goblin Shamans", 41, 7),
                    New MonsterGroupData("Goblins", 42, 7),
                    New MonsterGroupData("Moat Monsters", 43, 7),
                    New MonsterGroupData("Ogres", 44, 7),
                    New MonsterGroupData("Priestesses", 45, 7),
                    New MonsterGroupData("Strangler Vines", 46, 7),
                    New MonsterGroupData("Vorpal Bunnies", 47, 7),
                    New MonsterGroupData("Bishops", 48, 6),
                    New MonsterGroupData("Centaurs", 49, 6),
                    New MonsterGroupData("Grave Mists", 50, 6),
                    New MonsterGroupData("High Corsairs", 51, 6),
                    New MonsterGroupData("Hobgoblins", 52, 6),
                    New MonsterGroupData("Lifestealers", 53, 6),
                    New MonsterGroupData("Lvl 3 Samurai", 54, 6),
                    New MonsterGroupData("Master Ninjas", 55, 6),
                    New MonsterGroupData("Minor Daimyos", 56, 6),
                    New MonsterGroupData("Nightstalkers", 57, 6),
                    New MonsterGroupData("Werewolves", 58, 6),
                    New MonsterGroupData("Wights", 59, 6),
                    New MonsterGroupData("Boring Beetles", 60, 5),
                    New MonsterGroupData("Corr. Slimes", 61, 5),
                    New MonsterGroupData("D'Placer Beasts", 62, 5),
                    New MonsterGroupData("Gargoyles", 63, 5),
                    New MonsterGroupData("Gas Dragons", 64, 5),
                    New MonsterGroupData("Ghosts", 65, 5),
                    New MonsterGroupData("Hellhounds", 66, 5),
                    New MonsterGroupData("Komodo Dragons", 67, 5),
                    New MonsterGroupData("Masters/Dragons", 68, 5),
                    New MonsterGroupData("Priests of Fung", 69, 5),
                    New MonsterGroupData("Seraphim", 70, 5),
                    New MonsterGroupData("Weretigers", 71, 5),
                    New MonsterGroupData("Carriers", 72, 4),
                    New MonsterGroupData("Dark Riders", 73, 4),
                    New MonsterGroupData("Doppelgangers", 74, 4),
                    New MonsterGroupData("Evil Eyes", 75, 4),
                    New MonsterGroupData("Giant Mantises", 76, 4),
                    New MonsterGroupData("Goblin Princes", 77, 4),
                    New MonsterGroupData("Gorgons", 78, 4),
                    New MonsterGroupData("Lvl 6 Ninjas", 79, 4),
                    New MonsterGroupData("Masters/W. Wind", 80, 4),
                    New MonsterGroupData("Myrmidons", 81, 4),
                    New MonsterGroupData("Scrylls", 82, 4),
                    New MonsterGroupData("Wyverns", 83, 4),
                    New MonsterGroupData("Berserkers", 84, 3),
                    New MonsterGroupData("Bleebs", 85, 3),
                    New MonsterGroupData("Brass Dragons", 86, 3),
                    New MonsterGroupData("Champ Samurai", 87, 3),
                    New MonsterGroupData("Chimeras", 88, 3),
                    New MonsterGroupData("Fiends", 89, 3),
                    New MonsterGroupData("Major Daimyos", 90, 3),
                    New MonsterGroupData("Rocs", 91, 3),
                    New MonsterGroupData("Trolls", 92, 3),
                    New MonsterGroupData("Vampires", 93, 3),
                    New MonsterGroupData("Will O' Wisps", 94, 3),
                    New MonsterGroupData("Xenos", 95, 3),
                    New MonsterGroupData("Cyclopes", 96, 2),
                    New MonsterGroupData("Dragon Zombies", 97, 2),
                    New MonsterGroupData("Fire Giants", 98, 2),
                    New MonsterGroupData("Firedrakes", 99, 2),
                    New MonsterGroupData("Frost Giants", 100, 2),
                    New MonsterGroupData("Hatamotos", 101, 2),
                    New MonsterGroupData("Hydrae", 102, 2),
                    New MonsterGroupData("Liches", 103, 2),
                    New MonsterGroupData("Manticores", 104, 2),
                    New MonsterGroupData("Masters/Summer", 105, 2),
                    New MonsterGroupData("Murphy's Ghosts", 106, 2),
                    New MonsterGroupData("Succubi", 107, 2),
                    New MonsterGroupData("A Demon Lord", 108, 1),
                    New MonsterGroupData("Black Dragons", 109, 1),
                    New MonsterGroupData("Fleck", 110, 1),
                    New MonsterGroupData("Foaming Molds", 111, 1),
                    New MonsterGroupData("Gold Dragons", 112, 1),
                    New MonsterGroupData("Greater Demons", 113, 1),
                    New MonsterGroupData("High Masters", 114, 1),
                    New MonsterGroupData("Iron Golems", 115, 1),
                    New MonsterGroupData("Lycurgi", 116, 1),
                    New MonsterGroupData("Maelifics", 117, 1),
                    New MonsterGroupData("Poison Giants", 118, 1),
                    New MonsterGroupData("Vampire Lords", 119, 1)
                }
            End If
            Return mMasterMonsterGroupList
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

    Public Overrides Function GetCharacter(ByVal Tag As String) As CharacterBase
        Return mCharacters(CInt(Tag.Substring(Tag.Length - 1, 1)) - 1)
    End Function
    Public Overrides Sub Read()
        Dim binReader As BinaryReader = Nothing
        Dim iGame As Short = 0
        Try
            'Instead of individual characters Wizardry (4) supports up to 8 games...
            'This is comparable when you consider the nature of the game - Werdna is the only "Character".
            binReader = New BinaryReader(File.Open(Me.ScenarioDataPath, FileMode.Open))
            binReader.BaseStream.Position = Me.CharacterDataOffset
            For iGame = 1 To Me.CharactersMax
                Characters(iGame - 1).Read(binReader)
            Next iGame
        Catch ex As Exception
            Debug.WriteLine(String.Format("{0} encountered reading Game #{1}{2}{3}", New Object() {ex.GetType.Name, iGame, vbCrLf, ex.ToString}))
            Throw
        Finally
            If binReader IsNot Nothing Then binReader.Close() : binReader = Nothing
        End Try
    End Sub
    Public Overrides Sub Save()
        Dim binWriter As BinaryWriter = Nothing
        Dim iGame As Short = 0
        Try
            Backup()
            'Instead of individual characters Wizardry (4) supports up to 8 games...
            'This is comparable when you consider the nature of the game - Werdna is the only "Character".
            binWriter = New BinaryWriter(File.Open(Me.ScenarioDataPath, FileMode.Open, FileAccess.Write, FileShare.None))
            binWriter.BaseStream.Position = Me.CharacterDataOffset
            For iGame = 1 To Me.CharactersMax
                Characters(iGame - 1).Save(binWriter)
            Next iGame
        Catch ex As Exception
            Debug.WriteLine(String.Format("{0} encountered saving Game #{1}{2}{3}", New Object() {ex.GetType.Name, iGame, vbCrLf, ex.ToString}))
            Throw
        Finally
            If binWriter IsNot Nothing Then binWriter.Close() : binWriter = Nothing
        End Try
    End Sub
    Public Overrides Sub Show()
        mForm = New frmWizardry04(Me, mCaption, mIcon, mBoxArt)
        mForm.ShowDialog(mParent)
    End Sub
End Class