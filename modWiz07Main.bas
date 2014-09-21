Attribute VB_Name = "modWiz07Main"
'modWiz07Main - modWiz07Main.bas
'   Main module for Crusaders of the Dark Savant / Wizardry Gold...
'   Copyright © 2000, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Programmer:     Description:
'   08/26/00    Ken Clark       Created;
'=================================================================================================================================
Option Explicit

'=================================================================================================================================
'Note: This module is mostly taken from a C application named, oddly enough, WizEdit.
'      The original WizEdit was written for DOS at the end of 1995. Some of that C
'      code is imortalized here for reference.
'=================================================================================================================================
Global Const Wiz07ItemMapMax As Integer = 569
Global Const Wiz07RaceMapMax As Integer = 10
Global Const Wiz07ProfessionMapMax As Integer = 13
Global Const Wiz07ConditionMapMax As Integer = 11
Global Const Wiz07SpellMapMax As Integer = 95

'/* WIZARDRY.H
'*/
'
'struct Item
'{
'   unsigned short int     ItemCode;
'   unsigned short int     Weight;
'   unsigned short int     PictureCode;
'   unsigned char          Count;
'   unsigned char          Status;
'   unsigned short int     Unknown;
'   unsigned char          Filler;
'   unsigned char          AC;
'};
Type Wiz07Item
    ItemCode As Integer
    Weight As Integer
    PictureCode As Integer
    Count As Byte
    Status As Byte
    Unknown As Integer
    Filler As Byte
    AC As Byte
End Type

'struct POINTS
'{
'   unsigned short int   Current;
'   unsigned short int   Maximum;
'};
Type Wiz07Points
    Current As Integer
    Maximum As Integer
End Type

'struct Character
'{
'   char                 Name[8];
'   unsigned char        Unknown1[4];
'   unsigned long int    EXP;
'   unsigned long int    MKS;
'   unsigned long int    GP;
'   struct Points        HP;
'   struct Points        STA;
'   struct Points        CC;
'   unsigned short int   Level;
'   unsigned short int   Lives;
'   struct
'   {
'      struct Points     Fire;
'      struct Points     Water;
'      struct Points     Air;
'      struct Points     Earth;
'      struct Points     Mental;
'      struct Points     Divine;
'   } SpellPoints;
'   struct Item          ItemList[10];
'   struct Item          SwagBag[10];
'   unsigned char        Unknown2[64];
'   struct
'   {
'      unsigned char     STR;
'      unsigned char     INT;
'      unsigned char     PIE;
'      unsigned char     VIT;
'      unsigned char     DEX;
'      unsigned char     SPD;
'      unsigned char     PER;
'      unsigned char     KAR;
'   } Attributes;
'   struct
'   {
'      struct
'      {
'         unsigned char  Wand;
'         unsigned char  Sword;
'         unsigned char  Axe;
'         unsigned char  Mace;
'         unsigned char  Pole;
'         unsigned char  Throwing;
'         unsigned char  Sling;
'         unsigned char  Bow;
'         unsigned char  Shield;
'         unsigned char  HandToHand;
'      } Weaponry;
'      struct
'      {
'         unsigned char  Swimming;
'         unsigned char  Climbing;
'         unsigned char  Scouting;
'         unsigned char  Music;
'         unsigned char  Oratory;
'         unsigned char  Legerdemain;
'         unsigned char  Skulduggery;
'         unsigned char  Ninjutsu;
'      } Physical;
'      struct
'      {
'         unsigned char  Firearms;
'         unsigned char  Reflextion;
'         unsigned char  SnakeSpeed;
'         unsigned char  EagleEye;
'         unsigned char  PowerStrike;
'         unsigned char  MindControl;
'      } Personal;
'      struct
'      {
'         unsigned char  Artifacts;
'         unsigned char  Mythology;
'         unsigned char  Mapping;
'         unsigned char  Scribe;
'         unsigned char  Diplomacy;
'         unsigned char  Alchemy;
'         unsigned char  Theology;
'         unsigned char  Theosophy;
'         unsigned char  Thaumaturgy;
'         unsigned char  Kirijutsu;
'      } Academia;
'   } Skills;
'   unsigned char        Unknown3[36];
'   struct
'   {
'      unsigned char Aptitude[96];
'      unsigned EnergyBlast       : 1;
'      unsigned BlindingFlash     : 1;
'      unsigned PsionicFire       : 1;
'      unsigned Fireball          : 1;
'      unsigned FireShield        : 1;
'      unsigned DazzlingLights    : 1;
'      unsigned FireBomb          : 1;
'      unsigned Lightning         : 1;
'      unsigned PrismicMissile    : 1;
'      unsigned Firestorm         : 1;
'      unsigned NuclearBlast      : 1;
'
'      unsigned ChillingTouch     : 1;
'      unsigned Stamina           : 1;
'      unsigned Terror            : 1;
'      unsigned Weaken            : 1;
'      unsigned Slow              : 1;
'      unsigned Haste             : 1;
'      unsigned CureParalysis     : 1;
'      unsigned IceShield         : 1;
'      unsigned Restfull          : 1;
'      unsigned IceBall           : 1;
'      unsigned Paralyze          : 1;
'      unsigned Superman          : 1;
'      unsigned Deepfreeze        : 1;
'      unsigned DrainingCloud     : 1;
'      unsigned CureDisease       : 1;
'
'      unsigned Poison            : 1;
'      unsigned MissileShield     : 1;
'      unsigned ShrillSound       : 1;
'      unsigned StinkBomb         : 1;
'      unsigned AirPocket         : 1;
'      unsigned Silence           : 1;
'      unsigned PoisonGas         : 1;
'      unsigned CurePoison        : 1;
'      unsigned Whirlwind         : 1;
'      unsigned PurifyAir         : 1;
'      unsigned DeadlyPoison      : 1;
'      unsigned Levitate          : 1;
'      unsigned ToxicVapors       : 1;
'      unsigned NoxiousFumes      : 1;
'      unsigned Asphyxiation      : 1;
'      unsigned DeadlyAir         : 1;
'      unsigned DeathCloud        : 1;
'
'      unsigned AcidSplash        : 1;
'      unsigned ItchingSkin       : 1;
'      unsigned ArmorShield       : 1;
'      unsigned Direction         : 1;
'      unsigned KnockKnock        : 1;
'      unsigned Blades            : 1;
'      unsigned Armorplate        : 1;
'      unsigned Web               : 1;
'      unsigned WhippingRocks     : 1;
'      unsigned AcidBomb          : 1;
'      unsigned Armormelt         : 1;
'      unsigned Crush             : 1;
'      unsigned CreateLife        : 1;
'      unsigned CureStone         : 1;
'
'      unsigned MentalAttack      : 1;
'      unsigned Sleep             : 1;
'      unsigned Bless             : 1;
'      unsigned Charm             : 1;
'      unsigned CureLesserCnd     : 1;
'      unsigned DivineTrap        : 1;
'      unsigned DetectSecret      : 1;
'      unsigned Identify          : 1;
'      unsigned Confusion         : 1;
'      unsigned Watchbells        : 1;
'      unsigned HoldMonsters      : 1;
'      unsigned Mindread          : 1;
'      unsigned SaneMind          : 1;
'      unsigned PsionicBlast      : 1;
'      unsigned Illusion          : 1;
'      unsigned WizardsEye        : 1;
'      unsigned Spooks            : 1;
'      unsigned Death             : 1;
'      unsigned LocateObject      : 1;
'      unsigned MindFlay          : 1;
'      unsigned FindPerson        : 1;
'
'      unsigned HealWounds        : 1;
'      unsigned MakeWounds        : 1;
'      unsigned MagicMissile      : 1;
'      unsigned DispellUndead     : 1;
'      unsigned EnchantedBlade    : 1;
'      unsigned Blink             : 1;
'      unsigned MagicScreen       : 1;
'      unsigned Conjuration       : 1;
'      unsigned AntiMagic         : 1;
'      unsigned RemoveCurse       : 1;
'      unsigned Healfull          : 1;
'      unsigned Lifesteal         : 1;
'      unsigned AstralGate        : 1;
'      unsigned ZapUndead         : 1;
'      unsigned Recharge          : 1;
'      unsigned WordOfDeath       : 1;
'      unsigned Resurrection      : 1;
'      unsigned DeathWish         : 1;
'   } Spells;
'   unsigned char        Unknown4[12];
'   unsigned char        PictureCode;
'   unsigned char        Race;
'   unsigned char        Gender;
'   unsigned char        Profession;
'   unsigned char        Alive;         /* Under Investigation */
'   unsigned char        ConditionCode; /* Under Investigation */
'   unsigned char        Unknown5[12];
'};
Type Wiz07Character
    Name As String * 8                  'Null Terminated
    Unknown1(1 To 4) As Byte            '???
    
    'Secondary Statistics...
    EXP As Long                         'Experience Points
    MKS As Long                         'Monster Kills
    GP As Long                          'Gold Pieces
    HP As Wiz07Points                   'Hit Points
    STM As Wiz07Points                  'Stamina
    CC As Wiz07Points                   'Carrying Capacity
    Level As Integer                    'Level (Duh)
    Lives As Integer                    'Number of Lives Used
    
    'Spell Points...
    FireSpellPoints  As Wiz07Points
    WaterSpellPoints  As Wiz07Points
    AirSpellPoints  As Wiz07Points
    EarthSpellPoints  As Wiz07Points
    MentalSpellPoints  As Wiz07Points
    DivineSpellPoints  As Wiz07Points
    ItemList(1 To 10) As Wiz07Item      'List of Items (not stowed)
    SwagBag(1 To 10) As Wiz07Item       'List of Stowed items
    
    Unknown2(1 To 64) As Byte           '???
    
    'Basic Statistics...
    STR As Byte                         'Strength
    INT As Byte                         'Intellegence (I.Q.)
    PIE As Byte                         'Piety
    VIT As Byte                         'Vitality
    DEX As Byte                         'Dexterity
    SPD As Byte                         'Speed
    PER As Byte                         'Personality
    KAR As Byte                         'Karma
    
    'Weaponry Skills...
    Wand As Byte
    Sword As Byte
    Axe As Byte
    Mace As Byte
    Pole As Byte
    Throwing As Byte
    Sling As Byte
    Bow As Byte
    Shield As Byte
    HandToHand As Byte
    
    'Physical Skills...
    Swimming As Byte
    Climbing As Byte
    Scouting As Byte
    Music As Byte
    Oratory As Byte
    Legerdemain As Byte
    Skulduggery As Byte
    Ninjutsu As Byte
    
    'Personal Skills...
    Firearms As Byte
    Reflextion As Byte
    SnakeSpeed As Byte
    EagleEye As Byte
    PowerStrike As Byte
    MindControl As Byte
    
    'Academia Skills...
    Artifacts As Byte
    Mythology As Byte
    Mapping As Byte
    Scribe As Byte
    Diplomacy As Byte
    Alchemy As Byte
    Theology As Byte
    Theosophy As Byte
    Thaumaturgy As Byte
    Kirijutsu As Byte
    
    Unknown3a(1 To 28) As Byte           '???
    NaturalArmorClass As Byte
    Unknown3b(1 To 7) As Byte            '???
    
    'Temporary until I figure out how best to do these bit-strings...
    Aptitude(1 To 96) As Byte           'Aptitude - I don't remember how I determined this...
    SpellBooks(1 To 12) As Byte         'Need to mask as bits...
    
    Unknown4(1 To 12) As Byte           '???
    PictureCode As Byte
    Race As Byte
    Gender As Byte
    Profession As Byte
    AgeInWeeks As Byte
    ConditionCode As Byte               '??? Under Investigation...
    Unknown5(1 To 12) As Byte           '???
End Type

Private Spells(1 To 96) As String
Public Sub DumpWiz07(ByVal strFile As String)
    Dim i As Long
    Dim j As Long
    Dim Unit As Integer
    Dim BytesReadSoFar As Long
    Dim errorCode As Long
    Dim Sunwolf As Wiz07Character
    
    'Fire Spell Book...
    Spells(1) = "Energy Blast"
    Spells(2) = "Blinding Flash"
    Spells(3) = "Psionic Fire"
    Spells(4) = "Fireball"
    Spells(5) = "Fire Shield"
    Spells(6) = "Dazzling Lights"
    Spells(7) = "Fire Bomb"
    Spells(8) = "Lightning"
    Spells(9) = "Prismic Missile"
    Spells(10) = "Firestorm"
    Spells(11) = "NuclearBlast"
    'Water Spell Book...
    Spells(12) = "Chilling Touch"
    Spells(13) = "Stamina"
    Spells(14) = "Terror"
    Spells(15) = "Weaken"
    Spells(16) = "Slow"
    Spells(17) = "Haste"
    Spells(18) = "Cure Paralysis"
    Spells(19) = "IceShield"
    Spells(20) = "Restfull"
    Spells(21) = "IceBall"
    Spells(22) = "Paralyze"
    Spells(23) = "Superman"
    Spells(24) = "Deep Freeze"
    Spells(25) = "Draining Cloud"
    Spells(26) = "CureDisease"
    'Air Spell Book...
    Spells(27) = "Poison"
    Spells(28) = "Missile Shield"
    Spells(29) = "Shrill Sound"
    Spells(30) = "StinkBomb"
    Spells(31) = "Air Pocket"
    Spells(32) = "Silence"
    Spells(33) = "Poison Gas"
    Spells(34) = "Cure Poison"
    Spells(35) = "Whirlwind"
    Spells(36) = "Purify Air"
    Spells(37) = "Deadly Poison"
    Spells(38) = "Levitate"
    Spells(39) = "Toxic Vapors"
    Spells(40) = "Noxious Fumes"
    Spells(41) = "Asphyxiation"
    Spells(42) = "Deadly Air"
    Spells(43) = "Death Cloud"
    'Earth Spell Book...
    Spells(44) = "Acid Splash"
    Spells(45) = "Itching Skin"
    Spells(46) = "Armor Shield"
    Spells(47) = "Direction"
    Spells(48) = "Knock Knock"
    Spells(49) = "Blades"
    Spells(50) = "Armorplate"
    Spells(51) = "Web"
    Spells(52) = "Whipping Rocks"
    Spells(53) = "Acid Bomb"
    Spells(54) = "Armormelt"
    Spells(55) = "Crush"
    Spells(56) = "Create Life"
    Spells(57) = "Cure Stone"
    'Mental Spell Book...
    Spells(58) = "MentalAttack"
    Spells(59) = "Sleep"
    Spells(60) = "Bless"
    Spells(61) = "Charm"
    Spells(62) = "Cure Lesser Cnd"
    Spells(63) = "Divine Trap"
    Spells(64) = "Detect Secret"
    Spells(65) = "Identify"
    Spells(66) = "Confusion"
    Spells(67) = "Watchbells"
    Spells(68) = "Hold Monsters"
    Spells(69) = "Mindread"
    Spells(70) = "Sane Mind"
    Spells(71) = "Psionic Blast"
    Spells(72) = "Illusion"
    Spells(73) = "Wizards Eye"
    Spells(74) = "Spooks"
    Spells(75) = "Death"
    Spells(76) = "Locate Object"
    Spells(77) = "Mind Flay"
    Spells(78) = "Find Person"
    'Divine Spell Book...
    Spells(79) = "Heal Wounds"
    Spells(80) = "Make Wounds"
    Spells(81) = "Magic Missile"
    Spells(82) = "Dispell Undead"
    Spells(83) = "Enchanted Blade"
    Spells(84) = "Blink"
    Spells(85) = "Magic Screen"
    Spells(86) = "Conjuration"
    Spells(87) = "Anti Magic"
    Spells(88) = "Remove Curse"
    Spells(89) = "Healfull"
    Spells(90) = "Lifesteal"
    Spells(91) = "Astral Gate"
    Spells(92) = "Zap Undead"
    Spells(93) = "Recharge"
    Spells(94) = "Word Of Death"
    Spells(95) = "Resurrection"
    Spells(96) = "Death Wish"

    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open strFile For Binary Access Read Write Lock Read Write As #Unit
    Get #Unit, &H3635, Sunwolf

    With Sunwolf
        Debug.Print "Name:              " & vbTab & .Name
        Debug.Print "Unknown Region #1 (4 bytes): "
        Debug.Print strHex(.Unknown1, 4) & vbCrLf
        
        Debug.Print vbCrLf & "Secondary Statistics..."
        Debug.Print "Experience Points: " & vbTab & .EXP & vbTab & "0x" & Hex(.EXP)
        Debug.Print "Monster Kills:     " & vbTab & .MKS & vbTab & "0x" & Hex(.MKS)
        Debug.Print "Gold Pieces:       " & vbTab & .GP & vbTab & "0x" & Hex(.GP)
        Debug.Print "Hit Points:        " & vbTab & strPoints(.HP)
        Debug.Print "Stamina:           " & vbTab & strPoints(.STM)
        Debug.Print "Carrying Capacity: " & vbTab & strPoints(.CC)
        Debug.Print "Level:             " & vbTab & .Level
        Debug.Print "Lives:             " & vbTab & .Lives
        
        Debug.Print vbCrLf & "Spell Points..."
        Debug.Print "Fire:              " & vbTab & strPoints(.FireSpellPoints)
        Debug.Print "Water:             " & vbTab & strPoints(.WaterSpellPoints)
        Debug.Print "Air:               " & vbTab & strPoints(.AirSpellPoints)
        Debug.Print "Earth:             " & vbTab & strPoints(.EarthSpellPoints)
        Debug.Print "Mental:            " & vbTab & strPoints(.MentalSpellPoints)
        Debug.Print "Divine:            " & vbTab & strPoints(.DivineSpellPoints)
        
        Debug.Print vbCrLf & "List of Items (not stowed)..."
        For i = 1 To 10
            Debug.Print strItem(.ItemList(i))
        Next i
        
        Debug.Print vbCrLf & "List of Stowed items..."
        For i = 1 To 10
            Debug.Print strItem(.SwagBag(i))
        Next i
        
        Debug.Print vbCrLf & "Unknown Region #2 (64 bytes):"
        Debug.Print strHex(.Unknown2, 64) & vbCrLf
        
        Debug.Print vbCrLf & "Basic Statistics..."
        Debug.Print "Strength:          " & vbTab & .STR       '
        Debug.Print "Intellegence:      " & vbTab & .INT
        Debug.Print "Piety:             " & vbTab & .PIE
        Debug.Print "Vitality:          " & vbTab & .VIT
        Debug.Print "Dexterity:         " & vbTab & .DEX
        Debug.Print "Speed:             " & vbTab & .SPD
        Debug.Print "Personality:       " & vbTab & .PER
        Debug.Print "Karma:             " & vbTab & .KAR
        
        Debug.Print vbCrLf & "Weaponry Skills..."
        Debug.Print "Wand:              " & vbTab & .Wand
        Debug.Print "Sword:             " & vbTab & .Sword
        Debug.Print "Axe:               " & vbTab & .Axe
        Debug.Print "Mace:              " & vbTab & .Mace
        Debug.Print "Pole:              " & vbTab & .Pole
        Debug.Print "Throwing:          " & vbTab & .Throwing
        Debug.Print "Sling:             " & vbTab & .Sling
        Debug.Print "Bow:               " & vbTab & .Bow
        Debug.Print "Shield:            " & vbTab & .Shield
        Debug.Print "HandToHand:        " & vbTab & .HandToHand
        
        Debug.Print vbCrLf & "Physical Skills..."
        Debug.Print "Swimming:          " & vbTab & .Swimming
        Debug.Print "Climbing:          " & vbTab & .Climbing
        Debug.Print "Scouting:          " & vbTab & .Scouting
        Debug.Print "Music:             " & vbTab & .Music
        Debug.Print "Oratory:           " & vbTab & .Oratory
        Debug.Print "Legerdemain:       " & vbTab & .Legerdemain
        Debug.Print "Skulduggery:       " & vbTab & .Skulduggery
        Debug.Print "Ninjutsu:          " & vbTab & .Ninjutsu
        
        Debug.Print vbCrLf & "Personal Skills..."
        Debug.Print "Firearms:          " & vbTab & .Firearms
        Debug.Print "Reflextion:        " & vbTab & .Reflextion
        Debug.Print "SnakeSpeed:        " & vbTab & .SnakeSpeed
        Debug.Print "EagleEye:          " & vbTab & .EagleEye
        Debug.Print "PowerStrike:       " & vbTab & .PowerStrike
        Debug.Print "MindControl:       " & vbTab & .MindControl
        
        Debug.Print vbCrLf & "Academia Skills..."
        Debug.Print "Artifacts:         " & vbTab & .Artifacts
        Debug.Print "Mythology:         " & vbTab & .Mythology
        Debug.Print "Mapping:           " & vbTab & .Mapping
        Debug.Print "Scribe:            " & vbTab & .Scribe
        Debug.Print "Diplomacy:         " & vbTab & .Diplomacy
        Debug.Print "Alchemy:           " & vbTab & .Alchemy
        Debug.Print "Theology:          " & vbTab & .Theology
        Debug.Print "Theosophy:         " & vbTab & .Theosophy
        Debug.Print "Thaumaturgy:       " & vbTab & .Thaumaturgy
        Debug.Print "Kirijutsu:         " & vbTab & .Kirijutsu
        
        Debug.Print vbCrLf & "Unknown Region #3a (28 bytes):"
        Debug.Print strHex(.Unknown3a, 28) & vbCrLf
        
        Debug.Print "'Natural' Armor Class:" & vbTab & .NaturalArmorClass
        
        Debug.Print vbCrLf & "Unknown Region #3b (7 bytes):"
        Debug.Print strHex(.Unknown3b, 7) & vbCrLf
        
        Debug.Print vbCrLf & "Aptitute (96 bytes):"
        Debug.Print strHex(.Aptitude, 96) & vbCrLf   'Aptitude - I don't remember how I determined this...
        
        Debug.Print vbCrLf & "SpellBooks..."
        For j = 1 To 12
            For i = 1 To 8
                Debug.Print vbTab & strSpell(((j - 1) * 8) + i, .SpellBooks(j), i - 1)
            Next i
        Next j
    
        Debug.Print vbCrLf & "Unknown Region #4 (12 bytes):"
        Debug.Print strHex(.Unknown4, 12) & vbCrLf
        
        Debug.Print "PictureCode:       " & vbTab & .PictureCode
        Debug.Print "Race:              " & vbTab & strRace(.Race)
        Debug.Print "Gender:            " & vbTab & strGender(.Gender)
        Debug.Print "Profession:        " & vbTab & strProfession(.Profession)
        Debug.Print "Age:               " & vbTab & .AgeInWeeks \ 52 & " (" & .AgeInWeeks & " weeks)"
        Debug.Print "ConditionCode:     " & vbTab & strCondition(.ConditionCode)
        
        Debug.Print vbCrLf & "Unknown Region #5 (12 bytes):"
        Debug.Print strHex(.Unknown5, 12) & vbCrLf
    
        Debug.Print "Unknown Recap:"
        Debug.Print "Unknown Region #1 (4 bytes): "
        Debug.Print strHex(.Unknown1, 4) & vbCrLf
        Debug.Print vbCrLf & "Unknown Region #2 (64 bytes):"
        Debug.Print strHex(.Unknown2, 64) & vbCrLf
        Debug.Print vbCrLf & "Unknown Region #3a (28 bytes):"
        Debug.Print strHex(.Unknown3a, 28) & vbCrLf
        Debug.Print vbCrLf & "Unknown Region #3b (7 bytes):"
        Debug.Print strHex(.Unknown3b, 7) & vbCrLf
        Debug.Print vbCrLf & "Unknown Region #4 (12 bytes):"
        Debug.Print strHex(.Unknown4, 12) & vbCrLf
        Debug.Print vbCrLf & "Unknown Region #5 (12 bytes):"
        Debug.Print strHex(.Unknown5, 12) & vbCrLf
        Debug.Print vbCrLf & "Aptitute (96 bytes):"
        Debug.Print strHex(.Aptitude, 96) & vbCrLf   'Aptitude - I don't remember how I determined this...
    End With
    Close #Unit
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "DumpWiz07"
    Exit Sub
    Resume Next
End Sub
Public Sub PopulateWiz07Condition(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz07ConditionMapMax
            .AddItem strCondition(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz07Gender(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To 1
            .AddItem strGender(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz07Profession(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz07ProfessionMapMax
            .AddItem strProfession(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz07Race(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        For i = 0 To Wiz07RaceMapMax
            .AddItem strRace(i), CInt(i)
        Next i
    End With
End Sub
Public Sub PopulateWiz07SkillsAcademia(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        .AddItem "Artifacts"
        .AddItem "Mythology"
        .AddItem "Mapping"
        .AddItem "Scribe"
        .AddItem "Diplomacy"
        .AddItem "Alchemy"
        .AddItem "Theology"
        .AddItem "Theosophy"
        .AddItem "Thaumaturgy"
        .AddItem "Kirijutsu"
    End With
End Sub
Public Sub PopulateWiz07SkillsPersonal(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        .AddItem "Firearms"
        .AddItem "Reflextion"
        .AddItem "Snake Speed"
        .AddItem "Eagle Eye"
        .AddItem "Power Strike"
        .AddItem "Mind Control"
    End With
End Sub
Public Sub PopulateWiz07SkillsPhysical(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        .AddItem "Swimming"
        .AddItem "Climbing"
        .AddItem "Scouting"
        .AddItem "Music"
        .AddItem "Oratory"
        .AddItem "Legerdemain"
        .AddItem "Skulduggery"
        .AddItem "Ninjutsu"
    End With
End Sub
Public Sub PopulateWiz07SkillsWeaponry(x As ComboBox)
    Dim i As Byte
    With x
        .Clear
        .AddItem "Wand & Dagger"
        .AddItem "Sword"
        .AddItem "Axe"
        .AddItem "Mace & Flail"
        .AddItem "Pole & Staff"
        .AddItem "Throwing"
        .AddItem "Sling"
        .AddItem "Bows"
        .AddItem "Shield"
        .AddItem "Hands & Feet"
    End With
End Sub
Public Sub ReadWiz07(ByVal strFile As String, xCharacters() As Wiz07Character)
    Dim i As Long
    Dim j As Long
    Dim Unit As Integer
    Dim BytesReadSoFar As Long
    Dim errorCode As Long
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open strFile For Binary Access Read Write Lock Read Write As #Unit
    Get #Unit, &H3635, xCharacters(1)
    For i = 2 To 6
        Get #Unit, , xCharacters(i)
    Next i
    Close #Unit
    
    For i = 1 To 6
        xCharacters(i).Name = Replace(xCharacters(i).Name, Chr(0), " ")
    Next i
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "ReadWiz07"
    Exit Sub
    Resume Next
End Sub
Private Function strCondition(ByVal x As Byte) As String
'char *ConditionMap[] = {"OK","Asleep","Blinded","Dead","Poisoned","Stoned","Insane","Afraid","Nauseated","Paralyzed","Irritated","Diseased"};
    Select Case x
        Case 0
            strCondition = "OK"
        Case 1
            strCondition = "Asleep"
        Case 2
            strCondition = "Blinded"
        Case 3
            strCondition = "Dead"
        Case 4
            strCondition = "Poisoned"
        Case 5
            strCondition = "Stoned"
        Case 6
            strCondition = "Insane"
        Case 7
            strCondition = "Afraid"
        Case 8
            strCondition = "Nauseated"
        Case 9
            strCondition = "Paralyzed"
        Case 10
            strCondition = "Irritated"
        Case 11
            strCondition = "Diseased"
        Case Else
            strCondition = "Unknown"
    End Select
End Function
Private Function strGender(ByVal x As Byte) As String
    Select Case x
        Case 0
            strGender = "Male"
        Case 1
            strGender = "Female"
        Case Else
            strGender = "Unknown"
    End Select
End Function
Private Function strHex(ByRef xBytes() As Byte, nBytes As Integer) As String
    Dim i As Integer

    strHex = ""
    For i = 1 To nBytes
        strHex = strHex & Format(Hex(xBytes(i)), "00")
        If i Mod 4 = 0 Then strHex = strHex & " "
        If i Mod 32 = 0 Then strHex = strHex & vbCrLf
    Next i
End Function
Private Function strProfession(ByVal x As Byte) As String
'char *ProfessionMap[] = {"Fighter","Mage","Priest","Thief","Ranger","Alchemist","Bard","Psionic","Valkyrie","Bishop","Lord","Samurai","Monk","Ninja"};
    Select Case x
        Case 0
            strProfession = "Fighter"
        Case 1
            strProfession = "Mage"
        Case 2
            strProfession = "Priest"
        Case 3
            strProfession = "Thief"
        Case 4
            strProfession = "Ranger"
        Case 5
            strProfession = "Alchemist"
        Case 6
            strProfession = "Bard"
        Case 7
            strProfession = "Psionic"
        Case 8
            strProfession = "Valkyrie"
        Case 9
            strProfession = "Bishop"
        Case 10
            strProfession = "Lord"
        Case 11
            strProfession = "Samurai"
        Case 12
            strProfession = "Monk"
        Case 13
            strProfession = "Ninja"
        Case Else
            strProfession = "Unknown"
    End Select
End Function
Private Function strRace(ByVal x As Byte) As String
'char *RaceMap[] = {"Human","Elf","Dwarf","Gnome","Hobbit","Faerie","Lizardman","Dracon","Rawulf","Felpurr","Mook"};
    Select Case x
        Case 0
            strRace = "Human"
        Case 1
            strRace = "Elf"
        Case 2
            strRace = "Dwarf"
        Case 3
            strRace = "Gnome"
        Case 4
            strRace = "Hobbit"
        Case 5
            strRace = "Faerie"
        Case 6
            strRace = "Lizardman"
        Case 7
            strRace = "Dracon"
        Case 8
            strRace = "Rawulf"
        Case 9
            strRace = "Felpurr"
        Case 10
            strRace = "Mook"
        Case Else
            strRace = "Unknown"
    End Select
End Function
Private Function strItem(x As Wiz07Item) As String
    strItem = vbTab & "Code: " & x.ItemCode
End Function
Private Function strPoints(x As Wiz07Points) As String
    strPoints = x.Current & "/" & x.Maximum
End Function
Private Function strSpell(Spell As Integer, Data As Byte, Offset As Integer) As String
    Dim Temp As String
    If (Data And 2 ^ Offset) = 2 ^ Offset Then Temp = "[X]" Else Temp = "[ ]"
    strSpell = Temp & " " & Spells(Spell) '& vbTab & "[Spell: " & Spell & "; Data: " & Hex(Data) & "; Offset: " & Offset & "]"
End Function

