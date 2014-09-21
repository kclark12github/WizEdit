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
Type Item
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
Type Points
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
Type Character
    Name As String * 8                  'Null Terminated
    Unknown1(1 To 4) As Byte            '???
    
    'Secondary Statistics...
    EXP As Long                         'Experience Points
    MKS As Long                         'Monster Kills
    GP As Long                          'Gold Pieces
    HP As Points                        'Hit Points
    STA As Points                       'Stamina
    CC As Points                        'Carrying Capacity
    Level As Integer                    'Level (Duh)
    Lives As Integer                    'Number of Lives Used
    
    'Spell Points...
    FireSpellPoints  As Points
    WaterSpellPoints  As Points
    AirSpellPoints  As Points
    EarthSpellPoints  As Points
    MentalSpellPoints  As Points
    DivineSpellPoints  As Points
    ItemList(1 To 10) As Item           'List of Items (not stowed)
    SwagBag(1 To 10) As Item            'List of Stowed items
    
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
    
    Unknown3(1 To 36) As Byte           '???
    
    'Temporary until I figure out how best to do these bit-strings...
    Aptitude(1 To 96) As Byte           'Aptitude - I don't remember how I determined this...
    SpellBooks(1 To 48) As Integer      'Need to mask as bits...
    
    'Fire Spell Book...
'Global Const maskWiz07EnergyBlast = &H1
'Global Const maskWiz07BlindingFlash = &H2
'Global Const maskWiz07PsionicFire = &H3
'Global Const maskWiz07Fireball = &H4
'Global Const maskWiz07FireShield = &H5
'Global Const maskWiz07DazzlingLights = &H6
'Global Const maskWiz07FireBomb = &H7
'Global Const maskWiz07Lightning = &H8
'Global Const maskWiz07PrismicMissile = &H9
'Global Const maskWiz07Firestorm = &HA
'Global Const maskWiz07NuclearBlast = &HB
    'Water Spell Book...
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
    'Air Spell Book...
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
    'Earth Spell Book...
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
    'Mental Spell Book...
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
    '? Spell Book...
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

    Unknown4(1 To 12) As Byte           '???
    PictureCode As Byte
    Race As Byte
    Gender As Byte
    Profession As Byte
    Alive As Byte                       '??? Under Investigation...
    ConditionCode As Byte               '??? Under Investigation...
    Unknown5(1 To 12) As Byte           '???
End Type

'#define ItemMapMax 569
'#define RaceMapMax 10
'#define ProfessionMapMax 13
'#define ConditionMapMax 11
'#define SpellMapMax 95
Global Const Wiz07ItemMapMax = 569
Global Const Wiz07RaceMapMax = 10
Global Const Wiz07ProfessionMapMax = 13
Global Const Wiz07ConditionMapMax = 11
Global Const Wiz07SpellMapMax = 95

'extern char *RaceMap[];
'extern char *ProfessionMap[];
'extern char *ConditionMap[];
'extern char *SpellMap[];
'
'/* Function Prototypes */
'
'int ReadDBS(char *FileName, char **Buffer, long int *FileSize);
'int WriteDBS(char *FileName, char *Buffer, long int FileSize);
'char *MapItemCode(short int Code);
'void HexDump(unsigned char *p, int Bytes);
'void HexDumpf(FILE *FileUnit, unsigned char *p, int Bytes);
'
'
Private Function strPoints(x As Points) As String
    strPoints = x.Current & "/" & x.Maximum
End Function
Private Function strItem(x As Item) As String
    strItem = vbTab & "Code: " & x.ItemCode
End Function
Public Sub DumpWiz07(ByVal strFile As String)
    Dim i As Long
    Dim Unit As Integer
    Dim BytesReadSoFar As Long
    Dim errorCode As Long
    Dim Sunwolf As Character
    
    On Error GoTo ErrorHandler
    Unit = FreeFile
    Open strFile For Binary Access Read Write Lock Read Write As #Unit
    Get #Unit, &H3635, Sunwolf

    With Sunwolf
        Debug.Print "Name:              " & vbTab & .Name
        Debug.Print "Unknown Region #1: " & vbTab & Hex(.Unknown1(1)) & " " & Hex(.Unknown1(2)) & " " & Hex(.Unknown1(3)) & " " & Hex(.Unknown1(4))
        
        Debug.Print vbCrLf & "Secondary Statistics..."
        Debug.Print "Experience Points: " & vbTab & .EXP & vbTab & "0x" & Hex(.EXP)
        Debug.Print "Monster Kills:     " & vbTab & .MKS & vbTab & "0x" & Hex(.MKS)
        Debug.Print "Gold Pieces:       " & vbTab & .GP & vbTab & "0x" & Hex(.GP)
        Debug.Print "Hit Points:        " & vbTab & strPoints(.HP)
        Debug.Print "Stamina:           " & vbTab & strPoints(.STA)
        Debug.Print "Carrying Capacity: " & vbTab & strPoints(.CC)
        Debug.Print "Level:             " & vbTab & .Level
        Debug.Print "Lives:             " & vbTab & .Lives
        
        Debug.Print vbCrLf & "Spell Points..."
        Debug.Print "Fire" & vbTab & strPoints(.FireSpellPoints)
        Debug.Print "Water" & vbTab & strPoints(.WaterSpellPoints)
        Debug.Print "Air" & vbTab & strPoints(.AirSpellPoints)
        Debug.Print "Earth" & vbTab & strPoints(.EarthSpellPoints)
        Debug.Print "Mental" & vbTab & strPoints(.MentalSpellPoints)
        Debug.Print "Divine" & vbTab & strPoints(.DivineSpellPoints)
        
        Debug.Print vbCrLf & "List of Items (not stowed)..."
        For i = 1 To 10
            Debug.Print strItem(.ItemList(i))
        Next i
        
        Debug.Print vbCrLf & "List of Stowed items..."
        For i = 1 To 10
            Debug.Print strItem(.SwagBag(i))
        Next i
        
        Debug.Print vbCrLf & "Unknown Region #2 (64 bytes) [not shown]"
        'Unknown2(64) As Byte                '???
        
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
        
        Debug.Print vbCrLf & "Unknown Region #3 (36 bytes) [not shown]"
        'Unknown3(36) As Byte                '???
        
        Debug.Print vbCrLf & "Aptitute (96 bytes) [not shown]"
        'Aptitude(96) As Byte                'Aptitude - I don't remember how I determined this...
        
        'Temporary until I figure out how best to do these bit-strings...
        Debug.Print vbCrLf & "SpellBooks (96 bits) [not shown]"
        'SpellBooks(24) As Long              'Need to mask as bits...
        For i = 0 To 15
            Debug.Print "SpellBooks(1):" & i & " is " & ((.SpellBooks(1) And i) = i)
        Next i
        'Fire Spell Book...
    'Global Const maskWiz07EnergyBlast = &H1
    'Global Const maskWiz07BlindingFlash = &H2
    'Global Const maskWiz07PsionicFire = &H3
    'Global Const maskWiz07Fireball = &H4
    'Global Const maskWiz07FireShield = &H5
    'Global Const maskWiz07DazzlingLights = &H6
    'Global Const maskWiz07FireBomb = &H7
    'Global Const maskWiz07Lightning = &H8
    'Global Const maskWiz07PrismicMissile = &H9
    'Global Const maskWiz07Firestorm = &HA
    'Global Const maskWiz07NuclearBlast = &HB
        'Water Spell Book...
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
        'Air Spell Book...
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
        'Earth Spell Book...
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
        'Mental Spell Book...
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
        'Divine Spell Book...
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
    
        Debug.Print vbCrLf & "Unknown Region #4 (12 bytes) [not shown]"
        'Unknown4(12) As Byte                '???
        
        Debug.Print "PictureCode:       " & vbTab & .PictureCode
        Debug.Print "Race:              " & vbTab & .Race
        Debug.Print "Gender:            " & vbTab & .Gender
        Debug.Print "Profession:        " & vbTab & .Profession
        Debug.Print "?Alive?:           " & vbTab & .Alive
        Debug.Print "ConditionCode:     " & vbTab & .ConditionCode
        
        Debug.Print vbCrLf & "Unknown Region #5 (12 bytes) [not shown]"
        'Unknown5(12) As Byte                '???
    End With
    Close #Unit
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbExclamation, "DumpWiz07"
    Exit Sub
    Resume Next
End Sub
