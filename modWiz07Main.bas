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
'struct POINTS
'{
'   unsigned short int   Current;
'   unsigned short int   Maximum;
'};
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
'
'#define ItemMapMax 569
'#define RaceMapMax 10
'#define ProfessionMapMax 13
'#define ConditionMapMax 11
'#define SpellMapMax 95
'
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

