#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <alloc.h>
#include <io.h>
#include <conio.h>

#include <wizardry.h>
#include <input.h>
#include <itemmap.h>

int	titlecolor		=	MAGENTA;
int	headercolor		=	LIGHTRED;
int	promptcolor		=	CYAN;
int	errorcolor		=	RED;
int   warningcolor   =  LIGHTMAGENTA;
int	selectcolor		=	BROWN;
int	menucolor		=	GREEN;
int	datecolor		=	LIGHTBLUE;
int	timecolor		=	LIGHTGRAY;
int	screencolor		=	LIGHTGREEN;
int	inputcolor		=	CYAN;
int	displaycolor   =  YELLOW;

char *ItemMap[] = {""};
char *RaceMap[] = {"Human","Elf","Dwarf","Gnome","Hobbit","Faerie","Lizardman","Dracon","Rawulf","Felpurr","Mook"};
char *ProfessionMap[] = {"Fighter","Mage","Priest","Thief","Ranger","Alchemist","Bard","Psionic","Valkyrie","Bishop","Lord","Samurai","Monk","Ninja"};
char *ConditionMap[] = {"OK","Asleep","Blinded","Dead","Poisoned","Stoned","Insane","Afraid","Nauseated","Paralyzed","Irritated","Diseased"};
char *SpellMap[] = {
     /* Fire */     "EnergyBlast","BlindingFlash","PsionicFire","Fireball",
                    "FireShield","DazzlingLights","FireBomb","Lightning",
                    "PrismicMissile","Firestorm","NuclearBlast",
     /* Water */    "ChillingTouch","Stamina","Terror","Weaken","Slow",
                    "Haste","CureParalysis","IceShield","Restfull","IceBall",
                    "Paralyze","Superman","Deepfreeze","DrainingCloud",
                    "CureDisease",
     /* Air */      "Poison","MissileShield","ShrillSound","StinkBomb",
                    "AirPocket","Silence","PoisonGas","CurePoison",
                    "Whirlwind","PurifyAir","DeadlyPoison","Levitate",
                    "ToxicVapors","NoxiousFumes","Asphyxiation","DeadlyAir",
                    "DeathCloud",
     /* Earth */    "AcidSplash","ItchingSkin","ArmorShield","Direction",
                    "KnockKnock","Blades","Armorplate","Web","WhippingRocks",
                    "AcidBomb","Armormelt","Crush","CreateLife","CureStone",
     /* Mental */   "MentalAttack","Sleep","Bless","Charm","CureLesserCnd",
                    "DivineTrap","DetectSecret","Identify","Confusion",
                    "Watchbells","HoldMonsters","Mindread","SaneMind",
                    "PsionicBlast","Illusion","WizardsEye","Spooks","Death",
                    "LocateObject","MindFlay","FindPerson",
     /* Divine */   "HealWounds","MakeWounds","MagicMissile","DispellUndead",
                    "EnchantedBlade","Blink","MagicScreen","Conjuration",
                    "AntiMagic","RemoveCurse","Healfull","Lifesteal",
                    "AstralGate","ZapUndead","Recharge","WordOfDeath",
                    "Resurrection","DeathWish"};

int ReadDBS(char *FileName, char **Buffer, long int *FileSize)
{
   short int Status = FALSE;
   long BytesLeft = 0,
        Bytes = 0;
   FILE *DataFile = NULL;

   if (!(DataFile = fopen(FileName, "rb")))
   {
      SystemError("Cannot open data file, %s.", FileName);
      goto EXIT_LABEL;
   }
   if ((*FileSize = filelength(fileno(DataFile))) == -1L)
   {
      PrintError("Cannot determine size of %s.", FileName);
      goto EXIT_LABEL;
   }
   BytesLeft = coreleft();
   if (*FileSize > BytesLeft)
   {
      PrintError("Cannot allocate enough memory (%ld bytes) to read %s.",
                   BytesLeft, FileName);
      goto EXIT_LABEL;
   }
   if ((*Buffer = (char *)malloc(*FileSize)) == NULL)
   {
      PrintError("Could not allocate %ld bytes for %s.", FileName);
      goto EXIT_LABEL;
   }

   printf("   Reading %ld bytes from \"%s\"...", *FileSize, FileName);
   if ((Bytes = fread((void *) *Buffer, 1, *FileSize, DataFile)) != *FileSize)
   {
      PrintError("Read only %ld of %ld bytes from %s.",
                   Bytes, *FileSize, FileName);
      goto EXIT_LABEL;
   }

   Status = TRUE;

EXIT_LABEL:
   fclose(DataFile);
   return Status;
}

int WriteDBS(char *FileName, char *Buffer, long int FileSize)
{
   short int Status = FALSE;
   long Bytes = 0;
   FILE *DataFile = NULL;

   if (!(DataFile = fopen(FileName, "wb")))
   {
      SystemError("Cannot open data file, %s.", FileName);
      goto EXIT_LABEL;
   }

   printf("   Writing %ld bytes to %s...", FileSize, FileName);
   if ((Bytes = fwrite((void *) Buffer, 1, FileSize, DataFile)) != FileSize)
   {
      SystemError("Wrote only %d of %d bytes to %s.", Bytes, FileSize, FileName);
      goto EXIT_LABEL;
   }

   Status = TRUE;

EXIT_LABEL:
   fclose(DataFile);
   return Status;
}

char *MapItemCode(short int Code)
{
   static char HexString[5];
   if (Code < ItemMapMax)
      return Items[Code].Name;

   sprintf(HexString, "%04X", Code);
   return HexString;
}

void HexDump(unsigned char *p, int Bytes)
{
   short int i = 0;

   for (i = 1; i <= Bytes; i++)
   {
      printf("%02X", *p);
      p++;
      if (!(i % 32) && i != Bytes)
         printf("\n");
      else
          if (!(i % 4))
             printf(" ");
   }
   printf("\n");
}
void HexDumpf(FILE *FileUnit, unsigned char *p, int Bytes)
{
   short int i = 0;

   for (i = 1; i <= Bytes; i++)
   {
      fprintf(FileUnit, "%02X", *p);
      p++;
      if (!(i % 32) && i != Bytes)
         fprintf(FileUnit, "\n");
      else
          if (!(i % 4))
             fprintf(FileUnit, " ");
   }
   fprintf(FileUnit, "\n");
}

