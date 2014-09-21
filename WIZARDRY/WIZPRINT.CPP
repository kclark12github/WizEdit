#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <io.h>
#include <conio.h>

#include <wizardry.h>
#include <input.h>

#define TRUE (0l == 0l)
#define FALSE (0l != 0l)

char *GetTimeString(void);
char *GetDateString(void);

struct ProgramInfo Program;
char *Buffer = NULL;
long FileSize = 0;

void PrintAttributes(struct Character *cp)
{
   printf("\nAttributes:   \t\tExperience:   \t%10ld\n", cp->EXP);
   printf("   STR:\t%02d \t\tMonster Kills:\t%10ld\n", cp->Attributes.STR, cp->MKS);
   printf("   INT:\t%02d \t\tGold Pieces:  \t%10ld\n", cp->Attributes.INT, cp->GP);
   printf("   PIE:\t%02d \t\tHit Points:   \t%d of %d\n", cp->Attributes.PIE,
      cp->HP.Current, cp->HP.Maximum);
   printf("   VIT:\t%02d \t\tStamina:      \t%d of %d\n", cp->Attributes.VIT,
      cp->STA.Current, cp->STA.Maximum);
   printf("   DEX:\t%02d \t\tCapacity:     \t%5.1f lbs of %5.1f lbs\n",
      cp->Attributes.DEX,
      (float)cp->CC.Current/10.0, (float)cp->CC.Maximum/10.0);
   printf("   SPD:\t%02d \t\t\n", cp->Attributes.SPD);
   printf("   PER:\t%02d \t\t\n", cp->Attributes.PER);
   printf("   KAR:\t%02d \t\t\n", cp->Attributes.KAR);
}

void PrintSpellPoints(struct Character *cp)
{
   printf("\nSpell Points:\n");
   printf("   Fire:  \t%d of %d\tEarth: \t%d of %d\n",
      cp->SpellPoints.Fire.Current, cp->SpellPoints.Fire.Maximum,
      cp->SpellPoints.Earth.Current, cp->SpellPoints.Earth.Maximum);
   printf("   Water: \t%d of %d\tMental:\t%d of %d\n",
      cp->SpellPoints.Water.Current, cp->SpellPoints.Water.Maximum,
      cp->SpellPoints.Mental.Current, cp->SpellPoints.Mental.Maximum);
   printf("   Air:   \t%d of %d\tDivine:\t%d of %d\n",
      cp->SpellPoints.Air.Current, cp->SpellPoints.Air.Maximum,
      cp->SpellPoints.Divine.Current, cp->SpellPoints.Divine.Maximum);
}

void PrintItems(struct Character *cp)
{
   short int i = 0;
   float f = 0.0;

   printf("\nItems:\n");
   for (i = 0; i < 10; i++)
   {
      f = (float)cp->ItemList[i].Weight/10.0;
      printf("   %d) %s\n", i, MapItemCode(cp->ItemList[i].ItemCode));
      if (cp->ItemList[i].ItemCode)
         printf("      Weight: %5.1f lbs; Count: %4d; Status: %02X; AC: %4u; Unknown: %04X\n",
                      f,
                      cp->ItemList[i].Count,
                      cp->ItemList[i].Status,
                      cp->ItemList[i].AC,
                      cp->ItemList[i].Unknown);
   }

   printf("\nSwag Bag:\n");
   for (i = 0; i < 10; i++)
   {
      f = (float)cp->SwagBag[i].Weight/10.0;
      printf("   %d) %s\n", i, MapItemCode(cp->SwagBag[i].ItemCode));
      if (cp->SwagBag[i].ItemCode)
         printf("      Weight: %5.1f lbs; Count: %4d; Status: %02X; AC: %4u; Unknown: %04X\n",
                      f,
                      cp->SwagBag[i].Count,
                      cp->SwagBag[i].Status,
                      cp->SwagBag[i].AC,
                      cp->SwagBag[i].Unknown);
   }
}

void PrintUnknown(struct Character *cp)
{
   printf("\n");
   printf("Unknown Region #1: %s\n", cp->Name);
   HexDump(cp->Unknown1, sizeof(cp->Unknown1)); printf("\n");
   printf("Unknown Region #2: %s\n", cp->Name);
   HexDump(cp->Unknown2, sizeof(cp->Unknown2));
   printf("Unknown Region #3: %s\n", cp->Name);
   HexDump(cp->Unknown3, sizeof(cp->Unknown3));
   printf("Unknown Region #4: %s\n", cp->Name);
   HexDump(cp->Unknown4, sizeof(cp->Unknown4));
   printf("Unknown Region #5: %s\n", cp->Name);
   HexDump(cp->Unknown5, sizeof(cp->Unknown5));
}

#define B_YN 'Y' : 'N'

void PrintSpells(struct Character *cp)
{
   printf("\nSpells:\n");
   printf("   Fire               Water              Air\n");
   printf("   EnergyBlast:    %c  ChillingTouch:  %c  Poison:         %c\n", cp->Spells.EnergyBlast ? B_YN, cp->Spells.ChillingTouch ? B_YN, cp->Spells.Poison ? B_YN);
   printf("   BlindingFlash:  %c  Stamina:        %c  MissileShield:  %c\n", cp->Spells.BlindingFlash ? B_YN, cp->Spells.Stamina ? B_YN, cp->Spells.MissileShield ? B_YN);
   printf("   PsionicFire:    %c  Terror:         %c  ShrillSound:    %c\n", cp->Spells.PsionicFire ? B_YN, cp->Spells.Terror ? B_YN, cp->Spells.ShrillSound ? B_YN);
   printf("   Fireball:       %c  Weaken:         %c  StinkBomb:      %c\n", cp->Spells.Fireball ? B_YN, cp->Spells.Weaken ? B_YN, cp->Spells.StinkBomb ? B_YN);
   printf("   FireShield:     %c  Slow:           %c  AirPocket:      %c\n", cp->Spells.FireShield ? B_YN, cp->Spells.Slow ? B_YN, cp->Spells.AirPocket ? B_YN);
   printf("   DazzlingLights: %c  Haste:          %c  Silence:        %c\n", cp->Spells.DazzlingLights ? B_YN, cp->Spells.Haste ? B_YN, cp->Spells.Silence ? B_YN);
   printf("   FireBomb:       %c  CureParalysis:  %c  PoisonGas:      %c\n", cp->Spells.FireBomb ? B_YN, cp->Spells.CureParalysis ? B_YN, cp->Spells.PoisonGas ? B_YN);
   printf("   Lightning:      %c  IceShield:      %c  CurePoison:     %c\n", cp->Spells.Lightning ? B_YN, cp->Spells.IceShield ? B_YN, cp->Spells.CurePoison ? B_YN);
   printf("   PrismicMissile: %c  Restfull:       %c  Whirlwind:      %c\n", cp->Spells.PrismicMissile ? B_YN, cp->Spells.Restfull ? B_YN, cp->Spells.Whirlwind ? B_YN);
   printf("   Firestorm:      %c  IceBall:        %c  PurifyAir:      %c\n", cp->Spells.Firestorm ? B_YN, cp->Spells.IceBall ? B_YN, cp->Spells.PurifyAir ? B_YN);
   printf("   NuclearBlast:   %c  Paralyze:       %c  DeadlyPoison:   %c\n", cp->Spells.NuclearBlast ? B_YN, cp->Spells.Paralyze ? B_YN, cp->Spells.DeadlyPoison ? B_YN);
   printf("                      Superman:       %c  Levitate:       %c\n", cp->Spells.Superman ? B_YN, cp->Spells.Levitate ? B_YN);
   printf("                      Deepfreeze:     %c  ToxicVapors:    %c\n", cp->Spells.Deepfreeze ? B_YN, cp->Spells.ToxicVapors ? B_YN);
   printf("                      DrainingCloud:  %c  NoxiousFumes:   %c\n", cp->Spells.DrainingCloud ? B_YN, cp->Spells.NoxiousFumes ? B_YN);
   printf("                      CureDisease:    %c  Asphyxiation:   %c\n", cp->Spells.CureDisease ? B_YN, cp->Spells.Asphyxiation ? B_YN);
   printf("                                         DeadlyAir:      %c\n", cp->Spells.DeadlyAir ? B_YN);
   printf("                                         DeathCloud:     %c\n", cp->Spells.DeathCloud ? B_YN);

   printf("\n   Earth              Mental             Divine\n");
   printf("   AcidSplash:     %c  MentalAttack:   %c  HealWounds:     %c\n", cp->Spells.AcidSplash ? B_YN, cp->Spells.MentalAttack ? B_YN, cp->Spells.HealWounds ? B_YN);
   printf("   ItchingSkin:    %c  Sleep:          %c  MakeWounds:     %c\n", cp->Spells.ItchingSkin ? B_YN, cp->Spells.Sleep ? B_YN, cp->Spells.MakeWounds ? B_YN);
   printf("   ArmorShield:    %c  Bless:          %c  MagicMissile:   %c\n", cp->Spells.ArmorShield ? B_YN, cp->Spells.Bless ? B_YN, cp->Spells.MagicMissile ? B_YN);
   printf("   Direction:      %c  Charm:          %c  DispellUndead:  %c\n", cp->Spells.Direction ? B_YN, cp->Spells.Charm ? B_YN, cp->Spells.DispellUndead ? B_YN);
   printf("   KnockKnock:     %c  CureLesserCnd:  %c  EnchantedBlade: %c\n", cp->Spells.KnockKnock ? B_YN, cp->Spells.CureLesserCnd ? B_YN, cp->Spells.EnchantedBlade ? B_YN);
   printf("   Blades:         %c  DivineTrap:     %c  Blink:          %c\n", cp->Spells.Blades ? B_YN, cp->Spells.DivineTrap ? B_YN, cp->Spells.Blink ? B_YN);
   printf("   Armorplate:     %c  DetectSecret:   %c  MagicScreen:    %c\n", cp->Spells.Armorplate ? B_YN, cp->Spells.DetectSecret ? B_YN, cp->Spells.MagicScreen ? B_YN);
   printf("   Web:            %c  Identify:       %c  Conjuration:    %c\n", cp->Spells.Web ? B_YN, cp->Spells.Identify ? B_YN, cp->Spells.Conjuration ? B_YN);
   printf("   WhippingRocks:  %c  Confusion:      %c  AntiMagic:      %c\n", cp->Spells.WhippingRocks ? B_YN, cp->Spells.Confusion ? B_YN, cp->Spells.AntiMagic ? B_YN);
   printf("   AcidBomb:       %c  Watchbells:     %c  RemoveCurse:    %c\n", cp->Spells.AcidBomb ? B_YN, cp->Spells.Watchbells ? B_YN, cp->Spells.RemoveCurse ? B_YN);
   printf("   Armormelt:      %c  HoldMonsters:   %c  Healfull:       %c\n", cp->Spells.Armormelt ? B_YN, cp->Spells.HoldMonsters ? B_YN, cp->Spells.Healfull ? B_YN);
   printf("   Crush:          %c  Mindread:       %c  Lifesteal:      %c\n", cp->Spells.Crush ? B_YN, cp->Spells.Mindread ? B_YN, cp->Spells.Lifesteal ? B_YN);
   printf("   CreateLife:     %c  SaneMind:       %c  AstralGate:     %c\n", cp->Spells.CreateLife ? B_YN, cp->Spells.SaneMind ? B_YN, cp->Spells.AstralGate ? B_YN);
   printf("   CureStone:      %c  PsionicBlast:   %c  ZapUndead:      %c\n", cp->Spells.CureStone ? B_YN, cp->Spells.PsionicBlast ? B_YN, cp->Spells.ZapUndead ? B_YN);
   printf("                      Illusion:       %c  Recharge:       %c\n", cp->Spells.Illusion ? B_YN, cp->Spells.Recharge ? B_YN);
   printf("                      WizardsEye:     %c  WordOfDeath:    %c\n", cp->Spells.WizardsEye ? B_YN, cp->Spells.WordOfDeath ? B_YN);
   printf("                      Spooks:         %c  Resurrection:   %c\n", cp->Spells.Spooks ? B_YN, cp->Spells.Resurrection ? B_YN);
   printf("                      Death:          %c  DeathWish:      %c\n", cp->Spells.Death ? B_YN, cp->Spells.DeathWish ? B_YN);
   printf("                      LocateObject:   %c\n", cp->Spells.LocateObject ? B_YN);
   printf("                      MindFlay:       %c\n", cp->Spells.MindFlay ? B_YN);
   printf("                      FindPerson:     %c\n", cp->Spells.FindPerson ? B_YN);
}

void PrintSkills(struct Character *cp)
{
   printf("\nSkills:\n");
   printf("Weaponry          Physical          Personal          Academia\n");
   printf("   Wand&Dagger:%3.d   Swimming:   %3.d   Firearms:   %3.d   Artifacts:  %3.d\n",
      cp->Skills.Weaponry.Wand, cp->Skills.Physical.Swimming, cp->Skills.Personal.Firearms, cp->Skills.Academia.Artifacts);
   printf("   Sword:      %3.d   Climbing:   %3.d   Reflextion: %3.d   Mythology:  %3.d\n",
      cp->Skills.Weaponry.Sword, cp->Skills.Physical.Climbing, cp->Skills.Personal.Reflextion, cp->Skills.Academia.Mythology);
   printf("   Axe:        %3.d   Scouting:   %3.d   SnakeSpeed: %3.d   Mapping:    %3.d\n",
      cp->Skills.Weaponry.Axe, cp->Skills.Physical.Scouting, cp->Skills.Personal.SnakeSpeed, cp->Skills.Academia.Mapping);
   printf("   Mace&Flail: %3.d   Music:      %3.d   EagleEye:   %3.d   Scribe:     %3.d\n",
      cp->Skills.Weaponry.Mace, cp->Skills.Physical.Music, cp->Skills.Personal.EagleEye, cp->Skills.Academia.Scribe);
   printf("   Pole&Staff: %3.d   Oratory:    %3.d   PowerStrike:%3.d   Diplomacy:  %3.d\n",
      cp->Skills.Weaponry.Pole, cp->Skills.Physical.Oratory, cp->Skills.Personal.PowerStrike, cp->Skills.Academia.Diplomacy);
   printf("   Throwing:   %3.d   Legerdemain:%3.d   MindControl:%3.d   Alchemy:    %3.d\n",
      cp->Skills.Weaponry.Throwing, cp->Skills.Physical.Legerdemain, cp->Skills.Personal.MindControl, cp->Skills.Academia.Alchemy);
   printf("   Sling:      %3.d   Skulduggery:%3.d                     Theology:   %3.d\n",
      cp->Skills.Weaponry.Sling, cp->Skills.Physical.Skulduggery, cp->Skills.Academia.Theology);
   printf("   Bows:       %3.d   Ninjutsu:   %3.d                     Theosophy:  %3.d\n",
      cp->Skills.Weaponry.Bow, cp->Skills.Physical.Ninjutsu, cp->Skills.Academia.Theosophy);
   printf("   Shield:     %3.d                                       Thaumaturgy:%3.d\n",
      cp->Skills.Weaponry.Shield, cp->Skills.Academia.Thaumaturgy);
   printf("   Hands&Feet: %3.d                                       Kirijutsu:  %3.d\n",
      cp->Skills.Weaponry.HandToHand, cp->Skills.Academia.Kirijutsu);
}

void PrintCharacter(void *CharacterBuffer, short AttributesFlag,
     short ItemsFlag, short SkillsFlag, short SpellsFlag, short UnknownFlag)
{
   struct Character *cp = (struct Character *)CharacterBuffer;

   printf("\nCharacter Name: %s\n", cp->Name);
   printf("        Gender: %s\n", cp->Gender ? "Female" : "Male");
   printf("          Race: %s\n", RaceMap[cp->Race]);
   printf("    Profession: %s  Level %d\n", ProfessionMap[cp->Profession], cp->Level);
   printf("     Condition: %s (x%02X)  Life %d\n", ConditionMap[cp->ConditionCode], cp->ConditionCode, cp->Lives);
   printf("       ?Alive?: %u (x%02X)\n", cp->Alive, cp->Alive);

   if (AttributesFlag)
      PrintAttributes(cp);
   if (ItemsFlag)
      PrintItems(cp);
   if (SkillsFlag)
      PrintSkills(cp);
   if (SpellsFlag)
   {
      PrintSpellPoints(cp);
      PrintSpells(cp);
   }
   if (UnknownFlag)
      PrintUnknown(cp);

EXIT_LABEL:
   printf("\f");
   return;
}

int main(int argc, char **argv)
{
   short i,
         AttributesFlag = FALSE,
         ItemsFlag      = FALSE,
         SkillsFlag     = FALSE,
         SpellsFlag     = FALSE,
         UnknownFlag    = FALSE;
   char Args[513],
        Temp[33],
        FileName[133],
        CharacterName[8];
   char *p = NULL;

   normvideo();

   strcpy(Program.Name, "WIZPRINT");
   strcpy(Program.Title, "Wizardry VII: Crusaders of the Dark Savant");
   Program.Version = 1;
   Program.Revision = 1;
   Program.ScreenMode = FALSE;
   printf("[%s; %s\t%s %s]\n", Program.Name, Program.Title,
         GetDateString(), GetTimeString());
   strcpy(Args, strupr(argv[0]));
   for (i = 1; i <= argc; i++)
   {
      sprintf(Temp, " %s", strupr(argv[i]));
      strcat(Args, Temp);
   }

   if (argc < 3 || strstr(Args, "/?") || strstr(Args, "/H"))
   {
      if (argc < 3)
         printf("**** Error: Must specify WIZARDRY (.DBS) file name.\n");
      printf("     Usage: WIZPRINT wizardry.DBS <first/lead-character> [<Options>]\n");
      printf("     Where Options are:\n");
      printf("        /Basic        Only Display Basic Character Information\n");
      printf("        /Attributes   Display Character Attributes\n");
      printf("        /Items        Display Character's Items (inc. SwagBag)\n");
      printf("        /SKills       Display Character Skills\n");
      printf("        /SPells       Display Character Spells/Points\n");
      printf("        /Unknown      Display Hex Dump of Unknown Regions\n");
      printf("     Note: By default, all information is displayed.\n");
      goto EXIT_LABEL;
   }
   strcpy(FileName, strupr(argv[1]));
   strcpy(CharacterName, strupr(argv[2]));

   if (strstr(Args, "/A"))
      AttributesFlag = TRUE;
   if (strstr(Args, "/I"))
      ItemsFlag = TRUE;
   if (strstr(Args, "/SK"))
      SkillsFlag = TRUE;
   if (strstr(Args, "/SP"))
      SpellsFlag = TRUE;
   if (strstr(Args, "/U"))
      UnknownFlag = TRUE;
   if (!strstr(Args, "/"))
   {
      AttributesFlag = TRUE;
      ItemsFlag      = TRUE;
      SkillsFlag     = TRUE;
      SpellsFlag     = TRUE;
      UnknownFlag    = TRUE;
   }

   if (!ReadDBS(FileName, &Buffer, &FileSize))
      goto EXIT_LABEL;

   printf("   Searching for \"%s\"...\n", CharacterName);
   for (p = Buffer; p < Buffer+FileSize && strcmp(p, CharacterName); p++);
   if (!p)
   {
      printf("**** Error: Could not find \"%s\"...\n", CharacterName);
      goto EXIT_LABEL;
   }
   printf("   Found \"%s\" @ byte position %d (%d bytes)...\n\n", CharacterName, p-Buffer+1, sizeof(struct Character));

   for (i = 1; i <= 6; i++, p += sizeof(struct Character))
   {
      if (strstr(Args, "/DUMP") || strstr(Args, "/HEX"))
      {
         printf("\n%s\n", p); HexDump((unsigned char *)p, sizeof(struct Character));
      }
      else
         PrintCharacter((void *)p, AttributesFlag, ItemsFlag, SkillsFlag, SpellsFlag, UnknownFlag);
   }
   if (p < Buffer+FileSize)
   {
      printf("\n... more stuff at end of file (%d bytes)...\n\n", FileSize - (p-Buffer));
      HexDump((unsigned char *)p, FileSize - (p-Buffer));
   }

EXIT_LABEL:
   free(Buffer);
   normvideo();
   exit(0);
}
