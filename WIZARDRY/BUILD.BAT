@echo off
make -B -fWIZPRINT.MAK
make -B -fWIZREST.MAK
del *.OBJ
del *.CFG
