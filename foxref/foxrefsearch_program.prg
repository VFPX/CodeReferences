* Abstract...:
*	Search / Replace functionality for a .PRG, .H, .QPR, .MPR, etc
*
* Changes....:
*
#include "foxref.h"
#include "foxpro.h"

DEFINE CLASS RefSearchProgram AS RefSearch OF FoxRefSearch.prg
	Name = "RefSearchProgram"

	FUNCTION DoSearch()
		*** JRN 2010-03-12 : add reference to the file name for the PRG
		m.lSuccess = This.FindInText(JustFname(THIS.Filename), FINDTYPE_PRG, '', PRGNAME_LOC, SEARCHTYPE_EXPR, , "PRG", .T.)
		RETURN THIS.FindInCode(THIS.cFileText, FINDTYPE_CODE, '', '', SEARCHTYPE_NORMAL)
	ENDFUNC

	FUNCTION DoDefinitions()
		* THIS.FindDefinitions(FILETOSTR(THIS.Filename), '', '', SEARCHTYPE_NORMAL)
		THIS.FindDefinitions(THIS.cFileText, '', '', SEARCHTYPE_NORMAL)
	ENDFUNC
ENDDEFINE
