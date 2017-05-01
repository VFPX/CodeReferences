#include "foxref.h"
#include "foxpro.h"

* object we use to populate our collection of Matches
Define Class Match As Custom
	Name = "Match"

	MatchPos       = 0
	MatchLineNo    = 0
	MatchLen       = 0
Enddefine

* Abtract Match Engine Class
Define Class MatchEngine As Custom
	Name = "MatchEngine"

	Pattern        = ''
	PatternLen     = 0
	MatchCase      = .T.
	WholeWordsOnly = .F.

	oMatches = .Null.
	nMatchCnt = 0
	NoReplace = .F.

	Function InitEngine()
		Return .T.
	Endfunc

	Function Execute(cText As String) As Number
		Return 0
	Endfunc

	Function CreateMatch()
		Local oMatch

		m.oMatch = Createobject("Empty")
		AddProperty(m.oMatch, "MatchOffset", 0)
		AddProperty(m.oMatch, "MatchPos", 0)
		AddProperty(m.oMatch, "MatchLineNo", 0)
		AddProperty(m.oMatch, "MatchLen", 0)
		AddProperty(m.oMatch, "NoReplace", This.NoReplace)

		Return m.oMatch
	Endfunc

	* utility function for dermining if Whole Word
	Function IsAlphaNumeric(ch)
		Return Isalpha(ch) Or Between(ch, '0', '9') Or ch == '_'
	Endfunc

	Function SetPattern(cPattern)
		If Vartype(cPattern) <> 'C'
			Return .F.
		Endif

		This.PatternLen = Lenc(m.cPattern)
		This.Pattern = cPattern

		Return .T.
	Endfunc


Enddefine


* Standard Match Engine
Define Class MatchDefault As MatchEngine
	Name = "MatchDefault"

	Function Execute(cText As String)
		Local nPos
		Local nOccurrence
		Local i
		Local oMatch

		nOccurrence = 1
		i = 0


		This.oMatches = Createobject("Collection")
		If This.MatchCase
			nPos = At_c(This.Pattern, cText, nOccurrence)
		Else
			nPos = Atcc(This.Pattern, cText, nOccurrence)
		Endif

		* don't bother converting CR/LF to LF only unless there are matches
		If nPos > 0
			cText = Strtran(cText, Chr(13) + Chr(10), Chr(10))
			cText = Strtran(cText, Chr(13), Chr(10))

			If This.MatchCase
				nPos = At_c(This.Pattern, cText, nOccurrence)
			Else
				nPos = Atcc(This.Pattern, cText, nOccurrence)
			Endif
		Endif

		Do While nPos <> 0
			* if Whole Word Only is selected, then we must
			* make sure the word we found comprises a whole word
			*** JRN 2010-03-15 : Whole Words Only did not match word starting at first position in file / method
			*	IF (!THIS.WholeWordsOnly OR (m.nPos > 1 AND !THIS.IsAlphaNumeric(SUBSTRC(m.cText, m.nPos - 1, 1)) AND (m.nPos == LENC(m.cText) OR !THIS.IsAlphaNumeric(SUBSTRC(m.cText, m.nPos + THIS.PatternLen, 1)))))
			If (!This.WholeWordsOnly Or 														;
					(m.nPos = 1 or !This.IsAlphaNumeric(Substrc(m.cText, m.nPos - 1, 1)))		;
					And (m.nPos == Lenc(m.cText) Or !This.IsAlphaNumeric(Substrc(m.cText, m.nPos + This.PatternLen, 1)) ;
					))
				i = i + 1
				oMatch = This.CreateMatch()
				oMatch.MatchLen = This.PatternLen

				oMatch.MatchPos = nPos - Ratc(Chr(10), Leftc(cText, nPos))
				oMatch.MatchLineNo    = Occurs(Chr(10), Leftc(cText, nPos)) + 1

				This.oMatches.Add(oMatch)
			Endif

			nOccurrence = nOccurrence + 1
			If This.MatchCase
				nPos = At_c(This.Pattern, cText, nOccurrence)
			Else
				nPos = Atcc(This.Pattern, cText, nOccurrence)
			Endif
		Enddo

		This.nMatchCnt = i

		Return This.nMatchCnt
	Endfunc


Enddefine


* Regular Expression search engine (relies on VBScript)
Define Class MatchWildcard As MatchEngine
	Name = "MatchWildcard"

	oRegExpr = .Null.

	Function InitEngine()
		Local lSuccess

		m.lSuccess = .T.
		Try
			This.oRegExpr = Createobject("VBScript.RegExp")
			This.oRegExpr.Global = .T.
		Catch
			m.lSuccess = .F.
			This.oRegExpr = .Null.
			Messagebox(WSH_REQUIRED_LOC, MB_ICONEXCLAMATION, APPNAME_LOC)
		Endtry

		Return m.lSuccess
	Endfunc

	Function SetPattern(cPattern)
		Local lSuccess
		Local oException

		If Vartype(cPattern) <> 'C'
			Return .F.
		Endif
		If Empty(cPattern)
			Return .T.
		Endif

		lSuccess = .T.

		If This.WholeWordsOnly
			cPattern = "\b" + cPattern + "\b"
		Endif

		Try
			This.oRegExpr.Pattern = cPattern
		Catch To oException
			m.lSuccess = .F.
		Endtry

		* now check the pattern
		Try
			This.oRegExpr.Test("abcdefg")
		Catch To oException
			m.lSuccess = .F.
		Endtry

		If lSuccess
			This.PatternLen = Lenc(m.cPattern)
			This.Pattern = cPattern
		EndIf
		
		This.NoReplace = Len(cPattern) # Len(Chrtran( cPattern, RegExpSpecialChars, '')) 

		Return lSuccess
	Endfunc


	Function Execute(cText As String)
		Local oMatches
		Local oMatch
		Local oRegExprMatch
		Local nPos
		Local lSuccess
		Local oException
		Local nOccurs
		Local cFindChar
		Local cTextRegExpr

		This.oMatches = Createobject("Collection")


		This.nMatchCnt = 0

		lSuccess = .T.
		Try
			This.oRegExpr.IgnoreCase = !This.MatchCase
			This.oRegExpr.MultiLine = .T.
			****************************************************************
			*** JRN 2010-03-12 : handle continuation lines
			cTextRegExpr = Strtran (cText, ';' + Chr(13) + Chr(10), Space(3))
			****************************************************************
			oMatches = This.oRegExpr.Execute(cTextRegExpr)
		Catch To oException
			lSuccess = .F.
		Endtry


		If m.lSuccess
			This.nMatchCnt = oMatches.Count
			For Each oRegExprMatch In oMatches
				nPos = oRegExprMatch.FirstIndex + 1

				oMatch = This.CreateMatch()
				oMatch.MatchLen       = Lenc(oRegExprMatch.Value)


				cFindChar = Chr(10)
				nOccurs = Occurs(cFindChar, Leftc(cText, nPos))
				If nOccurs == 0
					cFindChar = Chr(13)
					nOccurs = Occurs(cFindChar, Leftc(cText, nPos))
				Endif
				oMatch.MatchLineNo    = nOccurs + 1
				oMatch.MatchPos       = nPos - Ratc(cFindChar, Leftc(cText, nPos))

				This.oMatches.Add(oMatch)
			Endfor
		Else
			Return -1  && error
		Endif

		Return This.nMatchCnt
	Endfunc

Enddefine

