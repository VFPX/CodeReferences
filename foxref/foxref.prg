* Abstract...:
*	Primary class for References application.
*
* Changes....:
*
#Ifdef TESTING
	* Below is a sample of using the FoxRef
	* class programmatically
	Public o As FoxRef Of FoxRef.prg

	cProject = "foxref"

	cPattern   = "oApp"
	cFileDir   = "c:\code\visgift"
	* cSkeleton  = "*.scx"
	cFileTypes = "*.prg *.scx"

	o = Newobject("FoxRef", "foxref.prg")

	If o.SetProject(cProject)
		o.Quiet = .F.
		o.FileTypes = cFileTypes
		o.Search("for")

		o.OverwritePrior = .F.
		o.WildCards = .T.
		o.Search("whil?")
	Else
		? "Unable to open project"
	Endif

	Return
#Endif

#include "foxpro.h"
#include "foxref.h"

Define Class FoxRef As Session
	Protected lIgnoreErrors As Boolean
	Protected lRefreshMode
	Protected cProgressForm
	Protected lCancel
	Protected tTimeStamp
	Protected lIgnoreErrors
	Protected Array aFileTypes[1]
	Protected nFileTypeCnt


	oSearchEngine   = .Null.

	Comments        = COMMENTS_INCLUDE
	MatchCase       = .T.
	WholeWordsOnly  = .F.

	SubFolders      = .F.
	WildCards       = .F.
	Quiet           = .F.  && quiet mode -- don't display search progress
	ShowProgress    = .T.  && show a progress form

	FileTypes       = ''
	ReportFile      = REPORT_FILE

	* MRU array lists
	Dimension aLookForMRU[10]
	Dimension aReplaceMRU[10]
	Dimension aFolderMRU[10]
	Dimension aFileTypesMRU[10]
	aLookForMRU     = ''
	aReplaceMRU     = ''
	aFolderMRU      = ''
	aFileTypesMRU   = ''

	Pattern         = ''
	OverwritePrior  = .T.

	RefTable        = ''
	ProjectFile     = ''
	FileDirectory   = ''

	cSetID          = ''
	lRefreshMode    = .F.

	lIgnoreErrors   = .F.

	oProgressForm   = .Null.
	lCancel         = .F.
	tTimeStamp      = .Null.

	nFileTypeCnt    = 0

	cTalk           = ''
	nLangOpt        = 0
	cMessage        = ''
	cEscapeState    = ''
	cSYS3054        = ''
	cSaveUDFParms   = ''
	cSaveLib        = ''
	cExclusive      = ''

	Procedure Init()
		This.cTalk = Set("TALK")
		Set Talk Off

		Set Deleted On

		This.cExclusive = Set("EXCLUSIVE")
		Set Exclusive Off

		This.cMessage = Set("MESSAGE",1)

		This.nLangOpt = _vfp.LanguageOptions
		_vfp.LanguageOptions=0

		This.cEscapeState = Set("ESCAPE")
		Set Escape Off

		This.cSYS3054 = Sys(3054)
		Sys(3054,0)

		This.cSaveLib      = Set("LIBRARY")

		This.cSaveUDFParms = Set("UDFPARMS")
		Set Udfparms To Value

		Set Exact Off


		This.RestorePrefs()

		* Create file type engine objects
		This.AddFileType('', FILETYPE_CLASS_DEFAULT, FILETYPE_LIBRARY_DEFAULT)  && default search engine
		This.AddFileType("PRG", FILETYPE_CLASS_PRG, FILETYPE_LIBRARY_PRG)  && program
		This.AddFileType("H",   FILETYPE_CLASS_H,   FILETYPE_LIBRARY_H)    && header
		This.AddFileType("SCX", FILETYPE_CLASS_SCX, FILETYPE_LIBRARY_SCX)  && form
		This.AddFileType("VCX", FILETYPE_CLASS_VCX, FILETYPE_LIBRARY_VCX)  && class library
		This.AddFileType("DBF", FILETYPE_CLASS_DBF, FILETYPE_LIBRARY_DBF)  && table
		This.AddFileType("DBC", FILETYPE_CLASS_DBC, FILETYPE_LIBRARY_DBC)  && database container
		This.AddFileType("FRX", FILETYPE_CLASS_FRX, FILETYPE_LIBRARY_FRX)  && report
		This.AddFileType("LBX", FILETYPE_CLASS_LBX, FILETYPE_LIBRARY_LBX)  && label
		This.AddFileType("MNX", FILETYPE_CLASS_MNX, FILETYPE_LIBRARY_MNX)  && menu
		This.AddFileType("SPR", FILETYPE_CLASS_SPR, FILETYPE_LIBRARY_SPR)  && screen
		This.AddFileType("QPR", FILETYPE_CLASS_QPR, FILETYPE_LIBRARY_QPR)  && query
		This.AddFileType("MPR", FILETYPE_CLASS_MPR, FILETYPE_LIBRARY_MPR)  && menu
	Endfunc


	Procedure Destroy()
		This.CloseProgress()

		If This.cEscapeState = "ON"
			Set Escape On
		Endif
		If This.cTalk = "ON"
			Set Talk On
		Endif
		If This.cExclusive = "ON"
			Set Exclusive On
		Endif
		Sys(3054,Int(Val(This.cSYS3054)))

		_vfp.LanguageOptions = This.nLangOpt

		If This.cSaveUDFParms = "REFERENCE"
			Set Udfparms To Reference
		Endif
	Endfunc

	*!*		PROCEDURE Error(nError, cMethod, nLine)
	*!*			IF THIS.lIgnoreErrors
	*!*				RETURN
	*!*			ENDIF
	*!*			DODEFAULT(nError, cMethod, nLine)
	*!*		ENDPROC


	* To-Do: change this to a collection
	* Add a filetype search -- This can be a collection in VFP8!!!
	Function AddFileType(cFileType, cClassName, cClassLibrary)
		Local nIndex
		Local lSuccess

		cFileType = Alltrim(Upper(cFileType))
		If Left(cFileType, 1) == '.'
			cFileType = Substr(cFileType, 2)
		Endif

		lSuccess = .F.

		nIndex = 0
		If Vartype(cFileType) <> 'C' Or Empty(cFileType)
			cFileType = ''
			nIndex = 1 && default search engine always the first row in the array
		Else
			If This.nFileTypeCnt > 0
				nIndex = Ascan(This.aFileTypes, cFileType, -1, -1, 1, 14)  && 14 = return row number, Exact On
			Endif
		Endif
		If Pcount() == 1
			If nIndex > 0
				* remove filetype
				=Adel(This.aFileTypes, nIndex, 2)
				This.nFileTypeCnt = This.nFileTypeCnt - 1

				lSuccess = .T.
			Endif
		Else
			If Vartype(cClassLibrary) <> 'C' Or Empty(cClassLibrary)
				cClassLibrary = FILETYPE_LIBRARY
			Endif
			oEngine = Newobject(cClassName, cClassLibrary)
			If Vartype(oEngine) == 'O'
				If nIndex == 0 Or This.nFileTypeCnt == 0
					This.nFileTypeCnt = This.nFileTypeCnt + 1
					nIndex = This.nFileTypeCnt
					Dimension This.aFileTypes[THIS.nFileTypeCnt, 2]
				Endif
				This.aFileTypes[nIndex, FILETYPE_EXTENSION] = cFileType
				This.aFileTypes[nIndex, FILETYPE_ENGINE]    = oEngine

				lSuccess = .T.
			Endif
		Endif

		Return lSuccess
	Endfunc

	Procedure SearchInit()
		Local i

		If This.WildCards
			This.oSearchEngine = Newobject("Wildcard", "foxmatch.prg")
		Else
			This.oSearchEngine = Newobject("MatchDefault", "foxmatch.prg")
		Endif
		This.oSearchEngine.MatchCase      = This.MatchCase
		This.oSearchEngine.WholeWordsOnly = This.WholeWordsOnly

		For i = 1 To This.nFileTypeCnt
			With This.aFileTypes[i, FILETYPE_ENGINE]
				.SetID          = This.cSetID
				.oSearchEngine  = This.oSearchEngine
				.Pattern        = This.Pattern
				.Comments       = This.Comments
			Endwith
		Endfor
	Endproc



	* Do a replacement on designated file
	Function ReplaceFile(cUniqueID, cReplaceText)
		Local nSelect
		Local nFileTypeIndex
		Local cFileType
		Local oFoxRefRecord
		Local lSuccess

		lSuccess = .F.
		If Used("FoxRefCursor") And Vartype(cUniqueID) == 'C' And !Empty(cUniqueID) And (FoxRefCursor.UniqueID == cUniqueID Or Seek(cUniqueID, "FoxRefCursor", "UniqueID"))
			nSelect = Select()

			cFilename = Addbs(Rtrim(FoxRefCursor.Folder)) + Rtrim(FoxRefCursor.Filename)
			If File(cFilename)
				cFileType = Upper(Justext(cFilename))

				nFileTypeIndex = Ascan(This.aFileTypes, cFileType, -1, -1, 1, 14)  && 14 = return row number, Exact On
				If nFileTypeIndex == 0
					nFileTypeIndex = 1  && this is the default search/replace engine to use
				Endif

				Select FoxRefCursor
				Scatter Memo Name oFoxRefRecord
				With This.aFileTypes[nFileTypeIndex, FILETYPE_ENGINE]
					lSuccess = .ReplaceWith(cReplaceText, oFoxRefRecord)
				Endwith
			Endif

			Select (nSelect)
		Endif

		Return lSuccess
	Endfunc




	Procedure FileTypes_Assign(cFileTypes)
		This.FileTypes = Chrtran(cFileTypes, ',;', '  ')
	Endfunc

	Function CreateRefTable(cRefTable)
		Local lSuccess
		Local cSafety

		lSuccess = .T.

		This.RefTable = ''

		cSafety = Set("SAFETY")
		Set Safety Off
		If Used(Juststem(cRefTable))
			Use In (cRefTable)
		Endif

		Create Table (cRefTable) Free ( ;
			UniqueID C(10), ;
			SetID C(10), ;
			RefID C(10), ;
			RefType C(1), ;
			DefType C(1), ;
			Folder C(240), ;
			Filename C(100), ;
			Symbol C(254), ;
			ClassName C(254), ;
			ProcName C(254), ;
			ProcLineNo i, ;
			LineNo i, ;
			ColPos i, ;
			MatchLen i, ;
			RefCode C(254), ;
			RecordID C(10), ;
			UpdField C(10), ;
			Checked L, ;
			TimeStamp T Null, ;
			oTimeStamp N(10) Null, ;
			Inactive L ;
			)
		Index On RefType Tag RefType
		Index On SetID Tag SetID
		Index On RefID Tag RefID
		Index On UniqueID Tag UniqueID
		Index On Filename Tag Filename
		Index On Checked Tag Checked


		* add the record that holds our results window search position & other options
		Insert Into (cRefTable) ( ;
			UniqueID, ;
			SetID, ;
			RefType, ;
			Folder, ;
			Filename, ;
			Symbol, ;
			ClassName, ;
			ProcName, ;
			ProcLineNo, ;
			LineNo, ;
			ColPos, ;
			MatchLen, ;
			RefCode, ;
			RecordID, ;
			UpdField, ;
			Timestamp, ;
			Inactive ;
			) Values ( ;
			SYS(2015), ;
			'', ;
			REFTYPE_INIT, ;
			THIS.ProjectFile, ;
			'', ;
			'', ;
			'', ;
			'', ;
			0, ;
			0, ;
			0, ;
			0, ;
			'', ;
			'', ;
			'', ;
			DATETIME(), ;
			.F. ;
			)
		*		  IIF(THIS.ProjectName == PROJECT_GLOBAL, cProjectOrDir, ''), ;
		*		  IIF(!(THIS.ProjectName == PROJECT_GLOBAL), cProjectOrDir, ''), ;
		*		  IIF(cScope == SCOPE_FOLDER, cProjectOrDir, ''), ;
		*		  IIF(cScope <> SCOPE_FOLDER, cProjectOrDir, ''), ;

		This.RefTable = cRefTable

		Use In (Juststem(cRefTable))

		Set Safety &cSafety

		Return lSuccess
	Endfunc

	* Open a FoxRef table
	* Return TRUE if table exists and it's in the correct format
	* [lCreate] = True to create table if it doesn't exist
	Function OpenRefTable(cRefTable)
		Local lSuccess

		If Used("FoxRefCursor")
			Use In FoxRefCursor
		Endif
		This.RefTable = ''

		lSuccess = .T.

		If !File(Forceext(cRefTable, "DBF"))
			lSuccess = This.CreateRefTable(cRefTable)
		Endif

		If lSuccess
			Use (cRefTable) Alias FoxRefCursor In 0 Shared Again
			If Type("FoxRefCursor.RefType") == 'C'
				This.RefTable = cRefTable
			Else
				lSuccess = .F.
				Messagebox(BADTABLE_LOC + Chr(10) + Chr(10) + Forceext(cRefTable, "DBF"), MB_ICONSTOP, APPNAME_LOC)
			Endif
		Endif

		Return lSuccess
	Endfunc

	* Abstract:
	*   Set a specific project to display result sets for.
	*	Pass an empty string or "global" to display result sets
	*	that are not associated with a project.
	*
	* Parameters:
	*   [cProject]
	Function SetProject(cProjectFile, lOverwrite)
		Local lInUse
		Local lSuccess
		Local cRefTable
		Local i
		Local lFoundProject

		lSuccess = .F.

		If Vartype(cProjectFile) <> 'C'
			cProjectFile = This.ProjectFile
		Endif
		If Empty(cProjectFile)
			* use the active project if a project name is not passsed
			If Application.Projects.Count > 0
				cProjectFile = Application.ActiveProject.Name
			Else
				cProjectFile = PROJECT_GLOBAL
			Endif

			lSuccess = .T.
		Else
			* make sure Project specified is open
			If cProjectFile == PROJECT_GLOBAL
				lSuccess = .T.
			Else
				cProjectFile = Upper(Forceext(Fullpath(cProjectFile), "PJX"))

				For i = 1 To Application.Projects.Count
					If Upper(Application.Projects(i).Name) == cProjectFile
						cProjectFile = Application.Projects(i).Name
						lSuccess = .T.
						Exit
					Endif
				Endfor
				If !lSuccess
					If File(cProjectFile)
						* open the project
						Modify Project (cProjectFile) Nowait

						* search again to find where in the Projects collection it is
						For i = 1 To Application.Projects.Count
							If Upper(Application.Projects(i).Name) == cProjectFile
								cProjectFile = Application.Projects(i).Name
								lSuccess = .T.
								Exit
							Endif
						Endfor
					Endif
				Endif
			Endif
		Endif


		If lSuccess
			If Empty(cProjectFile) Or cProjectFile == PROJECT_GLOBAL
				This.ProjectFile = PROJECT_GLOBAL
				cRefTable        = Addbs(Home()) + GLOBAL_TABLE + RESULT_EXT
			Else
				This.ProjectFile = Upper(cProjectFile)
				cRefTable        = Addbs(Justpath(cProjectFile)) + Juststem(cProjectFile) + RESULT_EXT
			Endif

			If lOverwrite
				lSuccess = This.CreateRefTable(cRefTable)
			Else
				lSuccess = This.OpenRefTable(cRefTable)
			Endif
		Endif

		Return lSuccess
	Endfunc


	Function Search(cPattern)
		If Vartype(cPattern) <> 'C' Or Empty(cPattern)
			Do Form FoxRefFind With This
		Else
			If This.SetProject(, This.OverwritePrior)
				If This.ProjectFile == PROJECT_GLOBAL Or Empty(This.ProjectFile)
					This.FolderSearch(cPattern)
				Else
					This.ProjectSearch(cPattern)
				Endif
			Endif
		Endif
	Endfunc


	* Determine if the Reference table we want to open is
	* actually one of ours.  If we're overwriting or a reference
	* table doesn't exist for this project, then create a new
	* Reference Table.
	*
	* Once we have a reference table, then we add a new record
	* that represents the search criteria for this particular
	* search.
	Function UpdateRefTable(cScope, cPattern, cProjectOrDir)
		Local nSelect
		Local cSafety
		Local cSearchOptions
		Local cRefTable

		nSelect = Select()

		If Vartype(cRefTable) <> 'C' Or Empty(cRefTable)
			cRefTable = This.RefTable
		Endif

		If Empty(cRefTable)
			Return .F.
		Endif

		If Used("FoxRefCursor")
			Use In FoxRefCursor
		Endif

		*!*			IF FILE(FORCEEXT(cRefTable, "DBF"))
		*!*				IF !THIS.OpenRefTable(cRefTable)
		*!*					RETURN .F.
		*!*				ENDIF
		*!*			ENDIF

		If !This.OpenRefTable(cRefTable)
			Return .F.
		Endif

		This.tTimeStamp = Datetime()


		* build a string representing the search options that
		* we can store to the FoxRef cursor
		cSearchOptions = Iif(This.Comments == COMMENTS_EXCLUDE, 'X', '') + ;
			IIF(This.Comments == COMMENTS_ONLY, 'C', '') + ;
			IIF(This.MatchCase, 'M', '') + ;
			IIF(This.WholeWordsOnly, 'W', '') + ;
			IIF(This.SubFolders, 'S', '') + ;
			IIF(This.WildCards, 'Z', '') + ;
			';' + This.FileTypes


		* if we've already searched for this same exact symbol
		* with the same exact criteria in the same exact project/folder,
		* then simply update what we have
		Select FoxRefCursor
		Locate For RefType == REFTYPE_SEARCH And Folder == Padr(cProjectOrDir, Len(FoxRefCursor.Folder)) And Symbol == Padr(cPattern + PATTERN_EOL, Len(FoxRefCursor.Symbol)) And RefCode == Padr(cSearchOptions, Len(FoxRefCursor.RefCode))
		This.lRefreshMode = Found()
		If This.lRefreshMode
			This.tTimeStamp = FoxRefCursor.Timestamp

			This.cSetID = FoxRefCursor.SetID
			Replace All Inactive With .T. ;
				FOR SetID == This.cSetID And (RefType == REFTYPE_RESULT Or RefType == REFTYPE_DEFINITION) ;
				IN FoxRefCursor
		Else
			This.cSetID = Sys(2015)

			* add the record that specifies the search criteria, etc
			Insert Into FoxRefCursor ( ;
				UniqueID, ;
				SetID, ;
				RefType, ;
				Folder, ;
				Filename, ;
				Symbol, ;
				ClassName, ;
				ProcName, ;
				ProcLineNo, ;
				LineNo, ;
				ColPos, ;
				MatchLen, ;
				RefCode, ;
				RecordID, ;
				UpdField, ;
				Timestamp, ;
				Inactive ;
				) Values ( ;
				SYS(2015), ;
				THIS.cSetID, ;
				REFTYPE_SEARCH, ;
				cProjectOrDir, ;
				'', ;
				cPattern + PATTERN_EOL, ;
				'', ;
				'', ;
				0, ;
				0, ;
				0, ;
				0, ;
				cSearchOptions, ;
				'', ;
				'', ;
				THIS.tTimeStamp, ;
				.F. ;
				)
		Endif

		Select (nSelect)

		Return .T.
	Endfunc


	* -- Search a Folder
	Function FolderSearch(cPattern, cFileDir)
		Local nFileCnt
		Local i, j
		Local cFileDir
		Local nFileTypesCnt
		Local cFileTypes
		Local lAutoYield
		Local Array aFileList[1]
		Local Array aFileTypes[1]

		If Vartype(cPattern) <> 'C'
			cPattern = This.Pattern
		Endif

		If Vartype(cFileDir) <> 'C' Or Empty(cFileDir)
			cFileDir = Addbs(This.FileDirectory)
		Else
			cFileDir = Addbs(cFileDir)
		Endif

		If !Directory(cFileDir)
			Return .F.
		Endif

		This.SetProject(PROJECT_GLOBAL, This.OverwritePrior)

		If !This.UpdateRefTable(SCOPE_FOLDER, cPattern, cFileDir)
			Return .F.
		Endif

		This.SearchInit()


		cFileTypes = Chrtran(This.FileTypes, ',;', '  ')
		nFileTypesCnt = Alines(aFileTypes, cFileTypes, .T., ' ')

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.


		This.ProcessFolder(cFileDir, cPattern, @aFileTypes, nFileTypesCnt)

		This.CloseProgress()

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU(cPattern)
		This.UpdateFolderMRU(cFileDir)
		This.UpdateFileTypesMRU(cFileTypes)

		Return .T.
	Endfunc

	* used in conjuction with FolderSearch() for
	* when we're searching subfolders
	Function ProcessFolder(cFileDir, cPattern, aFileTypes, nFileTypesCnt)
		Local nFileCnt
		Local nFolderCnt
		Local cFilename
		Local i, j
		Local Array aFileList[1]
		Local Array aFolderList[1]

		cFileDir = Addbs(cFileDir)
		For i = 1 To nFileTypesCnt
			If This.lCancel
				Exit
			Endif

			nFileCnt = Adir(aFileList, cFileDir + aFileTypes[i], '', 1)
			For j = 1 To nFileCnt
				If This.lCancel
					Exit
				Endif

				cFilename = aFileList[j, 1]

				This.FileSearch(cFileDir + cFilename, cPattern)
			Endfor
		Endfor

		* Process any sub-directories
		If !This.lCancel
			If This.SubFolders
				nFolderCnt = Adir(aFolderList, cFileDir + "*.*", 'D', 1)
				For i = 1 To nFolderCnt
					If !aFolderList[i, 1] == '.' And !aFolderList[i, 1] == '..' And 'D'$aFolderList[i, 5] And Directory(cFileDir + aFolderList[i, 1])
						This.ProcessFolder(cFileDir + aFolderList[i, 1], cPattern, @aFileTypes, nFileTypesCnt)
					Endif
					If This.lCancel
						Exit
					Endif
				Endfor
			Endif
		Endif
	Endfunc


	* -- Search files in a Project
	Function ProjectSearch(cPattern, cProjectFile)
		Local nFileIndex
		Local nProjectIndex
		Local oProjectRef
		Local oFileRef
		Local cFileTypes
		Local nFileTypesCnt
		Local lAutoYield
		Local lSuccess
		Local Array aFileTypes[1]
		Local Array aFileList[1]

		If Vartype(cPattern) <> 'C'
			cPattern = This.Pattern
		Endif

		*!*			IF VARTYPE(cProjectFile) == 'C' AND !EMPTY(cProjectFile)
		*!*				cProjectFile = UPPER(FORCEEXT(FULLPATH(cProjectFile), "PJX"))
		*!*			ELSE
		*!*				cProjectFile = ''
		*!*			ENDIF

		*!*			IF !THIS.SetProject(cProjectFile)
		*!*				RETURN .F.
		*!*			ENDIF
		If Vartype(cProjectFile) <> 'C' Or Empty(cProjectFile)
			cProjectFile = This.ProjectFile
		Endif
		If !This.SetProject(cProjectFile, This.OverwritePrior)
			Return .F.
		Endif


		If !This.UpdateRefTable(SCOPE_PROJECT, cPattern, This.ProjectFile)
			Return .F.
		Endif

		This.SearchInit()

		cFileTypes = This.FileTypes

		lAutoYield = _vfp.AutoYield
		_vfp.AutoYield = .T.

		For Each oProjectRef In Application.Projects
			If Upper(oProjectRef.Name) == This.ProjectFile
				* process each file in the project
				For Each oFileRef In oProjectRef.Files
					If This.WildCardMatch(cFileTypes, Justfname(oFileRef.Name))
						This.FileSearch(oFileRef.Name, cPattern)
					Endif
					If This.lCancel
						Exit
					Endif
				Endfor

				Exit
			Endif
		Endfor
		This.CloseProgress()

		_vfp.AutoYield = lAutoYield

		This.UpdateLookForMRU(cPattern)
		This.UpdateFileTypesMRU(cFileTypes)

		Return .T.
	Endfunc


	* Search a file
	Function FileSearch(cFilename, cPattern)
		Local cFileType
		Local nSelect
		Local cFileFind
		Local cFolderFind
		Local lSearch
		Local nSelect
		Local nFileTypeIndex
		Local Array aFileList[1]

		This.UpdateProgress(cFilename)
		If This.lCancel
			Return
		Endif

		If Vartype(cPattern) <> 'C'
			cPattern = This.Pattern
		Endif

		nSelect = Select()

		lSearch = .T.

		* determine which search engine to use based upon the filetype
		cFileType = Upper(Justext(cFilename))
		nFileTypeIndex = Ascan(This.aFileTypes, cFileType, -1, -1, 1, 14)  && 14 = return row number, Exact On
		If nFileTypeIndex == 0
			nFileTypeIndex = 1  && this is the default search engine to use
		Endif

		With This.aFileTypes[nFileTypeIndex, FILETYPE_ENGINE]
			.Filename       = cFilename
			* .FileTimeStamp  = THIS.FileTimeStamp


			* don't try to check the timestamp on a Table (DBF)
			* because it isn't always updated when we modify
			* certain things (such as values that are really stored in the
			* database container)
			If This.lRefreshMode And !.Norefresh
				* if we're refreshing results, then first see if
				* we've already searched this file and the timestamp
				* hasn't changed at all

				cFileFind   = Padr(Justfname(cFilename), 100)
				cFolderFind = Padr(Justpath(cFilename), 240)

				Select FoxRefCursor
				Locate For SetID == This.cSetID And (RefType == REFTYPE_RESULT Or RefType == REFTYPE_DEFINITION Or RefType == REFTYPE_NOMATCH) And Filename == cFileFind And Folder == cFolderFind
				If Found()
					If FoxRefCursor.Timestamp == .FileTimeStamp
						Update FoxRefCursor Set Inactive = .F. Where SetID == This.cSetID And (RefType == REFTYPE_RESULT Or RefType == REFTYPE_DEFINITION Or RefType == REFTYPE_NOMATCH) And Filename == cFileFind And Folder == cFolderFind
						lSearch = .F.
					Endif
				Endif
			Endif


			If lSearch
				.SearchFor(cPattern)
			Endif
		Endwith


		Select (nSelect)

	Endfunc



	* refresh results for all Sets in the Ref table or a single set
	Function RefreshResults(cSetID)
		Local nSelect
		Local lInUse
		Local i
		Local nCnt
		Local Array aRefList[1]

		nSelect = Select()

		If Vartype(cSetID) == 'C' And !Empty(cSetID)
			This.RefreshResultSet(cSetID)
		Else
			If File(Forceext(This.RefTable, "dbf"))
				nSelect = Select()

				lInUse = Used("FoxRefCursor")
				If !lInUse
					Use (This.RefTable) Alias FoxRefCursor In 0 Shared Again
				Endif

				Select SetID ;
					FROM FoxRefCursor ;
					WHERE RefType == REFTYPE_SEARCH And !Inactive ;
					INTO Array aRefList
				nCnt = _Tally
				*!*					IF !lInUse AND USED("FoxRefCursor")
				*!*						USE IN FoxRefCursor
				*!*					ENDIF

				For i = 1 To nCnt
					This.RefreshResultSet(aRefList[i])
				Endfor


				Select (nSelect)
			Endif
			This.cSetID = ''
		Endif
	Endfunc

	* refresh an existing search set
	Function RefreshResultSet(cSetID)
		Local lInUse
		Local nSelect
		Local lSuccess
		Local cScope
		Local cFolder
		Local cProject
		Local cPattern
		Local cSearchOptions

		lSuccess = .F.
		If File(Forceext(This.RefTable, "dbf"))
			nSelect = Select()

			lInUse = Used("FoxRefCursor")
			If !lInUse
				Use (This.RefTable) Alias FoxRefCursor In 0 Shared Again
			Endif

			Select FoxRefCursor
			Locate For RefType == REFTYPE_SEARCH And SetID == cSetID
			lSuccess = Found()
			If lSuccess
				cSearchOptions = Left(FoxRefCursor.RefCode, At(';', FoxRefCursor.RefCode) - 1)
				If 'X'$cSearchOptions
					This.Comments = COMMENTS_EXCLUDE
				Endif
				If 'C'$cSearchOptions
					This.Comments = COMMENTS_ONLY
				Endif
				This.MatchCase      = 'M' $ cSearchOptions
				This.WholeWordsOnly = 'W' $ cSearchOptions
				This.SubFolders     = 'S' $ cSearchOptions
				This.WildCards      = 'Z' $ cSearchOptions

				This.OverwritePrior = .F.

				This.FileTypes = Alltrim(Substr(FoxRefCursor.RefCode, At(';', FoxRefCursor.RefCode) + 1))
				cFolder  = Rtrim(FoxRefCursor.Folder)
				cProject = ''

				If Upper(Justext(cFolder)) == "PJX"
					cScope = SCOPE_PROJECT
					cProject = cFolder
				Else
					cScope = SCOPE_FOLDER
				Endif

				If PATTERN_EOL $ FoxRefCursor.Symbol
					cPattern = Left(FoxRefCursor.Symbol, At(PATTERN_EOL, FoxRefCursor.Symbol) - 1)
				Else
					cPattern = Rtrim(FoxRefCursor.Symbol)
				Endif
			Endif

			If !lInUse And Used("FoxRefCursor")
				Use In FoxRefCursor
			Endif


			If lSuccess
				Do Case
					Case cScope == SCOPE_FOLDER
						This.FolderSearch(cPattern, cFolder)
					Case cScope == SCOPE_PROJECT
						This.ProjectSearch(cPattern, cProject)
					Otherwise
						lSuccess = .F.
				Endcase
			Endif

			Select (nSelect)
		Endif

		Return lSuccess
	Endfunc

	Function SetChecked(cUniqueID, lChecked)
		If Pcount() < 2
			lChecked = .T.
		Endif
		If Used("FoxRefCursor") And Seek(cUniqueID, "FoxRefCursor", "UniqueID")
			Replace Checked With lChecked In FoxRefCursor
		Endif
	Endfunc


	* -- Show the Results form
	Function ShowResults()
		Local i

		* first see if there is an open Results window for this
		* project and display that if there is
		For i = 1 To _Screen.FormCount
			If Pemstatus(_Screen.Forms(i), "oFoxRef", 5) And ;
					UPPER(_Screen.Forms(i).Name) == "FRMFOXREFRESULTS" And ;
					VARTYPE(_Screen.Forms(i).oFoxRef) == 'O' And Upper(_Screen.Forms(i).oFoxRef.ProjectFile) == This.ProjectFile
				* _SCREEN.Forms(i).RefreshResults(THIS.cSetID)
				_Screen.Forms(i).SetRefTable(This.cSetID)
				Return
			Endif
		Endfor

		Do Form FoxRefResults With This
	Endfunc

	* goto a specific reference
	Function GotoReference(cUniqueID)
		Local nSelect
		Local cFilename
		Local cFileType
		Local cClassName
		Local cProcName

		If Vartype(cUniqueID) <> 'C' Or Empty(cUniqueID)
			Return .F.
		Endif

		If Used("FoxRefCursor") And Seek(cUniqueID, "FoxRefCursor", "UniqueID")
			nSelect = Select()

			cFilename  = Addbs(Rtrim(FoxRefCursor.Folder)) + Rtrim(FoxRefCursor.Filename)
			cClassName = Rtrim(FoxRefCursor.ClassName)
			cProcName  = Rtrim(FoxRefCursor.ProcName)
			cFileType  = Upper(Justext(cFilename))

			Do Case
				Case cFileType == "SCX"
					Editsource(cFilename, Max(FoxRefCursor.ProcLineNo, 1), cClassName, cProcName)

				Case cFileType == "VCX"
					Editsource(cFilename, Max(FoxRefCursor.ProcLineNo, 1), cClassName, cProcName)

				Case cFileType == "DBF"
					* do a TRY/CATCH here
					If Used(Juststem(cFilename))
						Select (Juststem(cFilename))
					Else
						Select 0
						Use (cFilename) Exclusive
					Endif
					Modify Structure

				Otherwise
					Editsource(cFilename, FoxRefCursor.Lineno)
			Endcase

			Select (nSelect)
		Endif

	Endfunc


	* -- goto the definition of a reference
	Function GotoDefinition(cUniqueID)
		Local nSelect
		Local lSuccess
		Local cFilename
		Local cClassName
		Local cProcName
		Local cSymbol
		Local nCnt

		If Vartype(cUniqueID) <> 'C' Or Empty(cUniqueID)
			Return .F.
		Endif

		lSuccess = .F.
		If Used("FoxRefCursor") And Seek(cUniqueID, "FoxRefCursor", "UniqueID")
			nSelect = Select()

			cSymbol    = Upper(FoxRefCursor.Symbol)
			cFolder    = FoxRefCursor.Folder
			cFilename  = FoxRefCursor.Filename
			cClassName = FoxRefCursor.ClassName
			cProcName  = FoxRefCursor.ProcName

			* now find all appropriate definitions
			Select ;
				UniqueID, ;
				Symbol, ;
				DefType, ;
				Filename, ;
				Folder, ;
				ClassName, ;
				ProcName, ;
				RefCode, ;
				IIF(Filename == cFilename And ClassName == cClassName And ProcName == cProcName, 1, ;
				IIF(Filename == cFilename And ClassName == cClassName, 2, ;
				IIF(Filename == cFilename, 3, ;
				4))) As Context ;
				FROM (This.RefTable) ;
				WHERE ;
				UPPER(Symbol) == cSymbol And ;
				RefType == REFTYPE_DEFINITION And ;
				!Inactive ;
				ORDER By Context ;
				INTO Cursor DefinitionCursor
			nCnt = _Tally

			Do Case
				Case nCnt == 0
					* no matches found
					Messagebox(NODEFINITION_LOC + Chr(10) + Chr(10) + Rtrim(FoxRefCursor.Symbol), MB_ICONEXCLAMATION, GOTODEFINITION_LOC)
				Case nCnt == 1
					* only a single match, so go right to it
					This.GotoReference(DefinitionCursor.UniqueID)
				Otherwise
					* more than one match found, so display a cursor of
					* the available matches.
					Do Form FoxRefGotoDef With This, Rtrim(FoxRefCursor.Symbol)
			Endcase

			If Used("DefinitionCursor")
				Use In DefinitionCursor
			Endif

			Select (nSelect)
		Endif

		Return lSuccess
	Endfunc

	* Show a progress form while searching
	Function UpdateProgress(cMsg)
		If This.ShowProgress
			If Vartype(This.oProgressForm) <> 'O'
				This.lCancel = .F.
				Do Form FoxRefProgress Name This.oProgressForm Linked
			Endif
			If This.oProgressForm.SetProgress(cMsg)  && TRUE is returned if Cancel button is pressed
				If Messagebox(SEARCH_CANCEL_LOC, MB_ICONQUESTION + MB_YESNO, APPNAME_LOC) == IDYES
					This.lCancel = .T.
				Else
					This.oProgressForm.lCancel = .F.
				Endif
			Endif
			DoEvents
		Endif
	Endfunc

	Function CloseProgress()
		If Vartype(This.oProgressForm) == 'O'
			This.oProgressForm.Release()
		Endif
		This.lCancel = .F.
	Endfunc


	* Export reference table
	Function ExportReferences(cExportType, cFilename, lSelectedOnly)
		Local nSelect
		Local cFor

		nSelect = Select()
		Select 0

		If Vartype(cExportType) <> 'C' Or Empty(cExportType)
			cExportType = EXPORTTYPE_DBF
		Endif
		If Vartype(lSelectedOnly) <> 'L'
			lSelectedOnly = .F.
		Endif
		If Vartype(cFilename) <> 'C'
			Return .F.
		Endif
		cFilename = Fullpath(cFilename)
		If !Directory(Justpath(cFilename))
			Return .F.
		Endif



		cFor = "RefType == [" + REFTYPE_RESULT + "] AND !Inactive"
		If lSelectedOnly
			cFor = cFor + " AND Checked"
		Endif


		Do Case
			Case cExportType == EXPORTTYPE_DBF
				Use (This.RefTable) Alias ExportCursor In 0 Shared Again
				Select ExportCursor

				Copy To (cFilename) ;
					FIELDS ;
					Symbol, ;
					Folder, ;
					Filename, ;
					ClassName, ;
					ProcName, ;
					ProcLineNo, ;
					LineNo, ;
					ColPos, ;
					MatchLen, ;
					RefCode, ;
					TimeStamp ;
					FOR &cFor

			Case cExportType == EXPORTTYPE_TXT
				Use (This.RefTable) Alias ExportCursor In 0 Shared Again
				Select ExportCursor

				Copy To (cFilename) ;
					FIELDS ;
					Symbol, ;
					Folder, ;
					Filename, ;
					ClassName, ;
					ProcName, ;
					ProcLineNo, ;
					LineNo, ;
					ColPos, ;
					MatchLen, ;
					RefCode, ;
					TimeStamp ;
					FOR &cFor ;
					DELIMITED

			Case cExportType == EXPORTTYPE_XML
				If lSelectedOnly
					Select ;
						Symbol, ;
						Folder, ;
						Filename, ;
						ClassName, ;
						ProcName, ;
						ProcLineNo, ;
						LineNo, ;
						ColPos, ;
						MatchLen, ;
						RefCode, ;
						TimeStamp ;
						FROM (This.RefTable) ;
						WHERE ;
						RefType == REFTYPE_RESULT And ;
						Checked And ;
						!Inactive ;
						INTO Cursor ExportCursor
				Else
					Select ;
						Symbol, ;
						Folder, ;
						Filename, ;
						ClassName, ;
						ProcName, ;
						ProcLineNo, ;
						LineNo, ;
						ColPos, ;
						MatchLen, ;
						RefCode, ;
						TimeStamp ;
						FROM (This.RefTable) ;
						WHERE ;
						RefType == REFTYPE_RESULT And ;
						!Inactive ;
						INTO Cursor ExportCursor
				Endif

				If Empty(Justext(cFilename))
					cFilename = Forceext(cFilename, "xml")
				Endif

				Cursortoxml("ExportCursor", cFilename, 1, 514, 0, '1')
		Endcase

		If Used("ExportCursor")
			Use In ExportCursor
		Endif

		Select (nSelect)

		Return .T.
	Endfunc


	* Print a report of found references
	Function PrintReferences(lPreview, cSetID, lSelectedOnly)
		Local nSelect
		Local cWhere
		Local lSuccess
		Local cRptFile

		nSelect = Select()
		Select 0

		lSuccess = .F.

		If Vartype(lPreview) <> 'L'
			lPreview = .F.
		Endif
		If Vartype(lSelectedOnly) <> 'L'
			lSelectedOnly = .F.
		Endif
		If Vartype(cSetID) <> 'C'
			cSetID = ''
		Endif

		cRptFile = This.ReportFile
		If Empty(Justext(cRptFile))
			cRptFile = Forceext(cRptFile, ".frx")
		Endif

		If File(cRptFile)
			cWhere = "RefType == [" + REFTYPE_RESULT + "] AND !Inactive"
			If !Empty(cSetID)
				cWhere = cWhere + " AND SetID == [" + cSetID + "]"
			Endif
			If lSelectedOnly
				cWhere = cWhere + " AND Checked"
			Endif

			Select ;
				SetID, ;
				Symbol, ;
				Folder, ;
				Filename, ;
				ClassName, ;
				ProcName, ;
				ProcLineNo, ;
				LineNo, ;
				ColPos, ;
				MatchLen, ;
				RefCode, ;
				TimeStamp ;
				FROM (This.RefTable) ;
				WHERE &cWhere ;
				INTO Cursor RptCursor

			If _Tally > 0
				If lPreview
					Report Form (cRptFile) Preview
				Else
					Report Form (cRptFile) Noconsole To Printer
				Endif

				lSuccess = .T.
			Endif

			If Used("RptCursor")
				Use In RptCursor
			Endif
		Endif

		Select (nSelect)

		Return lSuccess
	Endfunc

	* retrieve a preference from the FoxPro Resource file
	Function RestorePrefs()
		Local nSelect
		Local lSuccess
		Local nMemoWidth

		Local Array FOXREF_LOOKFOR_MRU[10]
		Local Array FOXREF_REPLACE_MRU[10]
		Local Array FOXREF_FOLDER_MRU[10]
		Local Array FOXREF_FILETYPES_MRU[10]

		Local FOXREF_COMMENTS
		Local FOXREF_MATCHCASE
		Local FOXREF_WHOLEWORDSONLY
		Local FOXREF_SUBFOLDERS
		Local FOXREF_OVERWRITE

		nSelect = Select()

		lSuccess = .F.

		If File(RESOURCE_FILE)    && resource file not found.
			Use (RESOURCE_FILE) In 0 Shared Again Alias FoxResource
			If Used("FoxResource")
				nMemoWidth = Set('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper(Alltrim(Type)) == "PREFW" ;
					AND Upper(Alltrim(Id)) == RESOURCE_ID ;
					AND !Deleted()

				If Found() And !Empty(Data) And ckval == Val(Sys(2007, Data)) And Empty(Name)
					Restore From Memo Data Additive

					If Type("FOXREF_LOOKFOR_MRU") == 'C'
						=Acopy(FOXREF_LOOKFOR_MRU, This.aLookForMRU)
					Endif
					If Type("FOXREF_REPLACE_MRU") == 'C'
						=Acopy(FOXREF_REPLACE_MRU, This.aReplaceMRU)
					Endif
					If Type("FOXREF_FOLDER_MRU") == 'C'
						=Acopy(FOXREF_FOLDER_MRU, This.aFolderMRU)
					Endif
					If Type("FOXREF_FILETYPES_MRU") == 'C'
						=Acopy(FOXREF_FILETYPES_MRU, This.aFileTypesMRU)
					Endif

					If Type("FOXREF_COMMENTS") == 'N'
						This.Comments = FOXREF_COMMENTS
					Endif

					If Type("FOXREF_MATCHCASE") == 'L'
						This.MatchCase = FOXREF_MATCHCASE
					Endif
					If Type("FOXREF_WHOLEWORDSONLY") == 'L'
						This.WholeWordsOnly = FOXREF_WHOLEWORDSONLY
					Endif
					If Type("FOXREF_SUBFOLDERS") == 'L'
						This.SubFolders = FOXREF_SUBFOLDERS
					Endif
					If Type("FOXREF_OVERWRITE") == 'L'
						This.OverwritePrior = FOXREF_OVERWRITE
					Endif

					lSuccess = .T.
				Endif

				* if no preferences or filetypes are empty,
				* then set a default
				If Empty(This.aFileTypesMRU[1])
					This.aFileTypesMRU[1] = FILETYPES_DEFAULT
				Endif

				Set Memowidth To (nMemoWidth)

				Use In FoxResource
			Endif
		Endif

		Select (nSelect)

		Return lSuccess

	Endfunc

	* retrieve a preference from the FoxPro Resource file
	Function SavePrefs()
		Local nSelect
		Local lSuccess
		Local nMemoWidth
		Local nCnt
		Local cData

		Local Array aFileList[1]
		Local Array FOXREF_LOOKFOR_MRU[10]
		Local Array FOXREF_FOLDER_MRU[10]
		Local Array FOXREF_FILETYPES_MRU[10]

		Local FOXREF_COMMENTS
		Local FOXREF_MATCHCASE
		Local FOXREF_WHOLEWORDSONLY
		Local FOXREF_SUBFOLDERS
		Local FOXREF_OVERWRITE

		=Acopy(This.aLookForMRU, FOXREF_LOOKFOR_MRU)
		=Acopy(This.aReplaceMRU, FOXREF_REPLACE_MRU)
		=Acopy(This.aFolderMRU, FOXREF_FOLDER_MRU)
		=Acopy(This.aFileTypesMRU, FOXREF_FILETYPES_MRU)

		FOXREF_COMMENTS       = This.Comments
		FOXREF_MATCHCASE      = This.MatchCase
		FOXREF_WHOLEWORDSONLY = This.WholeWordsOnly
		FOXREF_SUBFOLDERS     = This.SubFolders
		FOXREF_OVERWRITE      = This.OverwritePrior
		FOXREF_FILETYPES      = This.FileTypes


		nSelect = Select()

		lSuccess = .F.

		* make sure Resource file exists and is not read-only
		nCnt = Adir(aFileList, RESOURCE_FILE)
		If nCnt > 0 And Atc('R', aFileList[1, 5]) == 0
			Use (RESOURCE_FILE) In 0 Shared Again Alias FoxResource
			If Used("FoxResource") And !Isreadonly("FoxResource")
				nMemoWidth = Set('MEMOWIDTH')
				Set Memowidth To 255

				Select FoxResource
				Locate For Upper(Alltrim(Type)) == "PREFW" And Upper(Alltrim(Id)) == RESOURCE_ID And Empty(Name)
				If !Found()
					Append Blank In FoxResource
					Replace ;
						Type With "PREFW", ;
						ID With RESOURCE_ID, ;
						ReadOnly With .F. ;
						IN FoxResource
				Endif

				If !FoxResource.ReadOnly
					Save To Memo Data All Like FOXREF_*

					Replace ;
						Updated With Date(), ;
						ckval With Val(Sys(2007, FoxResource.Data)) ;
						IN FoxResource

					lSuccess = .T.
				Endif
				Set Memowidth To (nMemoWidth)

				*** JRN 2010-04-02 : Pack
				* Curiously, even if there are no changes, savings causes file explosion, so we pack afterwards
				Try
					Use (RESOURCE_FILE) Exclusive Alias FoxResource
					Pack
				Catch

				Finally
					Use

				Endtry

			Endif
		Endif

		Select (nSelect)

		Return lSuccess
	Endfunc

	Function SaveWindowPosition(nTop, nLeft, nHeight, nWidth)
		Local lInUse

		If File(Forceext(This.RefTable, "DBF"))
			lInUse = Used(This.RefTable)
			If !lInUse
				Use (This.RefTable) Alias FoxRefCursor In 0 Shared Again
			Endif

			Goto Top In FoxRefCursor
			If Type("FoxRefCursor.RefType") == 'C' And FoxRefCursor.RefType == REFTYPE_INIT
				Replace RefCode With ;
					TRANSFORM(nTop) + ',' + Transform(nLeft) + ',' + Transform(nHeight) + ',' + Transform(nWidth) ;
					IN FoxRefCursor
			Endif

			If !lInUse And Used("FoxRefCursor")
				Use In FoxRefCursor
			Endif
		Endif
	Endfunc

	Function UpdateLookForMRU(cPattern)
		Local nRow

		nRow = Ascan(This.aLookForMRU, cPattern, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aLookForMRU, nRow)
		Endif
		=Ains(This.aLookForMRU, 1)
		This.aLookForMRU[1] = cPattern
	Endfunc

	Function UpdateReplaceMRU(cReplaceText)
		Local nRow

		nRow = Ascan(This.aReplaceMRU, cReplaceText, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aReplaceMRU, nRow)
		Endif
		=Ains(This.aReplaceMRU, 1)
		This.aReplaceMRU[1] = cReplaceText
	Endfunc

	Function UpdateFolderMRU(cFolder)
		Local nRow

		nRow = Ascan(This.aFolderMRU, cFolder, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aFolderMRU, nRow)
		Endif
		=Ains(This.aFolderMRU, 1)
		This.aFolderMRU[1] = cFolder
	Endfunc


	Function UpdateFileTypesMRU(cFileTypes)
		Local nRow

		nRow = Ascan(This.aFileTypesMRU, cFileTypes, -1, -1, 1, 15)
		If nRow > 0
			=Adel(This.aFileTypesMRU, nRow)
		Endif
		=Ains(This.aFileTypesMRU, 1)
		This.aFileTypesMRU[1] = cFileTypes
	Endfunc

	* stolen from Class Browser
	Function WildCardMatch(tcMatchExpList, tcExpressionSearched, tlMatchAsIs)
		Local lcMatchExpList,lcExpressionSearched,llMatchAsIs,lcMatchExpList2
		Local lnMatchLen,lnExpressionLen,lnMatchCount,lnCount,lnCount2,lnSpaceCount
		Local lcMatchExp,lcMatchType,lnMatchType,lnAtPos,lnAtPos2
		Local llMatch,llMatch2

		If Alltrim(tcMatchExpList) == "*.*"
			Return .T.
		Endif

		If Empty(tcExpressionSearched)
			If Empty(tcMatchExpList) Or Alltrim(tcMatchExpList) == "*"
				Return .T.
			Endif
			Return .F.
		Endif
		lcMatchExpList=Lower(Alltrim(Strtran(tcMatchExpList,Tab," ")))
		lcExpressionSearched=Lower(Alltrim(Strtran(tcExpressionSearched,Tab," ")))
		lnExpressionLen=Len(lcExpressionSearched)
		If lcExpressionSearched==lcMatchExpList
			Return .T.
		Endif
		llMatchAsIs=tlMatchAsIs
		If Left(lcMatchExpList,1)==["] And Right(lcMatchExpList,1)==["]
			llMatchAsIs=.T.
			lcMatchExpList=Alltrim(Substr(lcMatchExpList,2,Len(lcMatchExpList)-2))
		Endif
		If Not llMatchAsIs And " "$lcMatchExpList
			llMatch=.F.
			lnSpaceCount=Occurs(" ",lcMatchExpList)
			lcMatchExpList2=lcMatchExpList
			lnCount=0
			Do While .T.
				lnAtPos=At(" ",lcMatchExpList2)
				If lnAtPos=0
					lcMatchExp=Alltrim(lcMatchExpList2)
					lcMatchExpList2=""
				Else
					lnAtPos2=At(["],lcMatchExpList2)
					If lnAtPos2<lnAtPos
						lnAtPos2=At(["],lcMatchExpList2,2)
						If lnAtPos2>lnAtPos
							lnAtPos=lnAtPos2
						Endif
					Endif
					lcMatchExp=Alltrim(Left(lcMatchExpList2,lnAtPos))
					lcMatchExpList2=Alltrim(Substr(lcMatchExpList2,lnAtPos+1))
				Endif
				If Empty(lcMatchExp)
					Exit
				Endif
				lcMatchType=Left(lcMatchExp,1)
				Do Case
					Case lcMatchType=="+"
						lnMatchType=1
					Case lcMatchType=="-"
						lnMatchType=-1
					Otherwise
						lnMatchType=0
				Endcase
				If lnMatchType#0
					lcMatchExp=Alltrim(Substr(lcMatchExp,2))
				Endif
				llMatch2=This.WildCardMatch(lcMatchExp,lcExpressionSearched, .T.)
				If (lnMatchType=1 And Not llMatch2) Or (lnMatchType=-1 And llMatch2)
					Return .F.
				Endif
				llMatch=(llMatch Or llMatch2)
				If lnAtPos=0
					Exit
				Endif
			Enddo
			Return llMatch
		Else
			If Left(lcMatchExpList,1)=="~"
				Return (Difference(Alltrim(Substr(lcMatchExpList,2)),lcExpressionSearched)>=3)
			Endif
		Endif
		lnMatchCount=Occurs(",",lcMatchExpList)+1
		If lnMatchCount>1
			lcMatchExpList=","+Alltrim(lcMatchExpList)+","
		Endif
		For lnCount = 1 To lnMatchCount
			If lnMatchCount=1
				lcMatchExp=Lower(Alltrim(lcMatchExpList))
				lnMatchLen=Len(lcMatchExp)
			Else
				lnAtPos=At(",",lcMatchExpList,lnCount)
				lnMatchLen=At(",",lcMatchExpList,lnCount+1)-lnAtPos-1
				lcMatchExp=Lower(Alltrim(Substr(lcMatchExpList,lnAtPos+1,lnMatchLen)))
			Endif
			For lnCount2 = 1 To Occurs("?",lcMatchExp)
				lnAtPos=At("?",lcMatchExp)
				If lnAtPos>lnExpressionLen
					If (lnAtPos-1)=lnExpressionLen
						lcExpressionSearched=lcExpressionSearched+"?"
					Endif
					Exit
				Endif
				lcMatchExp=Stuff(lcMatchExp,lnAtPos,1,Substr(lcExpressionSearched,lnAtPos,1))
			Endfor
			If Empty(lcMatchExp) Or lcExpressionSearched==lcMatchExp Or ;
					lcMatchExp=="*" Or lcMatchExp=="?" Or lcMatchExp=="%%"
				Return .T.
			Endif
			If Left(lcMatchExp,1)=="*"
				Return (Substr(lcMatchExp,2)==Right(lcExpressionSearched,Len(lcMatchExp)-1))
			Endif
			If Left(lcMatchExp,1)=="%" And Right(lcMatchExp,1)=="%" And ;
					SUBSTR(lcMatchExp,2,lnMatchLen-2)$lcExpressionSearched
				Return .T.
			Endif
			lnAtPos=At("*",lcMatchExp)
			If lnAtPos>0 And (lnAtPos-1)<=lnExpressionLen And ;
					LEFT(lcExpressionSearched,lnAtPos-1)==Left(lcMatchExp,lnAtPos-1)
				Return .T.
			Endif
		Endfor
		Return .F.
	Endfunc

Enddefine
