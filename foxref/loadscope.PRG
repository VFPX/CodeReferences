* Abstract:
*   Load up list of available scopes, including open Projects.
*
* Parameters:
*	[cGotoProjName]
#Include "foxref.h"
Lparameters cGotoProjName

Local llAddIt, lnRow, loProjects
Local cFindScope, cProcNameX, cProjName, nGotoIndex, nProjectIndex
With This.cboScope
	If Vartype (cGotoProjName) # 'C' Or Empty (cGotoProjName)
		cFindScope	  = This.cboScope.Value
		cGotoProjName = .List (.ListIndex, 3)
	Else
		cFindScope = SCOPE_PROJECT
	Endif
	m.cGotoProjName = Lower (m.cGotoProjName)

	.Clear()

	m.nGotoIndex = 1
	For nProjectIndex = 1 To Application.Projects.Count
		m.cProjName	 = Application.Projects (nProjectIndex).Name
		m.cProcNameX = IIf (Left (m.cProjName, 1) == '\', '\', '') + m.cProjName
		.AddItem (Strtran ('Project: ' + Thisform.GetDisplayRelativePath (Lower (cProjName)), '.PJX', '', 1, 1, 1))
		.AddListItem (SCOPE_PROJECT, .NewItemId, 2)
		.AddListItem (m.cProcNameX, .NewItemId, 3)
		If Lower (m.cProjName) == m.cGotoProjName And m.cFindScope == SCOPE_PROJECT
			m.nGotoIndex = .NewIndex
		Endif
	EndFor
	*** JRN 2010-08-21 : Added in MRU Projects from VFP's list
	loProjects = This.oFoxRef.GetMRUList ('MRUL')
	For lnRow = 1 To loProjects.Count
		cProjName = loProjects.Item (lnRow)
		m.cProcNameX = IIf (Left (m.cProjName, 1) == '\', '\', '') + m.cProjName
		llAddIt	  = .T.

		For nProjectIndex = 1 To Application.Projects.Count
			If Lower (m.cProcNameX) == Lower (.List (nProjectIndex, 3))
				llAddIt = .F.
			Endif
		Endfor

		If llAddIt
			.AddItem (Strtran ('Project: ' + Thisform.GetDisplayRelativePath (Lower (cProjName)), '.PJX', '', 1, 1, 1))
			.AddListItem (SCOPE_PROJECT, .NewItemId, 2)
			.AddListItem (m.cProcNameX, .NewItemId, 3)
			If Lower (m.cProjName) == m.cGotoProjName And m.cFindScope == SCOPE_PROJECT
				m.nGotoIndex = .NewIndex
			Endif
		Endif
	Endfor

	.AddItem (SCOPE_FOLDER_LOC)
	.AddListItem (SCOPE_FOLDER, .NewItemId, 2)
	.AddListItem ('', .NewItemId, 3)
	If SCOPE_FOLDER == m.cFindScope
		m.nGotoIndex = .NewIndex
	Endif

	.ListIndex = m.nGotoIndex
Endwith
