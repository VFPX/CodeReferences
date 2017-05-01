Lparameters tlDebugInfo

Local lcAppFile, lcProject

lcProject = 'FoxRef.PJX'
Select 0
Use (lcProject) Again
Locate For Type = 'H'
Replace Debug With .T. && tlDebugInfo
Use

lcAppFile = '..\FoxRef.APP'
Erase (lcAppFile)
Build App (lcAppFile) From FoxRef && Recompile
