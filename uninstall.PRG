#include FoxPro.h

#Define	SOURCE_FOLDER    Home()
#Define INSTALL_MESSAGE  'Original VFP "Code References" re-installed'

#Define CR				 Chr(13)
#Define WARNING_MESSAGE  'Warning: may not take effect until VFP restarted'

Local loRegistry As 'FoxReg' Of  'Registry.vcx'
Local lcFoxRefApp

lcFoxRefApp = Lower (Addbs (SOURCE_FOLDER) + "FoxRef.APP")

If File (lcFoxRefApp)
	loRegistry = Newobject('FoxReg', Home() + 'FFC\Registry.vcx')
	loRegistry.SetFoxOption('_FOXREF', ["] + lcFoxRefApp + ["])
	_FoxRef = lcFoxRefApp
	Messagebox (INSTALL_MESSAGE + CR + CR + WARNING_MESSAGE, MB_OK + MB_ICONINFORMATION, 'Un-installed')
Else
	Messagebox ("Executable " + lcFoxRefApp + " not found", MB_OK + MB_ICONEXCLAMATION, 'Un-installation failed')
Endif
