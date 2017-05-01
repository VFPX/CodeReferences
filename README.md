# Code References Version 1.2 Beta 

Project Manager: Jim Nelson

![](Code%20References_CodeReferences2.png)

The Code References project provides a way for you to search for occurrences of a code reference in a project, specific folder or subfolder. You can filter your search by file type. 

Code References is part of [XSource](https://github.com/VFPX/XSource), the source files for various Visual FoxPro components. The license governing XSource can be found in the XSource_EULA.txt included with all of the XSource releases.

**Enhancement in this release (since 1.1 Beta)**

* There is a new 'Code Filter' box at the top of the form, which acts as a filter on the Code column in the grid.  This performs a case-insensitive search ($) and remains in effect until removed.
* Class libraries listed in the TreeView on the left have sub-nodes corresponding to all the individual classes within them.
* All files opened by double-clicking on the results grid update the appropriate VFP MRU list, and the case of the file names is maintained.
* Drop-down list of projects uses the VFP MRU list of projects (instead of only those that are open).  Selecting a project from the list will open that project if it is closed.
* There are two new options in the Options form:
	* Files opened by double-clicking on the results grid can be checked out using Source Control.
	* Files listed in the TreeView on the left can be listed with their folder names (relative to the project or current folder).

## Summary of all Enhancements

### Search Screen:
* CommandButton to select current folder
* Drop down list for ‘Scope’ includes the entire VFP MRU list of projects.  Selecting a project from the list will open that project if it is closed.
* The search form is now re-sizable (well, at least, it can be made wider).
* Changes to file templates:
	* Searching for blanks in file names now supported
	* Blanks no longer supported as delimiters between templates; valid separators are comma and semi-colon

### New Searching Capabilities
* Included in search:
	* Class names (searches now done on Class and ClassLoc columns of VCX/SCX files)
	* Names of PRG files
* Regular expressions recognize continuation lines

### Results Screen
* Dockable; retains docking information in new session
* Does not allow items to be checked for replacement if they can’t be replaced:
	* Method and procedure names, etc, (as before)
	* Regular expressions (unless using no special characters)
* File names are shown as relative to the project or current directory
* Some re-arrangement of display between the 'class' and 'method' columns; new values in 'method' column
	* &lt;Class Def&gt;
	* &lt;Class&gt;
	* &lt;Property Def&gt;
	* &lt;Property&gt;
	* &lt;Method Def&gt;
	* &lt;Method&gt;
	* &lt;PRG File&gt;
	* &lt;Include File&gt;
	* &lt;Object&gt;
	* &lt;Procedure&gt;
	* &lt;Function&gt;
* A new 'Code Filter' box at the top of the form acts as a filter on the Code column in the grid.  This performs a case-insensitive search ($) and remains in effect until removed.
* All files opened by double-clicking on the results grid update the appropriate VFP MRU list, and the case of the file names is maintained.
* Changes in the TreeView
	* Class libraries listed in the TreeView on the left have sub-nodes corresponding to all the individual classes within them.
	* A new option (Options form) causes files list in the TreeView to be listed with their folder names (relative to the project or current folder).
* New column displays timestamp:
	* For VCXs and SCXs, the timestamp from row of source file
	* For all other sources, the timestamp of source file
* Cascading (i.e., multi-column) file sorts
* New sort options, by Folder or Extension, now available from right-click context menu on the grid cells.
* Sorting now allowed on ‘Code’ column

### Options form
* Files opened by double-clicking on the results grid can be checked out using Source Control.
* Files listed in the TreeView on the left can be listed with their folder names (relative to the project or current folder).

### Bug Fixes
* Minor bug with Whole Word not matching word starting at beginning of file / method 
* Column width had not been preserved between sessions for one of the columns

### Other / Miscellaneous
* Install / uninstall programs 
* Files that are opened are recorded in ‘MRU Files’ in PEM Editor (if open)
* Version # displayed
* Uses own resource file – in Home(7)
* Processing time is displayed in a WAIT WINDOW

## Release History
* Ver 1.2 Beta - Released 2010-10-09
* Ver 1.1 Beta 2 - Released 2010-05-17 (430 downloads)
* Ver 1.1 Beta 1 - Released 2010-04-13 (187 downloads)
