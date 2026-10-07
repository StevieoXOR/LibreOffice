The parent folder, named `config`, should go in the following location, if
1) you haven't changed any of LibreOffice's default filepath configurations and
2) assuming you're on Windows:
* C:\Users\<COMPUTER_USERNAME>\AppData\Roaming\LibreOffice\4\user\<`config`_FILE_GOES_HERE>

Note:
Since the only Modules I use are just Math and Writer (not Draw, Impress, Calc, ...), only those two folders will contain non-empty folders (pertaining to toolbars).
* If you for some reason need to reconstruct all Modules' folders:
  * List of Modules (each is a folder inside `soffice.cfg/modules/`:  BasicIDE, dbapp, scalc, sdraw, simpress, smath, StartModule, swriter
    * List of all folders inside each module folder:  menubar, popupmenu, statusbar, toolbar
      * The sole exception to this is `swriter`, which has an extra folder named `ui` with a file named `notebookbar.ui` inside it. I.e., `config/soffice.cfg/modules/swriter/ui/notebookbar.ui` exists.