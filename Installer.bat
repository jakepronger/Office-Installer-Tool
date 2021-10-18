@echo off
title Office 2021 Installer

:mainmenu
color 1F
net.exe session 1>nul 2>nul
if %errorlevel% EQU 2 (
    mshta vbscript:execute("CreateObject(""shell.Application"").shellExecute ""%~S0"", """", """", ""runas"", 1"^)^(window.close^)
    exit /b
)
cls
color 1F
echo[
echo Office 2021 Installer
echo[
echo Choose an option:
echo -----------------
echo 1. Download and install office.
echo 2. Download office to data folder to install later. (Downloaded save)
echo 3. Install office from an existing downloaded save.
echo 4. Create or view existing configuration files.
echo 5. View office applications installed.
echo 6. Repair office.
echo 7. Uninstall office.
echo 8. Delete all data.
echo 9. Exit to command prompt.
echo 10. Exit Installer.
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto download_and_install
) ELSE IF "%selection%" == "2" (
    goto download_install_later
) ELSE IF "%selection%" == "3" (
    goto install
) ELSE IF "%selection%" == "4" (
    goto configuration_files
) ELSE IF "%selection%" == "5" (
    goto viewinstalled
) ELSE IF "%selection%" == "6" (
    goto repair
) ELSE IF "%selection%" == "7" (
    goto uninstall
) ELSE IF "%selection%" == "8" (
    cls
    color 1F
    echo Removing all data...
    DEL /Q "%~dp0Data\SavedConfigurations\*"
    DEL /Q "%~dp0Data\SavedInstallers\*"
    CALL :REMOVETEMPFILES
    echo Removed all information from folders, TempInstall, SavedConfigurations and SavedInstallers.
    pause
    goto mainmenu
) ELSE IF "%selection%" == "9" (
    start cmd.exe
    exit
) ELSE IF "%selection%" == "10" (
    exit
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto mainmenu
)


:download_and_install
color 1F
cls
echo[
echo Office 2021 Customized - Download and install
echo[
echo Choose an option:
echo -----------------
echo 1. All office applications
echo 2. Choose office applications
echo 3. Go back
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto fulldownload
) ELSE IF "%selection%" == "2" (
    goto download_and_install_custom
) ELSE IF "%selection%" == "3" (
    goto mainmenu
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto download_and_install
)

:fulldownload
color 1F
cls
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    COLOR 4F
    cls
    echo[ &echo Office is already installed! &echo[
    pause
    goto download_and_install
) ELSE (
    echo Deleting files in TempInstall folder...
    CALL :REMOVETEMPFILES
    echo Checking for existing shortcuts to remove...
    CALL :deletedesktopshortcuts
    CALL :deletesearchshortcuts
    (echo ^<Configuration^>, &echo   ^<Add SourcePath^=^"%~dp0Data\TempInstall\^" OfficeClientEdition^=^"64^" Channel^=^"PerpetualVL2021^"^>&echo     ^<Product ID^=^"ProPlus2021Volume^"^> PIDKEY^=^"BN8D3-W2QKT-M7Q73-Y3JWK-KQC63^"^>&echo        ^<Language ID^=^"MatchOS^" ^/^>&echo     ^<^/Product^>&echo   ^<^/Add^>&echo   ^<Display Level^=^"None^" AcceptEULA^=^"TRUE^" ^/^>&echo ^<^/Configuration^>) >"%~dp0Data\TempInstall\config.xml"
    echo Created configuration file.
    echo Downloading files from Microsoft...
    "%~dp0Data\setup.exe" /download "%~dp0Data\TempInstall\config.xml"
    echo Installing Office...
    "%~dp0Data\setup.exe" /configure "%~dp0Data\TempInstall\config.xml"
    echo Verifying installation folder...
    IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
        CALL :REMOVETEMPFILES
        echo Deleted files in TempInstall folder.
        echo Creating shortcuts to the desktop...
        CALL :copyappshortcuts
        echo[
        echo Setup has been completed. Office 2021 has been installed.
        echo[
        pause
        goto mainmenu
    ) ELSE (
        CALL :REMOVETEMPFILES
        COLOR 4F
        cls
        echo[
        echo ERROR [Installation failed. Please restart your PC and try again.]
        echo[
        pause
        goto mainmenu
    )
)



:download_and_install_custom
color 1F
cls
echo[
echo Choose an option:
echo -----------------
echo 1. Choose applications to exclude and install now.
echo 2. Install using a configuration file in SavedConfigurations folder.
echo 3. Go back
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto download_and_install_custom_chooseapps_resetvars
) ELSE IF "%selection%" == "2" (
	goto download_and_install_custom_chooseconfig
) ELSE IF "%selection%" == "3" (
    goto download_and_install
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto download_and_install_custom
)

:download_and_install_custom_chooseapps
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    COLOR 4F
    cls
    echo[ &echo Office is already installed! &echo[
    pause
    goto download_and_install_custom
)
color 1F
cls
echo[
echo Use the numbers on the side to exclude ^& include applications.
echo[
IF "%Access%" =="" set Access=[92mINCLUDED
IF "%Excel%" ==""  set Excel=[92mINCLUDED
IF "%Groove%" =="" set Groove=[91mEXCLUDED
IF "%Lync%" ==""  set Lync=[92mINCLUDED
IF "%OneDrive%" =="" set OneDrive=[91mEXCLUDED
IF "%OneNote%" =="" set OneNote=[92mINCLUDED
IF "%Outlook%" =="" set Outlook=[92mINCLUDED
IF "%PowerPoint%" =="" set PowerPoint=[92mINCLUDED
IF "%Publisher%" =="" set Publisher=[92mINCLUDED
IF "%Teams%" =="" set Teams=[91mEXCLUDED
IF "%Word%" =="" set Word=[92mINCLUDED
IF "%ShortcutsRemove%" =="" set ShortcutsRemove=[92mINCLUDED
IF "%ShortcutsDesktop%" =="" set ShortcutsDesktop=[92mINCLUDED
echo Choose applications to exclude:
echo ---------------------------------
echo 1. Access                              ^|  ^[%Access%[97m^]
echo 2. Excel                               ^|  ^[%Excel%[97m^]
echo 3. Groove (Client for SharePoint)      ^|  ^[%Groove%[97m^]
echo 4. Lync (Skype for Business)           ^|  ^[%Lync%[97m^]
echo 5. OneDrive                            ^|  ^[%OneDrive%[97m^]
echo 6. OneNote                             ^|  ^[%OneNote%[97m^]
echo 7. Outlook                             ^|  ^[%Outlook%[97m^]
echo 8. PowerPoint                          ^|  ^[%PowerPoint%[97m^]
echo 9. Publisher                           ^|  ^[%Publisher%[97m^]
echo 10. Teams                              ^|  ^[%Teams%[97m^]
echo 11. Word                               ^|  ^[%Word%[97m^]
echo[
echo Other options:
echo ---------------
echo 12. Remove existing shortcuts.         ^|  ^[%ShortcutsRemove%[97m^]
echo 13. Copy shortcuts to the desktop.     ^|  ^[%ShortcutsDesktop%[97m^]
echo ^*. Add shortcuts to the taskbar.       ^|  ^[[91mUNAVAILABLE[97m^]
echo[
echo 14. Go back
echo 15. Agree and continue to setup.
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" ( CALL :include_and_exclude %Access%, Access, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "2" ( CALL :include_and_exclude %Excel%, Excel, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "3" ( CALL :include_and_exclude %Groove%, Groove, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "4" ( CALL :include_and_exclude %Lync%, Lync, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "5" ( CALL :include_and_exclude %OneDrive%, OneDrive, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "6" ( CALL :include_and_exclude %OneNote%, OneNote, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "7" ( CALL :include_and_exclude %Outlook%, Outlook, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "8" ( CALL :include_and_exclude %PowerPoint%, PowerPoint, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "9" ( CALL :include_and_exclude %Publisher%, Publisher, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "10" ( CALL :include_and_exclude %Teams%, Teams, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "11" ( CALL :include_and_exclude %Word%, Word, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "12" ( CALL :include_and_exclude %ShortcutsRemove%, ShortcutsRemove, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "13" ( CALL :include_and_exclude %ShortcutsDesktop%, ShortcutsDesktop, download_and_install_custom_chooseapps
) ELSE IF "%selection%" == "14" (
    goto download_and_install_custom
) ELSE IF "%selection%" == "15" (
    goto download_and_install_custom_chooseapps_install
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto download_and_install_custom_chooseapps
)

:download_and_install_custom_chooseapps_resetvars
CALL :resetappvars
set ShortcutsRemove=
set ShortcutsDesktop=
goto download_and_install_custom_chooseapps

:download_and_install_custom_chooseapps_install
color 1F
cls
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    CALL :resetappvars
    set ShortcutsRemove=
    set ShortcutsDesktop=
    COLOR 4F
    cls
    echo[ &echo Office is already installed! &echo[
    pause
    goto download_and_install
)
echo Deleting files in TempInstall folder...
CALL :REMOVETEMPFILES
(echo ^<Configuration^> &echo   ^<Add SourcePath^=^"%~dp0Data\TempInstall\^" OfficeClientEdition^=^"64^" Channel^=^"PerpetualVL2021^"^> &echo     ^<Product ID^=^"ProPlus2021Volume^"^> PIDKEY^=^"BN8D3-W2QKT-M7Q73-Y3JWK-KQC63^"^> &echo        ^<Language ID^=^"MatchOS^" ^/^>
IF "%Access%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Access" ^/^>
IF "%Excel%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Excel" ^/^>
IF "%Groove%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Groove" ^/^>
IF "%Lync%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Lync" ^/^>
IF "%OneDrive%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="OneDrive" ^/^>
IF "%OneNote%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="OneNote" ^/^>
IF "%Outlook%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Outlook" ^/^>
IF "%PowerPoint%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="PowerPoint" ^/^>
IF "%Publisher%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Publisher" ^/^>
IF "%Teams%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Teams" ^/^>
IF "%Word%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Word" ^/^>
echo     ^<^/Product^>&echo   ^<^/Add^> &echo   ^<Display Level^=^"None^" AcceptEULA^=^"TRUE^" ^/^> &echo ^<^/Configuration^>) >"%~dp0Data\TempInstall\config.xml"
echo Created configuration file.
IF "%ShortcutsRemove%" == "[92mINCLUDED" (
    echo Checking for existing shortcuts to remove...
    CALL :deletedesktopshortcuts
    CALL :deletesearchshortcuts
)
echo Downloading files from Microsoft...
"%~dp0Data\setup.exe" /download "%~dp0Data\TempInstall\config.xml"
echo Installing Office...
"%~dp0Data\setup.exe" /configure "%~dp0Data\TempInstall\config.xml"
echo Verifying installation folder...
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    CALL :REMOVETEMPFILES
    echo Deleted files in TempInstall folder.
    IF "%ShortcutsDesktop%" == "[92mINCLUDED" (
        echo Creating shortcuts to the desktop...
        CALL :copyappshortcuts
    )
    echo[ &echo Setup has been completed. Office 2021 Customized has been installed. &echo[
    pause
    cls
    echo[ &echo The following has been installed^:
    IF "%Access%" == "[92mINCLUDED" echo ^* Access
    IF "%Excel%" == "[92mINCLUDED" echo ^* Excel
    IF "%Groove%" == "[92mINCLUDED" echo ^* Groove
    IF "%Lync%" == "[92mINCLUDED" echo ^* Lync
    IF "%OneDrive%" == "[92mINCLUDED" echo ^* OneDrive
    IF "%OneNote%" == "[92mINCLUDED" echo ^* OneNote
    IF "%Outlook%" == "[92mINCLUDED" echo ^* Outlook
    IF "%PowerPoint%" == "[92mINCLUDED" echo ^* PowerPoint
    IF "%Publisher%" == "[92mINCLUDED" echo ^* Publisher
    IF "%Teams%" == "[92mINCLUDED" echo ^* Teams
    IF "%Word%" == "[92mINCLUDED" echo ^* Word
    echo[
    echo The following were NOT installed^:
    IF "%Access%" == "[91mEXCLUDED" echo ^* Access
    IF "%Excel%" == "[91mEXCLUDED" echo ^* Excel
    IF "%Groove%" == "[91mEXCLUDED" echo ^* Groove
    IF "%Lync%" == "[91mEXCLUDED" echo ^* Lync
    IF "%OneDrive%" == "[91mEXCLUDED" echo ^* OneDrive
    IF "%OneNote%" == "[91mEXCLUDED" echo ^* OneNote
    IF "%Outlook%" == "[91mEXCLUDED" echo ^* Outlook
    IF "%PowerPoint%" == "[91mEXCLUDED" echo ^* PowerPoint
    IF "%Publisher%" == "[91mEXCLUDED" echo ^* Publisher
    IF "%Teams%" == "[91mEXCLUDED" echo ^* Teams
    IF "%Word%" == "[91mEXCLUDED" echo ^* Word
    echo[
    IF "%ShortcutsRemove%" or "%ShortcutsDesktop%" == "[92mINCLUDED" echo The following actions have occured:
    IF "%ShortcutsRemove%" == "[92mINCLUDED" echo ^* Existing shortcuts were removed before installation.
    IF "%ShortcutsDesktop%" == "[92mINCLUDED" echo ^* Shortcuts have been added to the desktop.
    IF "%ShortcutsRemove%" or "%ShortcutsDesktop%" == "[92mINCLUDED" echo[
    CALL :resetappvars
    set ShortcutsRemove=
    set ShortcutsDesktop=
    echo Press any key to continue to main menu.
    pause >nul
    goto mainmenu
) ELSE (
    CALL :resetappvars
    set ShortcutsRemove=
    set ShortcutsDesktop=
    CALL :REMOVETEMPFILES
    COLOR 4F
    cls
    echo[ &echo ERROR [Installation failed. Please restart your PC and try again.] &echo[
    pause
    goto mainmenu
)





:download_and_install_custom_chooseconfig
color 1F
cls
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    COLOR 4F
    cls
    echo[ &echo Office is already installed! Try uninstalling it first! &echo[
    pause
    goto download_and_install_custom
    )
echo[ &echo To go back press enter without any text. &echo[ &echo Enter the name of the configuration file: (Without .xml) &echo[
set "selection="
set /p selection=
IF "%selection%" == "" (
    goto download_and_install_custom )
IF EXIST "%~dp0Data\SavedConfigurations\%selection%.xml" (
    COLOR 1F
    cls
    echo Downloading files from Microsoft...
    "%~dp0Data\setup.exe" /download "%~dp0Data\SavedConfigurations\%selection%.xml".
    echo Installing Office...
    "%~dp0Data\setup.exe" /configure "%~dp0Data\SavedConfigurations\%selection%.xml"
    echo Verifying installation files...
    IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
        CALL :REMOVETEMPFILES
		echo Deleted files in TempInstall folder.
        echo Creating shortcuts to the desktop...
        CALL :copyappshortcuts
        cls
        echo[ &echo Custom Office 2021 has been installed. &echo[
        echo What got installed? &echo[
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" echo ^* Access
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" echo ^* Excel
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" echo ^* OneNote
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" echo ^* Outlook
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" echo ^* PowerPoint
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE" echo ^* Publisher
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\LYNC.EXE" echo ^* Publisher
        IF EXIST "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" echo ^* Word
        echo[
        echo Shortcuts copied to the desktop ^[[92mYES[97m^]
        echo Shortcuts copied to the desktop ^[[91mNO[97m^]
        echo[
        pause
        goto mainmenu
    ) ELSE (
        COLOR 4F
        cls
        echo[ &echo ERROR [Installation failed to install office.] &echo[
        pause
        goto download_and_install_custom
    )
) ELSE (
    COLOR 4F
    cls
    echo "%~dp0Data\SavedConfigurations\%selection%.xml"
    echo[ &echo ERROR [File specified does not exist in SavedConfigurations folder.] &echo[
    pause
    goto download_and_install_custom_chooseconfig
)



:download_install_later
color 1F
cls
echo[ &echo Download office to install later. &echo[
echo 1. Choose apps now and download to install later.
echo 2. Choose a configuration file to download and install later.
echo 3. Go back.
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto download_install_later_chooseapps
) ELSE IF "%selection%" == "2" (
    goto download_install_later_chooseconfig
) ELSE IF "%selection%" == "3" (
    goto mainmenu
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto download_install_later
)

:download_install_later_chooseconfig
color 1F
cls
echo[ &echo Type the name of the configuration file below. ^(Must be in savedconfigurations folder.^) &echo[
set "configfilename="
set /p configfilename=
IF "%configfilename%" == "" (
    goto download_install_later
) ELSE (
    IF EXIST "%~dp0Data\SavedConfigurations\%configfilename%.xml" (
        cls
        echo What do you want your save to be called? echo[
        set "selection="
        set /p selection=
        IF EXIST "%~dp0Data\SavedInstallers\%configfilename%\" (
            echo save already exists
            pause
            goto download_install_later_chooseconfig
        ) ELSE (
            CALL :REMOVETEMPFILES
            echo cleared temp install folder
            MKDIR "%~dp0Data\SavedInstallers\%configfilename%\"
            xcopy /F "%~dp0Data\SavedConfigurations\%configfilename%.xml" "%~dp0Data\SavedInstallers\%configfilename%\config.xml"
            echo save made now Downloading
            %~dp0Data\setup.exe /download %~dp0Data\SavedInstallers\%configfilename%\config.xml
            echo copying data to saved installers folder...
            xcopy "%~dp0Data\TempInstall\Office" "%~dp0Data\SavedInstallers\%configfilename%\*" /E
            echo downloaded.
            echo clearing temp install files
            CALL :REMOVETEMPFILES
            echo done
            pause
            goto download_install_later
        )
        pause
        goto download_install_later
    ) ELSE (
        echo config doesnt exist
        pause
        goto download_install_later
    )
)



:confirm
color 1F
cls
echo[
echo Office 2021 Custom
echo[
echo You are agreeing to install the following applications:
echo * Word
echo * Powerpoint
echo * Excel
echo * Access
echo[
echo Choose an option:
echo -----------------
echo 1. Confirm and continue installation.
echo 2. Return to main menu
echo[

set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto setup
) else if "%selection%" == "2" (
    goto mainmenu
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto confirm
)

:setup
COLOR 1F
cls
echo[
echo Select an option:
echo ------------------
echo 1. Download office from the web. (Will wipe all office data!)
echo 2. Download office from the web and install and remove all files afterwards.
echo 3. Install office from an existant save.
echo 4. Go back.
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto download_install_later
) else if "%selection%" == "2" (
    goto download_and_install
) else if "%selection%" == "3" (
    goto install
) else if "%selection%" == "4" (
    goto confirm
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto setup
)

:repair
color 1F
cls
echo Verifying office is installed...
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    echo Opening office repair menu...
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=repair platform=x64 culture=en-us
    echo Process complete.
    echo[
    pause
    goto mainmenu
) ELSE (
    COLOR 4F
    cls
    echo[ &echoSetup failed to repair. ERROR [Office is not currently installed on this system.] &echo[
    pause
    goto mainmenu
)

:uninstall
color 1F
cls
net session >nul 2>&1
echo Checking if office is currently installed...
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    echo Uninstalling office...
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" scenario=install scenariosubtype=ARP source=None productstoremove=ProPlus2021Volume.16_en-us_x-none culture=en-us verion.16=16.0
    CALL :REMOVETEMPFILES
    IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
        COLOR 4F
        cls
        echo[
        echo Setup failed to uninstall. ERROR [Uninstaller failed to uninstall office.]
        echo[
        pause
        goto mainmenu
    ) ELSE (
    echo Removing existing shortcuts on the desktop...
    CALL :deletedesktopshortcuts
    CALL :deletesearchshortcuts
    echo Uninstalled.
    echo[
    pause
    goto mainmenu
    )
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Setup failed to uninstall. ERROR [Office is not currently installed on this system.] &echo[
    pause
    goto mainmenu
)


:Install
color 1F
cls
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    COLOR 4F
    cls
    echo[ &echo Office is already installed! &echo[
    pause
    goto mainmenu
)
echo[
echo Install office from a previous download that was saved. Enter below the name of your save.
echo[
set "selection="
set /p selection=
IF "%selection%" == "" (
    goto mainmenu
) ELSE (
    IF EXIST "%~dp0Data\SavedInstallers\%selection%\" (
        IF EXIST "%~dp0Data\SavedInstallers\%selection%\config.xml" (
            cls
            echo Deleting temporary files...
            CALL :REMOVETEMPFILES
            echo Deleting existing search index shortcuts...
            CALL :deletesearchshortcuts
            echo Deleting existing desktop shortcuts
            CALL :deletedesktopshortcuts
            echo Copying data...
            xcopy "%~dp0Data\SavedInstallers\%selection%\" "%~dp0Data\TempInstall\Office\" /E
            echo Installing...
            %~dp0Data\setup.exe /configure %~dp0Data\SavedInstallers\%selection%\config.xml
            echo Copying app shortcuts to the desktop...
            CALL :copyappshortcuts
            echo Deleting temporary files...
            CALL :REMOVETEMPFILES
            echo[ &echo Completed.
            pause
        ) ELSE (
            cls
            color 4F
            echo[ &echo Configuration file doesn't exist in this save. &echo[
            pause
            goto install )
    ) ELSE (
        cls
        color 4F
        echo[ &echo File doesn't exist. &echo[
        pause
        goto install )
)


:viewinstalled
color 1F
cls
IF EXIST "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe" (
    echo[ &echo Office applications that are currently installed include^:
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE" echo ^* Access
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" echo ^* Excel
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE" echo ^* OneNote
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" echo ^* Outlook
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" echo ^* PowerPoint
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE" echo ^* Publisher
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\LYNC.EXE" echo ^* Publisher
    IF EXIST "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" echo ^* Word
    echo[
    pause
    goto mainmenu
) ELSE (
    color 4F
    cls
    echo[ &echo Failed to view existing applications. ERROR [Office is not currently installed on this system.] &echo[
    pause
    goto mainmenu)



:configuration_files
color 1F
cls
echo[ &echo Choose an option^:
echo -------------------------
echo 1. Create a configuration file to use later.
echo 2. View a configuration file.
echo 3. Go back.
set "selection="
set /p selection=
IF "%selection%" == "1" (
    goto configuration_files_create
) ELSE IF "%selection%" == "2" (
    goto viewconfiguration
) ELSE IF "%selection%" == "3" (
    goto mainmenu
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto configuration_files)


:viewconfiguration
color 1F
cls
echo[ &echo View what apps are included ^& excluded in a configuration file! &echo[ &echo Type the name of your configuration file below thats in savedconfigurations folder. &echo[
set "selection="
set /p selection=
IF "%selection%" == "" (
    goto configuration_files )
IF EXIST "%~dp0Data\SavedConfigurations\%selection%.xml" (
    echo exists
    %~dp0Data\SavedConfigurations\%selection%.xml
    pause
    goto configuration_files
) ELSE (
    color 4F
    cls
    echo[ &echo ERROR [File specified does not exist in SavedConfigurations folder.] &echo[
    pause
    goto viewconfiguration
)

:configuration_files_create
color 1F
cls
echo[ &echo Create a configuration file. Use the numbers on the side to exclude ^& include applications. &echo[
IF "%Access%" =="" set Access=[92mINCLUDED
IF "%Excel%" =="" set Excel=[92mINCLUDED
IF "%Groove%" =="" set Groove=[91mEXCLUDED
IF "%Lync%" =="" set Lync=[92mINCLUDED
IF "%OneDrive%" =="" set OneDrive=[91mEXCLUDED
IF "%OneNote%" =="" set OneNote=[92mINCLUDED
IF "%Outlook%" =="" set Outlook=[92mINCLUDED
IF "%PowerPoint%" =="" set PowerPoint=[92mINCLUDED
IF "%Publisher%" =="" set Publisher=[92mINCLUDED
IF "%Teams%" =="" set Teams=[91mEXCLUDED
IF "%Word%" =="" set Word=[92mINCLUDED
echo Choose applications to exclude ^& include:
echo ---------------------------------
echo 1. Access                              ^|  ^[%Access%[97m^]
echo 2. Excel                               ^|  ^[%Excel%[97m^]
echo 3. Groove (Client for SharePoint)      ^|  ^[%Groove%[97m^]
echo 4. Lync (Skype for Business)           ^|  ^[%Lync%[97m^]
echo 5. OneDrive                            ^|  ^[%OneDrive%[97m^]
echo 6. OneNote                             ^|  ^[%OneNote%[97m^]
echo 7. Outlook                             ^|  ^[%Outlook%[97m^]
echo 8. PowerPoint                          ^|  ^[%PowerPoint%[97m^]
echo 9. Publisher                           ^|  ^[%Publisher%[97m^]
echo 10. Teams                              ^|  ^[%Teams%[97m^]
echo 11. Word                               ^|  ^[%Word%[97m^]
echo[
echo 12. Go back and don't save.
echo 13. Save now.
echo[
set "selection="
set /p selection=
IF "%selection%" == "1" ( CALL :include_and_exclude %Access%, Access, configuration_files_create
) ELSE IF "%selection%" == "2" ( CALL :include_and_exclude %Excel%, Excel, configuration_files_create
) ELSE IF "%selection%" == "3" ( CALL :include_and_exclude %Groove%, Groove, configuration_files_create
) ELSE IF "%selection%" == "4" ( CALL :include_and_exclude %Lync%, Lync, configuration_files_create
) ELSE IF "%selection%" == "5" ( CALL :include_and_exclude %OneDrive%, OneDrive, configuration_files_create
) ELSE IF "%selection%" == "6" ( CALL :include_and_exclude %OneNote%, OneNote, configuration_files_create
) ELSE IF "%selection%" == "7" ( CALL :include_and_exclude %Outlook%, Outlook, configuration_files_create
) ELSE IF "%selection%" == "8" ( CALL :include_and_exclude %PowerPoint%, PowerPoint, configuration_files_create
) ELSE IF "%selection%" == "9" ( CALL :include_and_exclude %Publisher%, Publisher, configuration_files_create
) ELSE IF "%selection%" == "10" ( CALL :include_and_exclude %Teams%, Teams, configuration_files_create
) ELSE IF "%selection%" == "11" ( CALL :include_and_exclude %Word%, Word, configuration_files_create
) ELSE IF "%selection%" == "12" (
    CALL :resetappvars
    cls
    goto configuration_files
) ELSE IF "%selection%" == "13" (
    cls
    goto configuration_files_create_save
) ELSE (
    COLOR 4F
    cls
    echo[ &echo Invalid input! &echo[
    pause
    goto configuration_files_create 
)

:include_and_exclude
IF "%~1%" == "[92mINCLUDED" (
    set %~2=[91mEXCLUDED
) ELSE IF "%~1%" == "[91mEXCLUDED" (
    set %~2=[92mINCLUDED
)
goto %~3


:configuration_files_create_save
color 1F
cls
echo[ &echo Name what you would like your configuration to be ^called without ^(.xml^). &echo You can go back to make changes by pressing space with no input. &echo[
set "selection="
set /p selection=
IF "%selection%" == "" ( goto configuration_files_create)
IF EXIST "%~dp0Data\SavedConfigurations\%selection%.xml" (
    color 4F
    cls
    echo[ &echo Configuration file already exists by that name. &echo[
    pause
    goto configuration_files_create_save )
(echo ^<Configuration^> &echo   ^<Add SourcePath^=^"%~dp0Data\TempInstall\^" OfficeClientEdition^=^"64^" Channel^=^"PerpetualVL2021^"^> &echo     ^<Product ID^=^"ProPlus2021Volume^"^> PIDKEY^=^"BN8D3-W2QKT-M7Q73-Y3JWK-KQC63^"^> &echo        ^<Language ID^=^"MatchOS^" ^/^>
IF "%Access%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Access" ^/^>
IF "%Excel%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Excel" ^/^>
IF "%Groove%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Groove" ^/^>
IF "%Lync%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Lync" ^/^>
IF "%OneDrive%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="OneDrive" ^/^>
IF "%OneNote%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="OneNote" ^/^>
IF "%Outlook%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Outlook" ^/^>
IF "%PowerPoint%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="PowerPoint" ^/^>
IF "%Publisher%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Publisher" ^/^>
IF "%Teams%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Teams" ^/^>
IF "%Word%" == "[91mEXCLUDED" echo        ^<ExcludeApp ID="Word" ^/^>
echo     ^<^/Product^>&echo   ^<^/Add^> &echo   ^<Display Level^=^"None^" AcceptEULA^=^"TRUE^" ^/^> &echo ^<^/Configuration^>) >"%~dp0Data\SavedConfigurations\%selection%.xml"
cls
echo Created configuration file at "%~dp0Data\SavedConfigurations\%selection%.xml".
echo You can now download and install your configuration file at download and install menu.
CALL :resetappvars
pause
goto mainmenu



:resetappvars
set Access=
set Excel=
set Groove=
set Lync=
set OneDrive=
set OneNote=
set Outlook=
set PowerPoint=
set Publisher=
set Teams=
set Word=

:deletedesktopshortcuts
IF EXIST "C:\Users\%USERNAME%\Desktop\Access.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Access.lnk", &echo Deleted Access shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Excel.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Excel.lnk", &echo Deleted Excel shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\OneDrive.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\OneDrive.lnk", &echo Deleted OneDrive shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\OneNote.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\OneNote.lnk", &echo Deleted OneNote shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Outlook.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Outlook.lnk", &echo Deleted Outlook shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Powerpoint.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Powerpoint.lnk", &echo Deleted PowerPoint shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Publisher.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Publisher.lnk", &echo Deleted Publisher shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Skype for Business.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Skype for Business.lnk", &echo Deleted Skype for Business shortcut from desktop.
IF EXIST "C:\Users\%USERNAME%\Desktop\Word.lnk" DEL /F "C:\Users\%USERNAME%\Desktop\Word.lnk", &echo Deleted Word shortcut from desktop.



:deletesearchshortcuts
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk", &echo Deleted Access shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk", &echo Deleted Excel shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk", &echo Deleted OneDrive shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk", &echo Deleted OneNote shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk", &echo Deleted Outlook shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk", &echo Deleted PowerPoint shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk", &echo Deleted Publisher shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Skype for Business.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Skype for Business.lnk", &echo Deleted Skype for Business shortcut from search index.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" DEL /F "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk", &echo Deleted Word shortcut from search index.

:copyappshortcuts
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Access to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Excel to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for OneDrive to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shorcut for OneNote to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Outlook to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PowerPoint.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for PowerPoint to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Publisher to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Skype for Business.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Skype for Business.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Skype for business to the desktop.
IF EXIST "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" xcopy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk" "C:\Users\%USERNAME%\Desktop\*">nul, &echo Added shortcut for Word to the desktop.

:REMOVETEMPFILES
RMDIR /S /Q "%~dp0Data\TempInstall\"
MKDIR "%~dp0Data\TempInstall"

REM 982 lines
