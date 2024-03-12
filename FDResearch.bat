@echo off
setlocal enabledelayedexpansion

set "work_folder=%USERPROFILE%\Downloads\WORK"

cd "%work_folder%" || exit /b

echo Warning: The use of this script is entirely at the end user`s responsibility.
echo The script author bears no liability.
echo We ask all users to carefully review the results of this script before using them.
echo Loading...
ping -n 2 127.0.0.1 > nul


for /d %%F in (*) do (
    set "folder_name=%%F"

    pushd "%%F"

    for %%G in (*) do (
        set "filename=%%~nG"
        set "extension=%%~xG"
        if /i not "!extension!" == ".xlsx" if /i not "!extension!" == ".xls" (
            ren "%%G" "!folder_name!_!filename!!extension!"
        )
    )

    popd
)

echo Done!

