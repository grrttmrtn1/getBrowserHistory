# getBrowserHistory

loadDLL function

You'll need to either base64 the DLL and add this to the respective variables for the dll or download the dlls to the device it's run on and ensure it's on any machine this is run on. If using base64 the dll will be embeded.

I did not include the base64 because I did not read the sqlite licensing for distribution. 

DLLs can be downloaded from here https://system.data.sqlite.org/index.html/doc/trunk/www/downloads.wiki, make sure to choose your respective .Net framework on the machine.

Easy way to get base64 of DLL

$dllcontent = Get-Content -Path "path to dll" -Encoding Byte
$Base64 = [System.Convert]::ToBase64String($dllcontent)
$Base64 | Out-File "path to a text file"

$sqliteDllBase64 is for System.Data.SQLite.dll
$sqliteinteropDllBase6 is for SQLite.Interop.dll

If you don't want to do this you can replace everything in loadDLL() with

Add-Type -Path "path to System.Data.SQLite.dll"

It seems that this dll assumes that SQLite.Interop.dll is in the same folder as it, so if you're moving files yourself I'd recommend placing them as such.
