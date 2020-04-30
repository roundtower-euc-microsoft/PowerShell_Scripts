$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
Import-Module $EXOModulePath
Connect-EXOPSSession