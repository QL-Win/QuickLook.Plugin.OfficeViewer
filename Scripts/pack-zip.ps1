Remove-Item ..\QuickLook.Plugin.HelloWorld.qlplugin -ErrorAction SilentlyContinue

$files = Get-ChildItem -Path ..\bin\Release\ -Exclude *.pdb,*.xml
Compress-Archive $files ..\QuickLook.Plugin.HelloWorld.zip
Move-Item ..\QuickLook.Plugin.HelloWorld.zip ..\QuickLook.Plugin.HelloWorld.qlplugin