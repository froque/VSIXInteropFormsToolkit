:: 64 bits
set PATH=%PATH%;"C:\Program Files (x86)\MSBuild\14.0\Bin\"
set PATH=%PATH%;"C:\Program Files (x86)\Microsoft Visual Studio\VB98\"

:: 32 bits
set PATH=%PATH%;"C:\Program Files\MSBuild\14.0\Bin\"
set PATH=%PATH%;"C:\Program Files\Microsoft Visual Studio\VB98\"

:: Project Hello World
MSBuild.exe /verbosity:minimal /p:Configuration=Release "Sample Applications\Hello World\Hello World Hybrid Application (.NET).sln"
cmd /c VB6.exe /make "Sample Applications\Hello World\HelloWorld\HelloWorld.vbp"
"Sample Applications\Hello World\HelloWorld\HelloWorldVB6.exe"

:: Project Hybrid Application
MSBuild.exe /verbosity:minimal /p:Configuration=Release "Sample Applications\MyCompany Application (Hybrid)\Line Of Business Hybrid Application (.NET).sln"
cmd /c VB6.exe /make "Sample Applications\MyCompany Application (Hybrid)\MyCompany_Components\MyCompany_Components.vbp"
cmd /c VB6.exe /make "Sample Applications\MyCompany Application (Hybrid)\MyCompany_Forms\MyCompany_Forms.vbp"
"Sample Applications\MyCompany Application (Hybrid)\MyCompany_Forms\MyCompany_Forms.exe"

:: Project UserControl
MSBuild.exe /verbosity:minimal /p:Configuration=Release "Sample Applications\UserControl\HybridAppControls\HybridAppControls.sln"
cmd /c VB6.exe /make "Sample Applications\UserControl\HybridAppControls\WordProcessor.vbp"
"Sample Applications\UserControl\HybridAppControls\WordProcessor.exe"
