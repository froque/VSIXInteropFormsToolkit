Option Explicit On
Option Strict On

'This assembly contains custom action code code to install a bootstrapper package + it also 
'contains the code that was in the "addinregistration" custom action in the 1.0 release
'This code has been additionally tweaked so that it works both with VS2005 and Orcas (VS2008)

Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Xml
Imports System.io
Imports System.IO.Path

''' <summary>
''' Custom action designed to install a specific bootstrapper package.
''' </summary>
''' <remarks></remarks>
Public Class BootstrapperPackageInstaller

    Private Const BOOTSTRAPPER_PACKAGE_NAME As String = "Microsoft InteropForms Toolkit"

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    'This logic works for VS2005 only.
    Function GetBootStrapperPackageDirForWhidbey() As String

        Dim objKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\.NETFramework")
        Using objKey

            If objKey Is Nothing Then
                Throw New Exception(My.Resources.BootstrapperDirectoryError)
            End If

            Return IO.Path.Combine(CStr(objKey.GetValue("sdkInstallRootv2.0")), "BootStrapper\Packages\")

        End Using

    End Function

    'This logic works for Orcas only.
    Function GetBootStrapperPackageDirForOrcas() As String

        Dim objKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\GenericBootstrapper\3.5")
        Using objKey

            If objKey Is Nothing Then
                Throw New Exception(My.Resources.BootstrapperDirectoryError)
            End If

            Return IO.Path.Combine(CStr(objKey.GetValue("Path")), "Packages\")

        End Using

	End Function

	'This logic works for Dev10 only.
	Function GetBootStrapperPackageDirForDev10() As String

		Dim objKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\GenericBootstrapper\4.0")
		Using objKey

			If objKey Is Nothing Then
				Throw New Exception(My.Resources.BootstrapperDirectoryError)
			End If

			Return IO.Path.Combine(CStr(objKey.GetValue("Path")), "Packages\")

		End Using

	End Function

    ''' <summary>
    ''' Used to figure out if VS is installed or not.  Use HKLM to do this correctly.
    ''' </summary>
    ''' <param name="ver"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsVSInstalled(ByVal ver As Integer) As Boolean

        Dim objKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software\Microsoft\VisualStudio\" & ver & ".0")

        Using objKey

            If objKey Is Nothing Then Return False
            Return CStr(objKey.GetValue("InstallDir")) <> ""

        End Using

    End Function

    ''' <summary>
    ''' Return the location of VS documents folder for the user
    ''' </summary>
    ''' <param name="ver"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetVSDocumentLocation(ByVal ver As Integer) As String

        Dim sTmp As String = ""
        sTmp = GetVSRegKey("VisualStudioLocation", ver)
        If sTmp <> "" Then Return sTmp 'This works 95% of the time

        'Otherwise, this logic will work in nearly all other cases.
        sTmp = My.Computer.FileSystem.SpecialDirectories.MyDocuments

		LogMsg("My Document location is " & sTmp)
		Dim col As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing

        'Search for an appropriate directory for our version
        If ver = 8 Then
            col = My.Computer.FileSystem.GetDirectories(sTmp, FileIO.SearchOption.SearchTopLevelOnly, "*Visual*2005")
		ElseIf ver = 9 Then
			col = My.Computer.FileSystem.GetDirectories(sTmp, FileIO.SearchOption.SearchTopLevelOnly, "*Visual*2008")
		ElseIf ver = 10 Then
			col = My.Computer.FileSystem.GetDirectories(sTmp, FileIO.SearchOption.SearchTopLevelOnly, "*Visual*2010")
		End If

        If col.Count = 0 Then            
            'If we get to this stage, it means that VS probably has never been launched.   In this case, we'll use the default settings
            If ver = 8 Then
                sTmp = Path.Combine(sTmp, "Visual Studio 2005")
            ElseIf ver = 9 Then
                sTmp = Path.Combine(sTmp, "Visual Studio 2008")
            ElseIf ver = 10 Then
                sTmp = Path.Combine(sTmp, "Visual Studio 2010")
            End If
            LogMsg("VS has never been launched.   Using default path of " & sTmp)
            Return sTmp
        Else
            Return col.Item(0)
        End If

    End Function


    ''' <summary>
    ''' Return standard VS installed regkey
    ''' </summary>
    ''' <param name="keyname"></param>
    ''' <param name="ver">Specify either 8 or 9 here</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetVSRegKey(ByVal keyname As String, ByVal ver As Integer) As String

        Dim objKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\VisualStudio\" & ver & ".0")

        Using objKey

            If objKey Is Nothing Then Return ""
            Return CStr(objKey.GetValue(keyname))

        End Using

    End Function

    'Figure out if Whidbey or Orcas are installed and then install files for them.
    Private Sub HandleAllUserFiles(ByVal blnInstall As Boolean)

		If IsVSInstalled(8) Then
			LogMsg("Whidbey installed")
			If blnInstall Then
				InstallUserFilesforWhidbey()
			Else
				UnInstallUserFilesforWhidbey()
			End If
		End If

		If IsVSInstalled(9) Then
			LogMsg("Orcas installed")
			If blnInstall Then
				InstallUserFilesforOrcas()
			Else
				UnInstallUserFilesforOrcas()
			End If
		End If

		If IsVSInstalled(10) Then
			LogMsg("Dev10 installed")
			If blnInstall Then
				InstallUserFilesforDev10()
			Else
				UnInstallUserFilesforDev10()
			End If
		End If
    End Sub

    'Figure out if Whidbey or Orcas are installed and then install files for them.
    Private Sub HandleBootstrapperFiles(ByVal blnInstall As Boolean)

        'Do appropriate action for Whidbey
        If IsVSInstalled(8) Then

			Dim sDestPath As String = GetBootStrapperPackageDirForWhidbey()
			LogMsg("Bootstrapper path for Whidbey is " & sDestPath)
            If blnInstall Then
                InstallBootStrapperFiles(sDestPath)
            Else
                RemoveBootStrapperFiles(sDestPath)
            End If

        End If

        'Do appropriate action for Orcas
        If IsVSInstalled(9) Then

			Dim sDestPath As String = GetBootStrapperPackageDirForOrcas()
			LogMsg("Bootstrapper path for Orcas is " & sDestPath)
            If blnInstall Then
                InstallBootStrapperFiles(sDestPath)
            Else
                RemoveBootStrapperFiles(sDestPath)
            End If

		End If

		'Do appropriate action for Dev10
		If IsVSInstalled(10) Then

			Dim sDestPath As String = GetBootStrapperPackageDirForDev10()
			LogMsg("Bootstrapper path for Dev10 is " & sDestPath)

			If blnInstall Then
				InstallBootStrapperFiles(sDestPath)
			Else
				RemoveBootStrapperFiles(sDestPath)
			End If

		End If
    End Sub

    Private Sub InstallUserFilesforWhidbey()
        InstallUserFiles(8, True)
    End Sub

	Private Sub UnInstallUserFilesforWhidbey()
		InstallUserFiles(8, False)
	End Sub

	Private Sub InstallUserFilesforOrcas()
		InstallUserFiles(9, True)
	End Sub

	Private Sub UnInstallUserFilesforOrcas()
		InstallUserFiles(9, False)
	End Sub


	Private Sub InstallUserFilesforDev10()
		InstallUserFiles(10, True)
	End Sub

	Private Sub UnInstallUserFilesforDev10()
		InstallUserFiles(10, False)
	End Sub

    'Either copy or delete a file
    Private Sub ProcessFile(ByVal src As String, ByVal dest As String, ByVal blnInstall As Boolean)

        Try
			If blnInstall Then
				LogMsg("Copying '" & src & "' to '" & dest & "'")
				My.Computer.FileSystem.CopyFile(src, dest, True)
			Else
				LogMsg("Deleting '" & dest & "'")
				If File.Exists(dest) Then
					File.Delete(dest)
				End If
			End If

        Catch ex As Exception

            If blnInstall Then
				LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageInstallError, vbNewLine, ex.Message), True)
            Else
				LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageUninstallError, vbNewLine, ex.Message), True)
            End If

        End Try

    End Sub

    Private Sub RemoveBootStrapperFiles(ByVal sDestPath As String)

        Dim sBaseTargetPath As String = sDestPath

        If Not IO.Directory.Exists(sBaseTargetPath) Then
            Exit Sub
        End If

        'Delete entire directory
        If IO.Directory.Exists(sBaseTargetPath & "\" & BOOTSTRAPPER_PACKAGE_NAME) Then
            My.Computer.FileSystem.DeleteDirectory(sBaseTargetPath & "\" & BOOTSTRAPPER_PACKAGE_NAME, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

    End Sub

    'Copy over the bootstrapper files to the correct location
    Private Sub InstallBootStrapperFiles(ByVal sDestPath As String)

        'We'll expect our sources to be here...
        Dim sBaseSourcePath As String = IO.Path.Combine(Me.Context.Parameters("targ"), "Setup\Redist")

        'Ensure our sources exist.
        Dim sBaseTargetPath As String = sDestPath
        If Not IO.Directory.Exists(sBaseTargetPath) Then
            Throw New Exception(sBaseTargetPath & " " & My.Resources.NotFoundError)
        End If

        'Create our dest directory
        sBaseTargetPath = IO.Path.Combine(sBaseTargetPath, BOOTSTRAPPER_PACKAGE_NAME)
        If Not IO.Directory.Exists(sBaseTargetPath) Then
            IO.Directory.CreateDirectory(sBaseTargetPath)
        End If

        'Now copy all the appropriate files
        My.Computer.FileSystem.CopyFile(sBaseSourcePath & "\Microsoft.InteropFormsRedist.msi", sBaseTargetPath & "\Microsoft.InteropFormsRedist.msi", True)
        My.Computer.FileSystem.CopyFile(sBaseSourcePath & "\product.xml", sBaseTargetPath & "\product.xml", True)

        If Not IO.Directory.Exists(sBaseTargetPath & "\" & "en") Then
            IO.Directory.CreateDirectory(sBaseTargetPath & "\" & "en")
        End If

        My.Computer.FileSystem.CopyFile(sBaseSourcePath & "\EULAredist.txt", sBaseTargetPath & "\en\" & "EULAredist.txt", True)
        My.Computer.FileSystem.CopyFile(sBaseSourcePath & "\package.xml", sBaseTargetPath & "\en\" & "package.xml", True)

    End Sub

    'Since we want the toolkit to work with Orcas and Whidbey, we need a custom action to add/delete all the files
    'as appropriate based on what is installed locally.
    Private Sub InstallUserFiles(ByVal ver As Integer, ByVal blnInstall As Boolean)

        Dim sSrc As String

        'This should be the target installed path.
		Dim sBaseSourcePath As String = IO.Path.Combine(Me.Context.Parameters("targ"), "Setup\User")

		LogMsg("Install Source Path is " & sBaseSourcePath)

        'Now, start copying the files over
		Dim sPath As String = GetVSDocumentLocation(ver)

		LogMsg("VSDocumentLocation is " & sPath)

        'Copy addin file.  Actually involves just modifying the .addin file and saving in target path
        sSrc = Combine(sBaseSourcePath, "InteropFormProxyGenerator.AddIn")
        Dim sDest As String = sPath & "\Addins\" & "InteropFormProxyGenerator.AddIn"

        Dim xmlDoc As New XmlDocument()
        Dim asmXmlEl As XmlElement = Nothing
        xmlDoc.Load(sSrc)

        Dim addinAssemblyPath As String = IO.Path.Combine(Me.Context.Parameters("targ"), "AddIn")
        asmXmlEl = xmlDoc.DocumentElement("Addin")("Assembly")
        asmXmlEl.InnerText = asmXmlEl.InnerText.Replace("[ADDIN_ASSEMBLY_PATH]", addinAssemblyPath)
        xmlDoc.Save(sSrc)

        ProcessFile(sSrc, sDest, blnInstall)
        If Not blnInstall Then
            ProcessFile(sDest, sSrc, blnInstall)        'Will uninstall the orig addin file.
        End If

        'Copy snippets
        sSrc = Combine(sBaseSourcePath, "InteropEvent.snippet")
        sDest = sPath & "\Code Snippets\Visual Basic\My Code Snippets\" & "InteropEvent.snippet"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "InteropFunc.snippet")
        sDest = sPath & "\Code Snippets\Visual Basic\My Code Snippets\" & "InteropFunc.snippet"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "InteropInit.snippet")
        sDest = sPath & "\Code Snippets\Visual Basic\My Code Snippets\" & "InteropInit.snippet"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "InteropProp.snippet")
        sDest = sPath & "\Code Snippets\Visual Basic\My Code Snippets\" & "InteropProp.snippet"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "InteropSub.snippet")
        sDest = sPath & "\Code Snippets\Visual Basic\My Code Snippets\" & "InteropSub.snippet"
        ProcessFile(sSrc, sDest, blnInstall)

        'Copy item templates
        sPath = GetVSRegKey("UserItemTemplatesLocation", ver)
        If sPath = "" Then      'On orcas, this isn't always explicitly set.  We'll use the default
            sPath = Combine(GetVSDocumentLocation(ver), "Templates\ItemTemplates")
        End If

        sSrc = Combine(sBaseSourcePath, "VB6 UserControl Item.zip")
        sDest = sPath & "\Visual Basic\" & "VB6 UserControl Item.zip"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "VB6 InteropForm Library Info.zip")
        sDest = sPath & "\Visual Basic\" & "VB6 InteropForm Library Info.zip"
        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "VB6 InteropForm.zip")
        sDest = sPath & "\Visual Basic\" & "VB6 InteropForm.zip"
        ProcessFile(sSrc, sDest, blnInstall)

        'Copy project templates
        sPath = GetVSRegKey("UserProjectTemplatesLocation", ver)
        If sPath = "" Then      'On orcas, this isn't always explicitly set.  We'll use the default
            sPath = Combine(GetVSDocumentLocation(ver), "Templates\ProjectTemplates")
        End If

        sSrc = Combine(sBaseSourcePath, "VB6 UserControl.zip")
        sDest = sPath & "\Visual Basic\Windows\" & "VB6 UserControl.zip"

        ProcessFile(sSrc, sDest, blnInstall)

        sSrc = Combine(sBaseSourcePath, "VB6 InteropForm Library.zip")
        sDest = sPath & "\Visual Basic\Windows\" & "VB6 InteropForm Library.zip"
        ProcessFile(sSrc, sDest, blnInstall)

    End Sub

    ''' <summary>
    ''' Writes out the needed files for an EN bootstrapper package
    ''' </summary>
    ''' <param name="savedState"></param>
    ''' <remarks></remarks>
    Public Overrides Sub Commit(ByVal savedState As System.Collections.IDictionary)
		MyBase.Commit(savedState)

		LogMsg("In Bootstrapper Commit")
        Try

            'Copy over and install all user files
            HandleAllUserFiles(True)

        Catch ex As Exception
			LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageAddInError, vbNewLine, ex.Message), True)
        End Try

        Try

            HandleBootstrapperFiles(True)

        Catch ex As Exception
			LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageInstallError, vbNewLine, ex.Message), True)
        End Try

    End Sub


    Public Overrides Sub Rollback(ByVal savedState As System.Collections.IDictionary)
		MyBase.Rollback(savedState)
        UninstallIfNecessary()
    End Sub

    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
		MyBase.Uninstall(savedState)
		LogMsg("In Bootstrapper Uninstall")
        UninstallIfNecessary()
    End Sub

    'Basically, delete the bootstrapper package we installed.
    Private Sub UninstallIfNecessary()

        Try

            'Uninstall all user files
            HandleAllUserFiles(False)

        Catch ex As Exception
			LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageAddinUninstallError, vbNewLine, ex.Message), True)
        End Try

        Try

            'Uninstall Bootstrapper files
            HandleBootstrapperFiles(False)

        Catch ex As Exception
			LogMsg(String.Format("{0}{1}{2}", My.Resources.PackageUninstallError, vbNewLine, ex.Message), True)
        End Try

	End Sub

	''' <summary>
	''' Helper logging utility
	''' </summary>
	''' <param name="msg"></param>
	''' <remarks></remarks>
	Private Sub LogMsg(ByVal msg As String, Optional ByVal displayDialog As Boolean = False)
		Dim strPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath, "interoptoolkitinstall.log")
		My.Computer.FileSystem.WriteAllText(strPath, "LOGMSG: " & msg & vbCrLf, True)
		If displayDialog Then MsgBox(msg, MsgBoxStyle.Exclamation)
	End Sub

End Class
