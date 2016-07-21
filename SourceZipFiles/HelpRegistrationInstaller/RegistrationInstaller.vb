Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Windows.Forms
Imports IWshRuntimeLibrary
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text

''' <summary>
''' Custom Action Installer class responsible for all aspects of installing and uninstalling the custom Help collection
''' for Interop Forms Toolkit.  
''' 
''' This installer handles registering/unregistering the help collection, creating/deleting shortcuts from the start menu, and 
''' launching Help via DExplore after the installer runs.  
''' 
''' This sample also registers all the .NET Interop 
''' </summary>
''' <remarks></remarks>
Public Class RegistrationInstaller

    Public Const MAX_PATH As Short = 260
    Declare Function SearchPath Lib "KERNEL32" Alias "SearchPathA" ( _
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpPath As String, _
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpFileName As String, _
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpExtension As String, _
        ByVal nBufferLength As Integer, ByVal lpBuffer As StringBuilder, _
        <MarshalAs(UnmanagedType.LPStr)> ByVal lpFilePart As String) As Integer


    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub

    Public Overrides Sub Commit(ByVal savedState As System.Collections.IDictionary)
		MyBase.Commit(savedState)
		LogMsg("In Help Commit")
        CreateShortcut()
        RegisterSamples()
        LaunchHelp()
    End Sub

    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)
		MyBase.Uninstall(savedState)
		LogMsg("In Help Uninstall")
        Me.DeleteShortcut()
        UnRegisterSamples()
    End Sub


    'Return where the sample apps are.
    Private Function GetBaseFolderPath() As String
        Dim execPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        Dim execDirectory As String = New System.IO.FileInfo(execPath).DirectoryName
        Return New DirectoryInfo(execDirectory & "..\..\Sample Applications").FullName
    End Function

    'Search and find where the appropriate file is located.
    Private Function FindFile(ByVal strFile As String) As String

        Dim strStartPath As String = GetBaseFolderPath()
        Dim col As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(strStartPath, FileIO.SearchOption.SearchAllSubDirectories, strFile)
        If col.Count > 0 Then
            Return col(0)
        Else
            Return ""
        End If

    End Function

    Private Sub RegisterSamples()
        RegisterFile("MyCompany.HelloWorld.Dll")
        RegisterFile("MyCompany.Customers.Dll")
        RegisterFile("MyCompany.Orders.Dll")
        RegisterFile("HybridAppControls.Dll")
    End Sub

    Private Sub UnRegisterSamples()
        UnRegisterFile("MyCompany.HelloWorld.Dll")
        UnRegisterFile("MyCompany.Customers.Dll")
        UnRegisterFile("MyCompany.Orders.Dll")
        UnRegisterFile("HybridAppControls.Dll")
    End Sub

    Private Sub RegisterFile(ByVal strFile As String)
        Try

            Dim proc As Process
            Dim output, strFilePath As String

            proc = New Process()
            ' Redirect the output stream of the child process.
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True

            'Get exactly where the file is at
			strFilePath = FindFile(strFile)

            If strFilePath = "" Then Throw New FileNotFoundException(strFile)

            'These arguments will register it properly
            proc.StartInfo.Arguments = Chr(34) & strFilePath & Chr(34) & " /codebase /tlb"

            'Figure out where regasm.exe is
            strFilePath = GetRegASMPath()
            If strFilePath = "" Then Throw New FileNotFoundException("regasm.exe")

			proc.StartInfo.FileName = strFilePath
			LogMsg("Registering: " & strFilePath & " " & proc.StartInfo.Arguments)
            proc.StartInfo.WorkingDirectory = GetBaseFolderPath()
            proc.StartInfo.CreateNoWindow = True
            proc.Start()
            output = proc.StandardOutput.ReadToEnd()
            output += proc.StandardError.ReadToEnd()
            proc.WaitForExit()

			LogMsg("Registration result was:" & proc.ExitCode & " " & vbCrLf & output & vbCrLf)

		Catch ex As Exception
			LogMsg(String.Format(My.Resources.RegisterSampleErrMsg, ex.Message), True)
        End Try
    End Sub

    Private Sub UnRegisterFile(ByVal strFile As String)
        Try

            Dim proc As Process
            Dim output, strFilePath, strParentDirectory As String

            proc = New Process()
            ' Redirect the output stream of the child process.
            proc.StartInfo.UseShellExecute = False
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.RedirectStandardError = True

            'Get exactly where the file is at
            strFilePath = FindFile(strFile)
            If strFilePath = "" Then Throw New FileNotFoundException(strFile)
            strParentDirectory = New IO.FileInfo(strFilePath).DirectoryName

            'These arguments will register it properly
            proc.StartInfo.Arguments = Chr(34) & strFilePath & Chr(34) & " /unregister"

            'Figure out where regasm.exe is
            strFilePath = GetRegASMPath()
            If strFilePath = "" Then Throw New FileNotFoundException("regasm.exe")

            proc.StartInfo.FileName = strFilePath
            proc.StartInfo.WorkingDirectory = GetBaseFolderPath()
            proc.StartInfo.CreateNoWindow = True
            proc.Start()
            output = proc.StandardOutput.ReadToEnd()
            output += proc.StandardError.ReadToEnd()
            proc.WaitForExit()

			LogMsg("UnRegister file: " & strFile & ":" & proc.ExitCode & " " & vbCrLf & output & vbCrLf)

            'Delete the .tlb that was created
            Kill(Path.Combine(strParentDirectory, "*.tlb"))

        Catch ex As Exception
			LogMsg(String.Format(My.Resources.UnRegisterSampleErrMsg, ex.Message), True)
        End Try
    End Sub


    'Reasonable way to find out where regasm.exe is located locally.
    Private Function GetRegASMPath() As String

        'We'll first see if regasm.exe is on the path
        Dim strPath As New StringBuilder(MAX_PATH)
        Dim strFilePath As String
        Dim intRet As Integer = SearchPath(vbNullString, "regasm.exe", vbNullString, strPath.Capacity + 1, strPath, vbNullString)
        If intRet <> 0 Then
            strFilePath = Left(strPath.ToString(), intRet)
            Return strFilePath
            Exit Function
        End If

        'Otherwise, we'll find it from the install root in the registry.
        Dim strInstallRoot As String = Registry.LocalMachine.OpenSubKey("Software\Microsoft\.NETFramework").GetValue("InstallRoot")
        strFilePath = Path.Combine(strInstallRoot, "v2.0.50727\regasm.exe")
        'If we find it, this should work.
        If System.IO.File.Exists(strFilePath) Then Return strFilePath

        'Otherwise, find all other regasm.exe binaries and pick one that looks reasonable
        Dim colOthers As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(strInstallRoot, FileIO.SearchOption.SearchAllSubDirectories, "regasm.exe")
        If colOthers.Count > 0 Then
            For Each strFilePath In colOthers
                If Not strFilePath.ToUpper.Contains("V1") Then  'Exclude any V1 versions since those won't work
                    Return strFilePath
                End If
            Next
        End If

        Return ""

    End Function

    Private Sub CreateShortcut()
        Try
            Dim shell As New WshShell
            Dim path As String
            Dim shortcut As IWshShortcut

            ' Set Location of .lnk file.
            'e.g. "C:\Documents and Settings\All Users\Start Menu\Programs\Microsoft Visual Basic Power Packs\Interop Forms Toolkit 2.1\Microsoft Interop Forms Toolkit 2.1 Documentation.lnk"
			path = getHelpLinkPath()
            LogMsg("destination link file was: " & path)

            ' Creates shortcut(.lnk) file.
            shortcut = shell.CreateShortcut(path)

            'This should be the target installed path.
            Dim linkSource As String = IO.Path.Combine(Me.Context.Parameters("targ"), "Documentation/Table of Contents.html")
            LogMsg("Initial help file is " & linkSource)

            shortcut.TargetPath = linkSource

            shortcut.Arguments = String.Empty
            shortcut.Save()
        Catch ex As Exception
			LogMsg(String.Format(My.Resources.CreateShortcutErrMsg, ex.Message), True)
        End Try

    End Sub

    Private Function getHelpLinkPath() As String

        Dim s As System.Collections.ObjectModel.ReadOnlyCollection(Of String)

        Dim shell As New IWshRuntimeLibrary.WshShell
        'Figure out where to start from
        Dim basePath As String = CStr(shell.SpecialFolders.Item("AllUsersStartMenu"))

        'Now search for our app directory.  By the time this code runs, this directory should be created.    
        'We do this since this path is localized on non ENU systems.
        s = My.Computer.FileSystem.GetDirectories(basePath, FileIO.SearchOption.SearchAllSubDirectories, My.Resources.ShortProductName)

        If s.Count = 0 Then
            Return IO.Path.Combine(My.Resources.StartMenuRelPathDefault, My.Resources.StartMenuLink)
        Else
            Return IO.Path.Combine(s(0), My.Resources.StartMenuLink)
        End If

    End Function

    Private Sub DeleteShortcut()
        Try
            ' Delete Start Menu Help File
            Dim shell As New WshShell
            Dim docsShortcutPath As String = ""

            docsShortcutPath = getHelpLinkPath()

            If My.Computer.FileSystem.FileExists(docsShortcutPath) Then
                My.Computer.FileSystem.DeleteFile(docsShortcutPath)
            End If

        Catch ex As Exception
			LogMsg(String.Format(My.Resources.DeleteShortcutErrMsg, ex.Message), True)
        End Try
    End Sub

    Private Sub LaunchHelp()
        Try
            Process.Start(getHelpLinkPath())
        Catch ex As Exception
			LogMsg(String.Format(My.Resources.LaunchHelpErrMsg, ex.Message), True)
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
