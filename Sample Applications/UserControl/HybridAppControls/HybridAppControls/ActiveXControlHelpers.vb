Imports Microsoft.InteropFormTools

#If COM_INTEROP_ENABLED Then

'Adds the InteropToolbox to the My namespace
Namespace My
    'The HideModuleNameAttribute hides the module name MyInteropToolbox so the syntax becomes My.InteropToolbox.   
    <Global.Microsoft.VisualBasic.HideModuleName()> _
    Module MyInteropToolbox

        Private _toolbox As New InteropToolbox

        Public ReadOnly Property InteropToolbox() As InteropToolbox
            Get
                Return _toolbox
            End Get
        End Property
    End Module
End Namespace

'Helper routines to do additional registration needed by ActiveX controls.
Friend Module ComRegistration

    Const OLEMISC_RECOMPOSEONRESIZE As Integer = 1
    Const OLEMISC_CANTLINKINSIDE As Integer = 16
    Const OLEMISC_INSIDEOUT As Integer = 128
    Const OLEMISC_ACTIVATEWHENVISIBLE As Integer = 256
    Const OLEMISC_SETCLIENTSITEFIRST As Integer = 131072

    Public Sub RegisterControl(ByVal t As Type)

        Try
            GuardNullType(t, "t")
            GuardTypeIsControl(t)

            'CLSID
            Dim key As String = "CLSID\" & t.GUID.ToString("B")
            Dim ver As String

            Using subkey As RegistryKey = Registry.ClassesRoot.OpenSubKey(key, True)

                'InProcServer32
                Dim InprocKey As RegistryKey = subkey.OpenSubKey("InprocServer32", True)
                If InprocKey IsNot Nothing Then
                    InprocKey.SetValue(Nothing, Environment.SystemDirectory & "\mscoree.dll")
                End If

                'Control
                Using controlKey As RegistryKey = subkey.CreateSubKey("Control")
                End Using

                'Misc
                Using miscKey As RegistryKey = subkey.CreateSubKey("MiscStatus")
                    Dim MiscStatusValue As Integer = OLEMISC_RECOMPOSEONRESIZE + _
                        OLEMISC_CANTLINKINSIDE + OLEMISC_INSIDEOUT + _
                        OLEMISC_ACTIVATEWHENVISIBLE + OLEMISC_SETCLIENTSITEFIRST

                    miscKey.SetValue("", MiscStatusValue.ToString, RegistryValueKind.String)
                End Using

                'ToolBoxBitmap32
                Using bitmapKey As RegistryKey = subkey.CreateSubKey("ToolBoxBitmap32")

                    'If you want to have different icons for each control in this assembly
                    'you can modify this section to specify a different icon each time.
                    'Each specified icon must be embedded as a win32resource in the
                    'assembly; the default one is at index 101, but you can additional ones.
                    bitmapKey.SetValue("", Assembly.GetExecutingAssembly.Location & ", 101", _
                                       RegistryValueKind.String)
                End Using

                'TypeLib
                Using typeLibKey As RegistryKey = subkey.CreateSubKey("TypeLib")
                    Dim libId As Guid = Marshal.GetTypeLibGuidForAssembly(t.Assembly)
                    typeLibKey.SetValue("", libId.ToString("B"), RegistryValueKind.String)
                End Using

                'Version
                Using versionKey As RegistryKey = subkey.CreateSubKey("Version")
                    Dim major, minor As Integer
                    Marshal.GetTypeLibVersionForAssembly(t.Assembly, major, minor)
                    versionKey.SetValue("", String.Format("{0}.{1}", major, minor))
                    ver = String.Format("{0}.{1}", major, minor)
                End Using

            End Using

        Catch ex As Exception
            LogAndRethrowException("ComRegisterFunction failed.", t, ex)
        End Try

    End Sub

    Public Sub UnregisterControl(ByVal t As Type)
        Try
            GuardNullType(t, "t")
            GuardTypeIsControl(t)

            'CLSID
            Dim key As String = "CLSID\" & t.GUID.ToString("B")
            Registry.ClassesRoot.DeleteSubKeyTree(key)

        Catch ex As Exception
            LogAndRethrowException("ComUnregisterFunction failed.", t, ex)
        End Try

    End Sub

    Private Sub GuardNullType(ByVal t As Type, ByVal param As String)
        If t Is Nothing Then
            Throw New ArgumentException("The CLR type must be specified.", param)
        End If
    End Sub

    Private Sub GuardTypeIsControl(ByVal t As Type)
        If Not GetType(Control).IsAssignableFrom(t) Then
            Throw New ArgumentException("Type argument must be a Windows Forms control.")
        End If
    End Sub

    Private Sub LogAndRethrowException(ByVal message As String, ByVal t As Type, ByVal ex As Exception)
        Try
            If t IsNot Nothing Then
                message &= vbCrLf & String.Format("CLR class '{0}'", t.FullName)
            End If

            Throw New ComRegistrationException(message, ex)

        Catch ex2 As Exception
            My.Application.Log.WriteException(ex2)
        End Try

    End Sub

End Module

<Serializable()> _
Public Class ComRegistrationException
    Inherits Exception

    Public Sub New()

    End Sub

    Public Sub New(ByVal message As String, ByVal inner As Exception)
        MyBase.New(message, inner)
    End Sub

End Class

'Helper functions to convert common COM types to their .NET equivalents
<ComVisible(False)> _
Friend Class ActiveXControlHelpers
    Inherits System.Windows.Forms.AxHost

    Friend Sub New()
        MyBase.New(Nothing)
    End Sub

    Friend Shared Shadows Function GetColorFromOleColor(ByVal oleColor As Integer) As Color
        Return AxHost.GetColorFromOleColor(CIntToUInt(oleColor))
    End Function

    Friend Shared Shadows Function GetOleColorFromColor(ByVal color As Color) As Integer
        Return CUIntToInt(AxHost.GetOleColorFromColor(color))
    End Function

    Friend Shared Function CUIntToInt(ByVal uiArg As UInteger) As Integer
        If uiArg <= Integer.MaxValue Then
            Return CInt(uiArg)
        End If
        Return CInt(uiArg - 2 * (CUInt(Integer.MaxValue) + 1))
    End Function

    Friend Shared Function CIntToUInt(ByVal iArg As Integer) As UInteger
        If iArg < 0 Then
            Return CUInt(UInteger.MaxValue + iArg + 1)
        End If
        Return CUInt(iArg)
    End Function

    Private Const KEY_PRESSED As Integer = &H1000
    Private Declare Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal ByValnVirtKey As Integer) As Short

    Private Shared Function CheckForAccessorKey() As Integer
        If My.Computer.Keyboard.AltKeyDown Then

            For i As Integer = Keys.A To Keys.Z
                If (GetKeyState(i) And KEY_PRESSED) <> 0 Then
                    Return i
                End If
            Next
        End If

        Return -1
    End Function

    <ComVisible(False)> _
    Friend Shared Sub HandleFocus(ByVal f As UserControl)
        If My.Computer.Keyboard.AltKeyDown Then
            HandleAccessorKey(f.GetNextControl(Nothing, True), f)
        Else
            'Move to the first control that can receive focus, taking into account
            'the possibility that the user pressed <Shift>+<Tab>, in which case we
            'need to start at the end and work backwards.
            Dim ctl As Control = f.GetNextControl(Nothing, Not My.Computer.Keyboard.ShiftKeyDown)
            While ctl IsNot Nothing
                If ctl.Enabled AndAlso ctl.CanSelect Then
                    ctl.Focus()
                    Exit While
                Else
                    ctl = f.GetNextControl(ctl, Not My.Computer.Keyboard.ShiftKeyDown)
                End If
            End While

        End If
    End Sub

    Private Shared Sub HandleAccessorKey(ByVal sender As Object, ByVal f As UserControl)
        Dim key As Integer = CheckForAccessorKey()
        If key = -1 Then Return

        Dim ctlCurrent As Control = f.GetNextControl(CType(sender, Control), False)

        Do
            ctlCurrent = f.GetNextControl(ctlCurrent, True)
            If ctlCurrent IsNot Nothing AndAlso Control.IsMnemonic(ChrW(key), ctlCurrent.Text) Then

                'VB6 handles conflicts correctly already, so if we handle it also we'll end up 
                'one control past where the focus should be
                If Not KeyConflict(ChrW(key), f) Then

                    'If we land on a label or other non-selectable control then go to the next 
                    'control in the tab order
                    If Not ctlCurrent.CanSelect Then
                        Dim ctlAfterLabel As Control = f.GetNextControl(ctlCurrent, True)
                        If ctlAfterLabel IsNot Nothing AndAlso ctlAfterLabel.CanFocus Then
                            ctlAfterLabel.Focus()
                        End If
                    Else
                        ctlCurrent.Focus()
                    End If
                    Exit Do
                End If
            End If

            'Loop until we hit the end of the tab order
            'If we've hit the end of the tab order we don't want to loop back because the
            'parent form's controls come next in the tab order.
        Loop Until ctlCurrent Is Nothing
    End Sub

    Private Shared Function KeyConflict(ByVal key As Char, ByVal u As UserControl) As Boolean
        Dim flag As Boolean = False

        For Each ctl As Control In u.Controls
            If Control.IsMnemonic(key, ctl.Text) Then
                If flag Then Return True
                flag = True
            End If
        Next
        Return False
    End Function

    'Handles <Tab> and <Shift>+<Tab>
    Friend Shared Sub TabHandler(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyCode = Keys.Tab Then
            Dim ctl As Control = CType(sender, Control)

            Dim userCtl As UserControl = GetParentUserControl(ctl)

            Dim firstCtl As Control = userCtl.GetNextControl(Nothing, True)
            Do Until (firstCtl Is Nothing OrElse firstCtl.CanSelect)
                firstCtl = userCtl.GetNextControl(firstCtl, True)
            Loop

            Dim lastCtl As Control = userCtl.GetNextControl(Nothing, False)
            Do Until (lastCtl Is Nothing OrElse lastCtl.CanSelect)
                lastCtl = userCtl.GetNextControl(lastCtl, False)
            Loop

            If ctl Is lastCtl OrElse ctl Is firstCtl OrElse _
                lastCtl.Contains(ctl) OrElse firstCtl.Contains(ctl) Then

                userCtl.SelectNextControl(CType(sender, Control), lastCtl Is userCtl.ActiveControl, _
                                          True, True, True)
            End If
        End If
    End Sub

    Private Shared Function GetParentUserControl(ByVal ctl As Control) As UserControl
        If ctl Is Nothing Then Return Nothing

        Do Until ctl.Parent Is Nothing
            ctl = ctl.Parent
        Loop
        If ctl IsNot Nothing Then
            Return DirectCast(ctl, UserControl)
        End If

        Return Nothing
    End Function

    Friend Shared Sub WireUpHandlers(ByVal ctl As Control, ByVal ValidationHandler As EventHandler)
        If ctl IsNot Nothing Then
            AddHandler ctl.KeyDown, AddressOf ActiveXControlHelpers.TabHandler
            AddHandler ctl.LostFocus, ValidationHandler

            If ctl.HasChildren Then
                For Each child As Control In ctl.Controls
                    WireUpHandlers(child, ValidationHandler)
                Next
            End If
        End If
    End Sub

End Class

#End If