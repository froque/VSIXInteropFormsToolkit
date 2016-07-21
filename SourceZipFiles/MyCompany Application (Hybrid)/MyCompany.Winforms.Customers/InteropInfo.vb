Imports Microsoft.InteropFormTools

' This ComVisible attribute is placed outside of the 
' AssemblyInfo.vb file since the Project Properties
' Designer will change the attribute in that file when
' the "Register For COM Interop" option is checked.
' If the attribute is added to that AssemblyInfo.vb
' it will cause a compile error so simply delete it.
<Assembly: System.Runtime.InteropServices.ComVisible(False)> 

' This property is included in the My namespace, so it will show up in IntelliSense after referencing My.
Namespace My
    ' The HideModuleNameAttribute hides the module name MyInteropToolbox so the syntax becomes My.InteropToolbox.   
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