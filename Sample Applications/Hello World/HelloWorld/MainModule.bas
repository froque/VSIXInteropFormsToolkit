Attribute VB_Name = "MainModule"
' Entry Point for Application
Sub Main()
    
    ' Instantiate the Toolbox
    Set g_InteropToolbox = New InteropToolbox
    g_InteropToolbox.Initialize
    
    ' Call Initialize method only when first creating the toolbox
    ' This aids in the debugging experience
    g_InteropToolbox.Initialize

    ' Signal Application Startup
    g_InteropToolbox.EventMessenger.RaiseApplicationStartedupEvent
    
    
    ' Do application logic
    Dim f1 As New Form1
    f1.Show vbModal
        
        
    ' Signal Application Shutdown
    g_InteropToolbox.EventMessenger.RaiseApplicationShutdownEvent
    
End Sub
