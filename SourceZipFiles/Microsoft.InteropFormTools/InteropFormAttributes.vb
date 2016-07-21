Imports System.Runtime.InteropServices

''' <summary>
''' Decorating an Interop Form with this attribute 
''' will cause the Generate InteropForm Wrapper Classes Addin
''' to generate an Interop Form wrapper class visible to VB6.
''' </summary>
''' <remarks></remarks>
<ComVisible(False), _
AttributeUsage(AttributeTargets.Class)> _
Public Class InteropFormAttribute
    Inherits Attribute
End Class

''' <summary>
''' Decorating a constructor in an Interop Form with this attribute 
''' will cause the Generate InteropForm Wrapper Classes Addin
''' to generate an Initialize method in the 
''' InteropForm wrapper class that maps back to the constructor.
''' </summary>
''' <remarks></remarks>
<ComVisible(False), _
AttributeUsage(AttributeTargets.Constructor)> _
Public Class InteropFormInitializerAttribute
    Inherits Attribute
End Class


''' <summary>
''' Decorating a method (Sub or Function) in an Interop Form with this attribute 
''' will cause the Generate InteropForm Wrapper Classes Addin
''' to generate an method in the InteropForm wrapper 
''' class that maps back to the method (Sub or Function).
''' </summary>
''' <remarks></remarks>
<ComVisible(False), _
AttributeUsage(AttributeTargets.Method)> _
Public Class InteropFormMethodAttribute
    Inherits Attribute
End Class

''' <summary>
''' Decorating a property in an Interop Form with this attribute 
''' will cause the Generate InteropForm Wrapper Classes Addin
''' to generate a property in the InteropForm wrapper 
''' class that maps back to the property.
''' </summary>
''' <remarks></remarks>
<ComVisible(False), _
AttributeUsage(AttributeTargets.Property)> _
Public Class InteropFormPropertyAttribute
    Inherits Attribute
End Class

''' <summary>
''' Decorating an event in an Interop Formwith this attribute 
''' will cause the Generate InteropForm Wrapper Classes Addin
''' to generate an event in the InteropForm wrapper 
''' class for the event.
''' </summary>
''' <remarks></remarks>
<ComVisible(False), _
AttributeUsage(AttributeTargets.Event)> _
Public Class InteropFormEventAttribute
    Inherits Attribute
End Class