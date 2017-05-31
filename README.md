# VSIXInteropFormsToolkit
Visual Studio 2015 extension, with command tool to generate interop wrapper classes. adapted from add-in.

https://visualstudiogallery.msdn.microsoft.com/31b6a154-3c85-4892-8cea-797460579912

See:

http://www.codeproject.com/Articles/15690/VB-C-Interop-Form-Toolkit

http://www.codeproject.com/Articles/18954/Interop-Forms-Toolkit-Tutorial

https://www.microsoft.com/en-us/download/details.aspx?id=3264

https://interoptoolkitcs.codeplex.com (also imported to https://github.com/froque/interoptoolkitcs)

### Info
At present state this tool does not override files previously generated. 
When forms with unsupported parameter types are found, the initialize method will be generated with a paramenter with the root type "Object".

### How to Debug 

Open project properties 
- Start Action 
	- choose "Start external program"
	- and write path to visual studio. (Ex: "C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\devenv.exe")
- On Command line arguments "/rootsuffix Exp"
- Start application
- Create project in newly opened VS window
- Run tool
