using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;

using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using System.Windows.Forms;

using System.IO;
using System.CodeDom;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using Microsoft.InteropFormTools;

namespace VSIXInteropFormsToolkit
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerateInteropFormCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("3621b928-eb17-49b3-8ef4-f9381eca5bd2");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateInteropFormCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private GenerateInteropFormCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }

            this._attrTypeForm = typeof(InteropFormAttribute);
            this._attrTypeInitializer = typeof(InteropFormInitializerAttribute);
            this._attrTypeProperty = typeof(InteropFormPropertyAttribute);
            this._attrTypeMethod = typeof(InteropFormMethodAttribute);
            this._attrTypeEvent = typeof(InteropFormEventAttribute);
            this._supportedTypes = null;
            this.EVENT_ARGS_COMMENT = String.Format("{0}{1}{2}{3}{4}",
                new object[] {Resource.EVENT_ARGS_COMMENT1, "\r\n",
                    Resource.EVENT_ARGS_COMMENT2, "\r\n",
                    Resource.EVENT_ARGS_COMMENT3 });
            try
            {
                this.LoadSupportedTypes();
            }
            catch (Exception exception1)
            {
                this.DisplayError(Resource.LoadSupportedTypesErrMsg + "\n" + exception1.ToString());
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerateInteropFormCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new GenerateInteropFormCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            this._applicationObject = Package.GetGlobalService(typeof(DTE)) as DTE2;
            this.CreateInteropFormProxiesForSolution();
        }


        private void CreateInteropFormProxiesForSolution()
        {
            this._applicationObject.StatusBar.Text = Resource.ADDIN_STATUS_GENERATING;
            foreach (Project project1 in this._applicationObject.Solution.Projects)
            {
                if ((project1.ProjectItems != null) && (project1.ProjectItems.Count > 0))
                {
                    this.CreateInteropFormProxiesForProject(project1, project1.ProjectItems);
                }
            }
            this._applicationObject.StatusBar.Text = Resource.ADDIN_STATUS_GENERATED_OK;
        }

        private void CreateInteropFormProxiesForProject(Project currentAssembly, ProjectItems projItemCollection)
        {
            IsVB = (String.Compare(currentAssembly.CodeModel.Language, Resource.LanguageVB, false) == 0);
            foreach (ProjectItem item1 in projItemCollection)
            {
                try
                {
                    if ((String.Compare(item1.Kind, Resource.DOCUMENT_TYPE, false) == 0) && (item1.FileCodeModel != null))
                    {
                        List<CodeClass> list1 = this.GetInteropFormClasses(currentAssembly, item1);
                        this.CreateInteropFormProxiesForDocument(list1, currentAssembly, item1);
                        continue;
                    }
                    if ((item1.ProjectItems != null) && (item1.ProjectItems.Count > 0))
                    {
                        this.CreateInteropFormProxiesForProject(currentAssembly, item1.ProjectItems);
                    }
                    continue;
                }
                catch (Exception exception2)
                {
                    this.DisplayError(String.Format(Resource.ADDIN_STATUS_GENERATED_ERROR_FULL, currentAssembly.Name) + "\n" + exception2.ToString());
                    continue;
                }
            }
        }

        private void CreateInteropFormProxiesForDocument(List<CodeClass> interopFormClasses, Project currentAssembly, ProjectItem interopFormDoc)
        {
            ProjectItem item2 = null;
            if (interopFormClasses.Count <= 0)
                return;

            FileInfo info1 = new FileInfo(interopFormDoc.get_FileNames(0));
            DirectoryInfo info3 = new DirectoryInfo(info1.DirectoryName + @"\" + Resource.INTEROP_FORM_PROXY_FOLDER_NAME);
            foreach (ProjectItem item3 in currentAssembly.ProjectItems)
            {
                if ((String.Compare(item3.Kind, Resource.FOLDER_TYPE, false) == 0) &&
                    (String.Compare(item3.Name, Resource.INTEROP_FORM_PROXY_FOLDER_NAME, false) == 0))
                {
                    item2 = item3;
                    break;
                }
            }

            if (item2 == null)
            {
                if (!info3.Exists)
                {
                    item2 = currentAssembly.ProjectItems.AddFolder(info3.Name, "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}");
                }
                else
                {
                    info3.Delete(true);
                    item2 = currentAssembly.ProjectItems.AddFolder(info3.Name, "{6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C}");
                }
            }
            string text1 = info3.FullName + @"\" + info1.Name.Replace(info1.Extension, ".wrapper" + info1.Extension);
            FileInfo info2 = new FileInfo(text1);
            foreach (ProjectItem item4 in item2.ProjectItems)
            {
                if (String.Compare(item4.Kind, Resource.DOCUMENT_TYPE, false) != 0 || String.Compare(item4.Name, info2.Name, false) != 0)
                    continue; ;

                ProjectItem item1 = item4;
                if (currentAssembly.DTE.SourceControl.IsItemUnderSCC(text1) && !item4.Collection.ContainingProject.DTE.SourceControl.IsItemCheckedOut(text1))
                {
                    item4.Collection.ContainingProject.DTE.SourceControl.CheckOutItem(text1);
                }
                break;
            }

            if (info2.Exists)
            {
                info2.Delete();
            }
            CodeCompileUnit unit1 = new CodeCompileUnit();
            CodeNamespaceImport import1 = new CodeNamespaceImport(this._attrTypeForm.Namespace);
            System.CodeDom.CodeNamespace namespace1 = new System.CodeDom.CodeNamespace();
            namespace1.Name = "Interop";
            unit1.Namespaces.Add(namespace1);
            namespace1.Imports.Add(import1);
            foreach (CodeClass class1 in interopFormClasses)
            {
                string text2 = class1.FullName;
                CodeTypeDeclaration declaration1 = new CodeTypeDeclaration(class1.Name);
                namespace1.Types.Add(declaration1);
                declaration1.IsClass = true;
                declaration1.IsPartial = true;
                CodePrimitiveExpression expression2 = new CodePrimitiveExpression(true);
                CodeSnippetExpression expression1 = new CodeSnippetExpression("System.Runtime.InteropServices.ClassInterfaceType.AutoDual");
                declaration1.CustomAttributes.Add(new CodeAttributeDeclaration("System.Runtime.InteropServices.ClassInterface", new System.CodeDom.CodeAttributeArgument[] { new System.CodeDom.CodeAttributeArgument(expression1) }));
                declaration1.CustomAttributes.Add(new CodeAttributeDeclaration("System.Runtime.InteropServices.ComVisible", new System.CodeDom.CodeAttributeArgument[] { new System.CodeDom.CodeAttributeArgument(expression2) }));
                declaration1.BaseTypes.Add(new CodeTypeReference(typeof(InteropFormProxyBase).Name));
                CodeTypeDeclaration declaration2 = new CodeTypeDeclaration("I" + declaration1.Name + "EventSink");
                declaration2.CustomAttributes.Add(new CodeAttributeDeclaration("System.Runtime.InteropServices.InterfaceTypeAttribute", new System.CodeDom.CodeAttributeArgument[] { new System.CodeDom.CodeAttributeArgument(new CodeSnippetExpression("System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIDispatch")) }));
                declaration2.CustomAttributes.Add(new CodeAttributeDeclaration("System.Runtime.InteropServices.ComVisible", new System.CodeDom.CodeAttributeArgument[] { new System.CodeDom.CodeAttributeArgument(expression2) }));
                declaration2.IsInterface = true;
                CodeConstructor constructor1 = new CodeConstructor();
                CodeSnippetStatement statement1 = GetStatementInitializeForm(text2);
                CodeSnippetStatement statement2 = GetStatementRegisterForm();
                constructor1.Statements.Add(statement1);
                constructor1.Statements.Add(statement2);
                constructor1.Attributes = MemberAttributes.Public;
                declaration1.Members.Add(constructor1);
                if (class1.Members.Count > 0)
                {
                    foreach (CodeElement element1 in class1.Members)
                    {
                        switch (element1.Kind)
                        {
                            case vsCMElement.vsCMElementFunction:
                                CodeFunction2 function1 = (CodeFunction2)element1;
                                if (function1.Access == vsCMAccess.vsCMAccessPublic)
                                {
                                    foreach (CodeElement element2 in function1.Attributes)
                                    {
                                        if (this.AttributesMatch(element2, this._attrTypeInitializer))
                                        {
                                            this.AddInitializeMethodForConstructor(declaration1, class1, function1);
                                            break;
                                        }
                                        if (this.AttributesMatch(element2, this._attrTypeMethod))
                                        {
                                            this.AddMethod(declaration1, class1, function1);
                                            break;
                                        }
                                    }
                                }
                                break;
                            case vsCMElement.vsCMElementProperty:
                                CodeProperty property1 = (CodeProperty)element1;
                                if (property1.Access == vsCMAccess.vsCMAccessPublic)
                                {
                                    foreach (CodeElement element3 in property1.Attributes)
                                    {
                                        if (this.AttributesMatch(element3, this._attrTypeProperty))
                                        {
                                            this.AddProperty(declaration1, class1, property1);
                                            break;
                                        }
                                    }
                                }
                                break;
                            case vsCMElement.vsCMElementEvent:
                                CodeEvent event1 = (CodeEvent)element1;
                                if (event1.Access == vsCMAccess.vsCMAccessPublic)
                                {
                                    foreach (CodeElement element3 in event1.Attributes)
                                    {
                                        if (this.AttributesMatch(element3, this._attrTypeEvent))
                                        {
                                            this.AddEvent(currentAssembly,
                                                declaration1, class1, event1, declaration2);
                                            break;
                                        }
                                    }
                                }
                                break;
                        }
                    }
                }
                if (declaration2.Members.Count > 0)
                {
                    namespace1.Types.Add(declaration2);
                    declaration1.CustomAttributes.Add(new CodeAttributeDeclaration("System.Runtime.InteropServices.ComSourceInterfaces", new System.CodeDom.CodeAttributeArgument[] { new System.CodeDom.CodeAttributeArgument(new CodeTypeOfExpression(declaration2.Name)) }));
                }
            }
            StreamWriter writer1 = new StreamWriter(info2.Create());
            writer1.AutoFlush = true;
            CodeDomProvider provider = GetProvider();
            CodeGeneratorOptions options1 = new CodeGeneratorOptions();
            provider.GenerateCodeFromCompileUnit(unit1, writer1, options1);
            writer1.Close();
            writer1.Dispose();
            item2.ProjectItems.AddFromFile(info2.FullName);
        }

        private void DisplayError(string errorMessage)
        {
            MessageBox.Show(errorMessage, Resource.DISPLAY_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private CodeSnippetStatement GetStatementInitializeForm(string typeName)
        {
            return new CodeSnippetStatement(String.Format(CultureInfo.InvariantCulture, "\t\t\tFormInstance = {0} {1}(){2}", statementNew, typeName, statementTerminator));
        }

        private CodeSnippetStatement GetStatementRegisterForm()
        {
            return new CodeSnippetStatement(String.Format(CultureInfo.InvariantCulture, "\t\t\tRegisterFormInstance(){0}", statementTerminator));
        }


        private bool AttributesMatch(CodeElement ce, System.Type attrType)
        {
            bool flag2 = false;
            string text1 = String.Empty;
            if (ce == null)
                return flag2;

            if (!String.IsNullOrEmpty(ce.Name))
            {
                text1 = ce.Name;
                flag2 = (String.Compare(text1, attrType.Name, true) == 0) ||
                    (String.Compare(text1, attrType.Name.Replace("Attribute", ""), true) == 0);
            }
            if (!flag2 && !String.IsNullOrEmpty(ce.Name))
            {
                text1 = ce.FullName;
                flag2 = (String.Compare(text1, attrType.FullName, true) == 0) ||
                    (String.Compare(text1, attrType.FullName.Replace("Attribute", ""), true) == 0);
            }
            return flag2;
        }


        private void AddInitializeMethodForConstructor(CodeTypeDeclaration proxyClass,
            CodeClass interopFormClass, CodeFunction method)
        {
            CodeMemberMethod method1 = new CodeMemberMethod();
            method1.Name = "Initialize";
            method1.Attributes = MemberAttributes.Public;
            method1.CustomAttributes.Add(new CodeAttributeDeclaration("System.Diagnostics.DebuggerStepThrough"));
            string text1 = String.Format(CultureInfo.InvariantCulture, "\t\t\tFormInstance = {0} {1}(", statementNew, interopFormClass.FullName);
            bool flag1 = false;
            foreach (CodeParameter parameter1 in method.Parameters)
            {
                if (!this.IsSupported(parameter1.Type))
                {
                    this.DisplayWarning(String.Format(Resource.InitMethodErrMsg, parameter1.Type.AsFullName, parameter1.Name, parameter1.Type.AsFullName));
                    return;
                }
                CodeParameterDeclarationExpression expression1 = new CodeParameterDeclarationExpression(parameter1.Type.AsFullName, parameter1.Name);
                method1.Parameters.Add(expression1);
                if (flag1)
                {
                    text1 += ", ";
                }
                text1 += parameter1.Name;
                flag1 = true;
            }
            text1 = text1 + ")" + statementTerminator;
            method1.Statements.Add(new CodeSnippetStatement(String.Format("\t\t\tUnregisterFormInstance(){0}", statementTerminator)));
            method1.Statements.Add(new CodeSnippetStatement(text1));
            method1.Statements.Add(new CodeSnippetStatement(String.Format("\t\t\tRegisterFormInstance(){0}", statementTerminator)));
            proxyClass.Members.Add(method1);
        }


        private void AddEvent(Project currentAssembly,
            CodeTypeDeclaration proxyClass, CodeClass interopFormClass,
            CodeEvent evt, CodeTypeDeclaration proxyClassEventSinkInterface)
        {
            CodeDelegate2 delegate1 = null;
            try
            {
                delegate1 = (CodeDelegate2)currentAssembly.CodeModel.CodeTypeFromFullName(evt.Type.AsFullName);
            }
            catch (Exception exception2)
            {
                foreach (CodeElement element1 in evt.ProjectItem.FileCodeModel.CodeElements)
                {
                    if (element1.IsCodeType)
                    {
                        CodeType type1 = (CodeType)element1;
                        foreach (CodeElement element2 in type1.Children)
                        {
                            if ((element2.Kind == vsCMElement.vsCMElementDelegate) &
                                (String.Compare(element2.FullName, evt.Type.AsFullName, false) == 0))
                            {
                                delegate1 = (CodeDelegate2)element2;
                            }
                        }
                        continue;
                    }
                }
            }

            if (delegate1 == null)
            {
                this.DisplayWarning(string.Format(Resource.EventErrMsg, evt.Name, evt.Type.AsFullName));
            }
            else
            {
                CodeMemberMethod method1 = null;
                foreach (CodeTypeMember member1 in proxyClass.Members)
                {
                    if (String.Compare(member1.Name, "HookCustomEvents", false) == 0)
                    {
                        method1 = (CodeMemberMethod)member1;
                    }
                }
                if (method1 == null)
                {
                    method1 = new CodeMemberMethod();
                    method1.Name = "HookCustomEvents";
                    method1.Attributes = MemberAttributes.Family | MemberAttributes.Override;
                    method1.Statements.Add(this.GetStatementCastFormInstance(interopFormClass));
                    proxyClass.Members.Add(method1);
                }
                CodeMemberEvent event1 = new CodeMemberEvent();
                event1.Attributes = MemberAttributes.Public;
                event1.Type = new CodeTypeReference(evt.Type.AsFullName);
                event1.Name = evt.Name;
                CodeMemberMethod method3 = new CodeMemberMethod();
                method3.Name = evt.Name;
                CodeTypeDelegate delegate2 = new CodeTypeDelegate(evt.Name + "Handler");
                bool flag1 = false;
                CodeMemberMethod method2 = new CodeMemberMethod();
                method2.Name = "castFormInstance_" + evt.Name;
                CodeDelegateInvokeExpression expression1 = new CodeDelegateInvokeExpression(new CodeEventReferenceExpression(new CodeThisReferenceExpression(), event1.Name));
                foreach (CodeParameter parameter1 in delegate1.Parameters)
                {
                    CodeParameterDeclarationExpression expression2;
                    CodeArgumentReferenceExpression expression3;
                    CodeParameterDeclarationExpression expression4;
                    if ((parameter1.Type.CodeType != null) && this.IsEventArgs(parameter1.Type.CodeType))
                    {
                        if (!flag1)
                        {
                            proxyClass.Members.Add(delegate2);
                            event1.Type = new CodeTypeReference(delegate2.Name);
                        }
                        expression4 = new CodeParameterDeclarationExpression("System.EventArgs", parameter1.Name);
                        expression2 = new CodeParameterDeclarationExpression(parameter1.Type.AsFullName, parameter1.Name);
                        expression3 = new CodeArgumentReferenceExpression(expression4.Name);
                        event1.Comments.Add(new CodeCommentStatement(this.EVENT_ARGS_COMMENT));
                        method3.Comments.Add(new CodeCommentStatement(this.EVENT_ARGS_COMMENT));
                    }
                    else
                    {
                        if (!this.IsSupported(parameter1.Type))
                        {
                            this.DisplayWarning(String.Format(Resource.EventErrMsg2, parameter1.Type.AsFullName, evt.Name));
                            return;
                        }
                        expression4 = new CodeParameterDeclarationExpression(parameter1.Type.AsFullName, parameter1.Name);
                        expression2 = new CodeParameterDeclarationExpression(parameter1.Type.AsFullName, parameter1.Name);
                        expression3 = new CodeArgumentReferenceExpression(expression4.Name);
                    }
                    method3.Parameters.Add(expression4);
                    delegate2.Parameters.Add(expression4);
                    method2.Parameters.Add(expression2);
                    expression1.Parameters.Add(expression3);
                }
                method2.Statements.Add(expression1);
                method1.Statements.Add(new CodeAttachEventStatement(new CodeEventReferenceExpression(new CodeVariableReferenceExpression("castFormInstance"), event1.Name), new CodeDelegateCreateExpression(event1.Type, new CodeThisReferenceExpression(), method2.Name)));
                proxyClassEventSinkInterface.Members.Add(method3);
                proxyClass.Members.Add(method2);
                proxyClass.Members.Add(event1);
            }
        }

        private bool IsEventArgs(CodeType parmType)
        {
            if (String.Compare(parmType.FullName, "system.eventargs", true) == 0)
                return true;

            foreach (CodeElement element1 in parmType.Bases)
            {
                if (String.Compare(element1.FullName, "system.eventargs", true) == 0)
                    return true;

                if (element1.IsCodeType && this.IsEventArgs((CodeType)element1))
                    return true;

            }
            return false;
        }

        private CodeSnippetStatement GetStatementCastFormInstance(CodeClass interopFormClass)
        {
            string statementFormat = (IsVB) ?
                "\t\t\tDim castFormInstance As {0} = FormInstance" :
                "\t\t\t{0} castFormInstance = ({0})FormInstance;";

            return new CodeSnippetStatement(String.Format(CultureInfo.InvariantCulture, statementFormat, interopFormClass.FullName));
        }

        private void AddProperty(CodeTypeDeclaration proxyClass,
            CodeClass interopFormClass, CodeProperty prop)
        {
            CodeMemberProperty property1 = new CodeMemberProperty();
            property1.Name = prop.Name;
            property1.Attributes = MemberAttributes.Public;
            property1.Type = new CodeTypeReference(prop.Type.AsFullName);
            if (!this.IsSupported(prop.Type))
            {
                this.DisplayWarning(String.Format(Resource.PropertyErrMsg, prop.Type.AsFullName, property1.Name));
                return;
            }
            if (prop.Getter != null)
            {
                property1.HasGet = true;
                property1.GetStatements.Add(this.GetStatementCastFormInstance(interopFormClass));
                property1.GetStatements.Add(new CodeMethodReturnStatement(new CodePropertyReferenceExpression(new CodeVariableReferenceExpression("castFormInstance"), prop.Name)));
            }
            if (prop.Setter != null)
            {
                property1.HasSet = true;
                property1.SetStatements.Add(this.GetStatementCastFormInstance(interopFormClass));
                property1.SetStatements.Add(new CodeAssignStatement(new CodePropertyReferenceExpression(new CodeVariableReferenceExpression("castFormInstance"), prop.Name), new CodePropertySetValueReferenceExpression()));
            }
            proxyClass.Members.Add(property1);
        }
        private void AddMethod(CodeTypeDeclaration proxyClass,
            CodeClass interopFormClass, CodeFunction method)
        {
            CodeMemberMethod method1 = new CodeMemberMethod();
            method1.Name = method.Name;
            method1.Attributes = MemberAttributes.Public;
            CodePrimitiveExpression expression1 = new CodePrimitiveExpression(true);
            method1.CustomAttributes.Add(new CodeAttributeDeclaration("System.Diagnostics.DebuggerStepThrough"));
            method1.Statements.Add(this.GetStatementCastFormInstance(interopFormClass));
            bool lShouldReturn;
            CodeSnippetStatement methodStatement = GetStatementMethod(method, method1, out lShouldReturn);
            if (lShouldReturn)
                return;
            method1.Statements.Add(methodStatement);
            proxyClass.Members.Add(method1);
        }


        private CodeSnippetStatement GetStatementMethod(CodeFunction method,
            CodeMemberMethod method1, out bool shouldReturn)
        {
            shouldReturn = false;
            string text1;
            if (method.FunctionKind == vsCMFunction.vsCMFunctionFunction &&
                method.Type.TypeKind != vsCMTypeRef.vsCMTypeRefVoid)
            {
                if (!this.IsSupported(method.Type))
                {
                    this.DisplayWarning(String.Format(Resource.MethodErrMsg1, method.Type.AsFullName, method.Name));
                    shouldReturn = true;
                    return null;
                }
                method1.ReturnType = new CodeTypeReference(method.Type.AsFullName);
                text1 = String.Format("\t\t\t{0} ", statementReturn);
            }
            else
            {
                text1 = "\t\t\t";
            }
            text1 += "castFormInstance." + method.Name + "(";
            bool flag1 = false;
            foreach (CodeParameter parameter1 in method.Parameters)
            {
                if (!this.IsSupported(parameter1.Type))
                {
                    this.DisplayWarning(String.Format(Resource.MethodErrMsg2, parameter1.Type.AsFullName, method.Name));
                    shouldReturn = true;
                    return null;
                }
                CodeParameterDeclarationExpression expression2 = new CodeParameterDeclarationExpression(parameter1.Type.AsFullName, parameter1.Name);
                method1.Parameters.Add(expression2);
                if (flag1)
                {
                    text1 += ", ";
                }
                text1 += parameter1.Name;
                flag1 = true;
            }
            text1 = text1 + ")" + statementTerminator;
            CodeSnippetStatement methodStatement = new CodeSnippetStatement(text1);
            return methodStatement;
        }

        private CodeDomProvider GetProvider()
        {
            if (IsVB)
            {
                return new Microsoft.VisualBasic.VBCodeProvider();
            }
            else
            {
                return new CSharpCodeProvider();
            }
        }

        private bool IsSupported(CodeTypeRef typeToCheck)
        {
            foreach (System.Type type1 in this._supportedTypes)
            {
                if (String.Compare(typeToCheck.AsFullName, type1.FullName, false) == 0)
                    return true;

            }
            return false;
        }

        private void DisplayWarning(string errorMessage)
        {
            MessageBox.Show(errorMessage, Resource.DISPLAY_CAPTION, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void LoadSupportedTypes()
        {
            this._supportedTypes = new List<System.Type>();
            this._supportedTypes.Add(typeof(int));
            this._supportedTypes.Add(typeof(string));
            this._supportedTypes.Add(typeof(bool));
            this._supportedTypes.Add(typeof(object));
        }

        public bool IsVB
        {
            get
            {
                return m_isVB;
            }
            set
            {
                if (value)
                {
                    statementTerminator = String.Empty;
                    statementNew = "New";
                    statementReturn = "Return";
                }
                else
                {
                    statementTerminator = ";";
                    statementNew = "new";
                    statementReturn = "return";
                }
                m_isVB = value;
            }
        }


        private void FindInteropFormClasses(Project currentAssembly, CodeElements codeElements, List<CodeClass> interopFormClasses)
        {
            if (codeElements == null)
                return;

            foreach (CodeElement element1 in codeElements)
            {
                if ((element1.Kind == vsCMElement.vsCMElementAttribute) && this.AttributesMatch(element1, this._attrTypeForm))
                {
                    CodeClass class1 = (CodeClass)codeElements.Parent;
                    interopFormClasses.Add(class1);
                }
                if (element1.Children.Count > 0)
                {
                    this.FindInteropFormClasses(currentAssembly, element1.Children, interopFormClasses);
                }
            }
        }
        private List<CodeClass> GetInteropFormClasses(Project assemblyProj, ProjectItem projItem)
        {
            List<CodeClass> list2 = new List<CodeClass>();
            if (projItem.FileCodeModel != null)
            {
                this.FindInteropFormClasses(assemblyProj, projItem.FileCodeModel.CodeElements, list2);
            }
            return list2;
        }

        private DTE2 _applicationObject;
        private System.Type _attrTypeEvent;
        private System.Type _attrTypeForm;
        private System.Type _attrTypeInitializer;
        private System.Type _attrTypeMethod;
        private System.Type _attrTypeProperty;
        private List<System.Type> _supportedTypes;
        private string EVENT_ARGS_COMMENT;

        private bool m_isVB;
        private string statementTerminator;
        private string statementNew;
        private string statementReturn;

    }
}
