﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Public Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Public ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Public Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to this Sample.
        '''</summary>
        Public ReadOnly Property AboutButton_ToolTip_Description() As String
            Get
                Return ResourceManager.GetString("AboutButton_ToolTip_Description", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Recent Files.
        '''</summary>
        Public ReadOnly Property ApplicationMenu_ResentFiles() As String
            Get
                Return ResourceManager.GetString("ApplicationMenu_ResentFiles", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Chart Tools.
        '''</summary>
        Public ReadOnly Property ContextualTabGroup_ChartTools() As String
            Get
                Return ResourceManager.GetString("ContextualTabGroup_ChartTools", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Frame Tools.
        '''</summary>
        Public ReadOnly Property ContextualTabGroup_FrameTools() As String
            Get
                Return ResourceManager.GetString("ContextualTabGroup_FrameTools", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Table Tools.
        '''</summary>
        Public ReadOnly Property ContextualTabGroup_TableTools() As String
            Get
                Return ResourceManager.GetString("ContextualTabGroup_TableTools", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to TX Text Control Words.
        '''</summary>
        Public ReadOnly Property MainWindow_Caption_Product() As String
            Get
                Return ResourceManager.GetString("MainWindow_Caption_Product", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to [untitled].
        '''</summary>
        Public ReadOnly Property MainWindow_Caption_Untitled() As String
            Get
                Return ResourceManager.GetString("MainWindow_Caption_Untitled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to The following exception is thrown:
        '''	
        '''&quot;{0}&quot;
        '''
        '''To get help, please visit the Text Control Support at www.textcontrol.com..
        '''</summary>
        Public ReadOnly Property MessageBox_Application_ThreadException_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_Application_ThreadException_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Change Form Layout.
        '''</summary>
        Public ReadOnly Property MessageBox_ChangeFormLayout_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_ChangeFormLayout_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to To change the program&apos;s view to a {0} to {1} text appearance, the application must be restarted. 
        '''Click OK to restart or Cancel to not restart the program..
        '''</summary>
        Public ReadOnly Property MessageBox_ChangeFormLayout_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_ChangeFormLayout_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Open Hyperlink.
        '''</summary>
        Public ReadOnly Property MessageBox_OpenHyperlink_CouldNotOpenLink_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenHyperlink_CouldNotOpenLink_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Could not open link to &apos;{0}&apos;..
        '''</summary>
        Public ReadOnly Property MessageBox_OpenHyperlink_CouldNotOpenLink_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenHyperlink_CouldNotOpenLink_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Open linked File.
        '''</summary>
        Public ReadOnly Property MessageBox_OpenHyperlink_FileDoesNotExist_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenHyperlink_FileDoesNotExist_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to The linked file &apos;{0}&apos; does not exist..
        '''</summary>
        Public ReadOnly Property MessageBox_OpenHyperlink_FileDoesNotExist_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenHyperlink_FileDoesNotExist_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Open Recent File.
        '''</summary>
        Public ReadOnly Property MessageBox_OpenRecentFile_FileDoesNotExist_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenRecentFile_FileDoesNotExist_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to The selected file does not exist. Do you want to remove it from the list of recent files?.
        '''</summary>
        Public ReadOnly Property MessageBox_OpenRecentFile_FileDoesNotExist_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_OpenRecentFile_FileDoesNotExist_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Incorrect Password.
        '''</summary>
        Public ReadOnly Property MessageBox_PasswordDialog_IncorrectPassword_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_PasswordDialog_IncorrectPassword_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to The entered password is incorrect..
        '''</summary>
        Public ReadOnly Property MessageBox_PasswordDialog_IncorrectPassword_Text() As String
            Get
                Return ResourceManager.GetString("MessageBox_PasswordDialog_IncorrectPassword_Text", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to left.
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentBeforeRestart_Left() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentBeforeRestart_Left", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to right.
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentBeforeRestart_Right() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentBeforeRestart_Right", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to To change the program&apos;s view to a {0} to {1} text appearance, the application must be restarted. Also, changes to &apos;{2}&apos; were not saved.
        '''Click Yes to save these changes and restart, No to restart without saving, or Cancel to not restart..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentBeforeRestart_ToFile() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentBeforeRestart_ToFile", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to To change the program&apos;s view to a {0} to {1} text appearance, the application must be restarted. Also, changes to the document were not saved.
        '''Click Yes to save these changes and restart, No to restart without saving, or Cancel to not restart..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentBeforeRestart_Untitled() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentBeforeRestart_Untitled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Exit.
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnExit_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnExit_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save changes to &apos;{0}&apos; before the program closes? 
        '''Click Yes to save and close, No to close without saving, or Cancel to not close..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnExit_ToFile() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnExit_ToFile", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save the document before the program closes? 
        '''Click Yes to save and close, No to close without saving, or Cancel to not close..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnExit_Untitled() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnExit_Untitled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to New.
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnNew_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnNew_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save changes to &apos;{0}&apos; before a new document is created? 
        '''Click Yes to save and create, No to create without saving, or Cancel to not to create a new document..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnNew_ToFile() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnNew_ToFile", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save the document before a new document is created? 
        '''Click Yes to save and create, No to create without saving, or Cancel to not to create a new document..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnNew_Untitled() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnNew_Untitled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Open.
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnOpen_Caption() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnOpen_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save changes to &apos;{0}&apos; before the new document is opened? 
        '''Click Yes to save and open, No to open without saving, or Cancel to not to open the new document..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnOpen_ToFile() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnOpen_ToFile", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Do you want to save the document before the new document is opened? 
        '''Click Yes to save and open, No to open without saving, or Cancel to not to open the new document..
        '''</summary>
        Public ReadOnly Property MessageBox_SaveDirtyDocumentOnOpen_Untitled() As String
            Get
                Return ResourceManager.GetString("MessageBox_SaveDirtyDocumentOnOpen_Untitled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Cancel.
        '''</summary>
        Public ReadOnly Property PasswordDialog_Cancel() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_Cancel", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Password.
        '''</summary>
        Public ReadOnly Property PasswordDialog_Caption() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_Caption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Enter a password to open the document..
        '''</summary>
        Public ReadOnly Property PasswordDialog_EnterPassword() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_EnterPassword", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to OK.
        '''</summary>
        Public ReadOnly Property PasswordDialog_OK() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_OK", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to _Password:.
        '''</summary>
        Public ReadOnly Property PasswordDialog_Password() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_Password", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to The document &apos;{0}&apos; is protected..
        '''</summary>
        Public ReadOnly Property PasswordDialog_ProtectedDocument() As String
            Get
                Return ResourceManager.GetString("PasswordDialog_ProtectedDocument", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Application View.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_ApplicationView() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_ApplicationView", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to L.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_ApplicationView_KeyTip() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_ApplicationView_KeyTip", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Left to Right.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_LeftToRight() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_LeftToRight", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to L.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_LeftToRight_KeyTip() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_LeftToRight_KeyTip", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Restart the program to change its view to a left to right text appearance..
        '''</summary>
        Public ReadOnly Property RibbonViewTab_LeftToRight_ToolTip_Description() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_LeftToRight_ToolTip_Description", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Left to Right.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_LeftToRight_ToolTip_Title() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_LeftToRight_ToolTip_Title", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Right to Left.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_RightToLeft() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_RightToLeft", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Restart the program to change its view to a right to left text appearance..
        '''</summary>
        Public ReadOnly Property RibbonViewTab_RightToLeft_ToolTip_Description() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_RightToLeft_ToolTip_Description", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Right to Left.
        '''</summary>
        Public ReadOnly Property RibbonViewTab_RightToLeft_ToolTip_Title() As String
            Get
                Return ResourceManager.GetString("RibbonViewTab_RightToLeft_ToolTip_Title", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Column: .
        '''</summary>
        Public ReadOnly Property StatusBar_Column() As String
            Get
                Return ResourceManager.GetString("StatusBar_Column", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Line: .
        '''</summary>
        Public ReadOnly Property StatusBar_Line() As String
            Get
                Return ResourceManager.GetString("StatusBar_Line", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Page: .
        '''</summary>
        Public ReadOnly Property StatusBar_Page() As String
            Get
                Return ResourceManager.GetString("StatusBar_Page", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Section: .
        '''</summary>
        Public ReadOnly Property StatusBar_Section() As String
            Get
                Return ResourceManager.GetString("StatusBar_Section", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
