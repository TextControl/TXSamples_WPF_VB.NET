'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'     Implements TX Text Control Words.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits RibbonWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_strFilesDirectory As String = "Files\"

        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow Constructor
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New()
            ' Add an unhandled exception handler
            Dim currentDomain As AppDomain = AppDomain.CurrentDomain
            AddHandler currentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException

            Me.InitializeComponent()
            ' Set some texts
            Me.Title = My.Resources.MainWindow_Caption_Product
            Me.m_rtRibbonTableLayoutTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_TableTools
            Me.m_rtRibbonFormulaTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_TableTools
            Me.m_rtRibbonFrameLayoutTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_FrameTools
            Me.m_ctgTableTools.Header = My.Resources.ContextualTabGroup_TableTools
            Me.m_ctgFrameTools.Header = My.Resources.ContextualTabGroup_FrameTools

            ' Set texts of the template buttons
            Me.m_rmbtnOpenSampleTemplate.Header = My.Resources.ApplicationMenu_OpenSampleTemplate
            Me.m_rmbtnOpenSampleTemplate.ToolTip = My.Resources.ApplicationMenu_OpenSampleTemplate_ToolTip_Description
            Me.m_rmbtnOpenSampleTemplate.KeyTip = My.Resources.ApplicationMenu_OpenSampleTemplate_KeyTip
            Me.m_rbtnSampleInvoiceTemplate.Header = My.Resources.ApplicationMenu_SampleInvoiceTemplate
            Me.m_rbtnSampleInvoiceTemplate.KeyTip = My.Resources.ApplicationMenu_SampleInvoiceTemplate_KeyTip
            Me.m_rbtnSamplePackingListTemplate.Header = My.Resources.ApplicationMenu_SamplePackingListTemplate
            Me.m_rbtnSamplePackingListTemplate.KeyTip = My.Resources.ApplicationMenu_SamplePackingListTemplate_KeyTip
            Me.m_rbtnSampleShippingLabelTemplate.Header = My.Resources.ApplicationMenu_SampleShippingLabelTemplate
            Me.m_rbtnSampleShippingLabelTemplate.KeyTip = My.Resources.ApplicationMenu_SampleShippingLabelTemplate_KeyTip

            ' Set texts of the 'Result' tab and its buttons
            Me.m_ctgReportingResult.Header = My.Resources.ContextualTabGroup_Reporting
            Me.m_rtMergeResultsTab.Header = My.Resources.MergeResultsTab
            Me.m_rtMergeResultsTab.ContextualTabGroupHeader = My.Resources.ContextualTabGroup_Reporting
            Me.m_rtMergeResultsTab.KeyTip = My.Resources.MergeResultsTab_KeyTip
            Me.m_rgMergeResultsTab_ResultGroup.Header = My.Resources.MergeResultsTab_Result
            Me.m_rgMergeResultsTab_ResultGroup.KeyTip = My.Resources.MergeResultsTab_Result_KeyTip
            Me.m_rbtnExitMergeResultsTab.Label = My.Resources.MergeResultsTab_Exit
            Me.m_rbtnExitMergeResultsTab.KeyTip = My.Resources.MergeResultsTab_Exit_KeyTip
            Me.m_rbtnExitMergeResultsTab.ToolTipTitle = My.Resources.MergeResultsTab_Exit_ToolTip_Title
            Me.m_rbtnExitMergeResultsTab.ToolTipDescription = My.Resources.MergeAndExport_ToolTip_Description
            Me.m_rgNavigateGroup.Header = My.Resources.MergeResultsTab_Navigate
            Me.m_rgNavigateGroup.KeyTip = My.Resources.MergeResultsTab_Navigate_KeyTip
            Me.m_rbtnFirstRecord.Label = My.Resources.MergeResultsTab_FirstRecord
            Me.m_rbtnFirstRecord.KeyTip = My.Resources.MergeResultsTab_FirstRecord_KeyTip
            Me.m_rbtnPreviousRecord.Label = My.Resources.MergeResultsTab_PreviousRecord
            Me.m_rbtnPreviousRecord.KeyTip = My.Resources.MergeResultsTab_PreviousRecord_KeyTip
            Me.m_rbtnNextRecord.Label = My.Resources.MergeResultsTab_NextRecord
            Me.m_rbtnNextRecord.KeyTip = My.Resources.MergeResultsTab_NextRecord_KeyTip
            Me.m_rbtnLastRecord.Label = My.Resources.MergeResultsTab_LastRecord
            Me.m_rbtnLastRecord.KeyTip = My.Resources.MergeResultsTab_LastRecord_KeyTip
            Me.m_rgExportGroup.Header = My.Resources.MergeResultsTab_Export
            Me.m_rgExportGroup.KeyTip = My.Resources.MergeResultsTab_Export_KeyTip
            Me.m_rbtnExportMergeResult.Label = My.Resources.ExportMergeResult
            Me.m_rbtnExportMergeResult.KeyTip = My.Resources.ExportMergeResult_KeyTip
            Me.m_rbtnExportMergeResult.ToolTipTitle = My.Resources.ExportMergeResult_ToolTip_Title
            Me.m_rbtnExportMergeResult.ToolTipDescription = My.Resources.ExportMergeResult_ToolTip_Description

            ' Get and set saved application settings.
            LoadRecentFiles()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------
        ' SetAppMenuDesign Method
        ' Updates the design of the menu items.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetAppMenuDesign()
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_New.ToString(), Me.m_rmiNew)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Open.ToString(), Me.m_rmiOpen)
            SetOpenSampleTemplateButtonDesign()
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Save.ToString(), Me.m_rmiSave)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_SaveAs.ToString(), Me.m_rmiSaveAs)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Print.ToString(), Me.m_rsmiPrint)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_Print.ToString(), Me.m_rbtnPrint)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_PrintQuick.ToString(), Me.m_rbtnPrintQuick)
            Me.SetMenuItemDesign(ResourceProvider.FileMenuItem.TXITEM_About.ToString(), Me.m_rtbtnAbout, My.Resources.AboutButton_ToolTip_Description)

            ' Set Recent Files overview label text
            Me.m_rgcRecentFiles.Header = My.Resources.ApplicationMenu_ResentFiles
        End Sub

        '-----------------------------------------------------------------------------------------------------
        ' SetRibbonButtonDesign Method
        ' Sets the icons, text, key tip and tool tip for a specific RibbonButton.
        '
        ' Parameters:
        '      resourceID:     The id that is used to identify the corresponding texts and icons.
        '      menuItem:	   The ribbon menu item to update.
        '      args:           Optional strings that are used to format the displayed texts.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetMenuItemDesign(ByVal resourceID As String, ByVal menuItem As RibbonMenuItem, ParamArray args As String())
            menuItem.Name = resourceID
            menuItem.ImageSource = ResourceProvider.GetLargeIcon(resourceID, Me)
            menuItem.KeyTip = ResourceProvider.GetKeyTip(resourceID)

            If args.Length > 0 Then
                menuItem.Header = String.Format(ResourceProvider.GetText(resourceID), args)
                menuItem.ToolTipTitle = String.Format(ResourceProvider.GetToolTipTitle(resourceID), args)
                menuItem.ToolTipDescription = String.Format(ResourceProvider.GetToolTipDescription(resourceID), args)
            Else
                menuItem.Header = ResourceProvider.GetText(resourceID)
                menuItem.ToolTipTitle = ResourceProvider.GetToolTipTitle(resourceID)
                menuItem.ToolTipDescription = ResourceProvider.GetToolTipDescription(resourceID)
            End If
        End Sub


        '-----------------------------------------------------------------------------------------------------
        ' SetAppMenuBehavior Method
        ' Connects all necessary handlers to the application menu items.
        '-----------------------------------------------------------------------------------------------------
        Private Sub SetAppMenuBehavior()
            ' Common
            AddHandler Me.m_txTextControl.Changed, AddressOf Me.TextControl_Changed ' Updates the internal 'is dirty document' flag

            ' New:
            AddHandler Me.m_rmiNew.Click, AddressOf Me.New_Click

            ' Open:
            AddHandler Me.m_rmiOpen.Click, AddressOf Me.Open_Click

            ' Save:
            AddHandler Me.m_rmiSave.Click, AddressOf Me.Save_Click

            ' Save As:
            AddHandler Me.m_rmiSaveAs.Click, AddressOf Me.SaveAs_Click

            ' Print:
            AddHandler Me.m_rsmiPrint.Click, AddressOf Me.Print_Click ' Print(Split Button)
            AddHandler Me.m_rbtnPrint.Click, AddressOf Me.Print_Click ' Print (Drop Down Button)
            AddHandler Me.m_rbtnPrintQuick.Click, AddressOf Me.PrintQuick_Click ' Print Quick (Drop Down Button)
            AddHandler Me.m_txTextControl.PropertyChanged, AddressOf Me.TextControl_PropertyChanged_Print ' Add handler for the print buttons Enable state handling
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' GetSmallIcon Method
        ' Creates a small bitmap icon from an embedded SVG resource.
        '
        ' Parameters:
        '      path:   The path to the embedded SVG resource.
        '
        ' Returns:     The created small bitmap.
        '-----------------------------------------------------------------------------------------------------------
        Private Function GetSmallIcon(ByVal path As String) As BitmapSource
            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim bmp As BitmapSource = Nothing

            Using sStream = asm.GetManifestResourceStream(path)
                bmp = ResourceProvider.GetSmallIcon(sStream, Me)
            End Using
            Return bmp
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' GetLargeIcon Method
        ' Creates a large bitmap icon from an embedded SVG resource.
        '
        ' Parameters:
        '      path:   The path to the embedded SVG resource.
        '
        ' Returns:     The created large bitmap.
        '-----------------------------------------------------------------------------------------------------------
        Private Function GetLargeIcon(ByVal path As String) As BitmapSource
            Dim asm As Assembly = Assembly.GetExecutingAssembly()
            Dim bmp As BitmapSource = Nothing

            Using sStream = asm.GetManifestResourceStream(path)
                bmp = ResourceProvider.GetLargeIcon(sStream, Me)
            End Using
            Return bmp
        End Function


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------
        ' MainWindow_Loaded Handler 
        ' Sets the requested behavior for all added controls.
        '-----------------------------------------------------------------------------------------------------
        Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Main Window:
            UpdateMainWindowCaption() ' Set caption

            ' Application Menu items:
            SetAppMenuDesign()
            SetAppMenuBehavior()

            ' QAT:
            SetQATItemsDesign()

            ' Sidebars:
            SetSidebarBehavior()

            ' Mini Toolbar
            Me.m_txTextControl.ShowMiniToolbar = MiniToolbarButton.LeftButton Or MiniToolbarButton.RightButton

            ' Contextual Tabs:
            SetContextualTabsBehavior()

            ' Toolbars:
            SetRulerBarsDesign()
            SetStatusBarDesign()

            ' Reporting Preview:
            AddSampleDatabaseButton()
            AddMergeGroup()
            SetMergeResultsTabDesign()
            SetResultTabBehavior()

            ' About:
            Me.UpdateAboutSidebar()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' TextControl_Loaded_MainWindow Handler 
        ' Sets the focus to the TextControl.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub TextControl_Loaded_MainWindow(ByVal sender As Object, ByVal e As RoutedEventArgs)
            m_txTextControl.Focus()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' MainWindow_Closing Handler
        ' Saves the recent files to the My.Setting.Default.RecentFiles property.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As CancelEventArgs)
            ' Save the recent files to the My.Setting.Default.RecentFiles property
            SaveRecentFiles()
        End Sub

        '-----------------------------------------------------------------------------------------------------
        ' CurrentDomain_UnhandledException Handler
        ' Handles an exception by showing a message box that explains the reason for the exception.
        '-----------------------------------------------------------------------------------------------------
        Friend Shared Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
            Dim strProductName As String = CType(Attribute.GetCustomAttribute(Assembly.GetExecutingAssembly(), GetType(AssemblyProductAttribute)), AssemblyProductAttribute).Product
            Dim ex As Exception = CType(e.ExceptionObject, Exception)

            ' TX Text Control Feature is not available
            If TypeOf ex Is LicenseLevelException Then
                MessageBox.Show(String.Format(My.Resources.MessageBox_Application_ThreadException_Text, ex.Message), strProductName, MessageBoxButton.OK, MessageBoxImage.Information)
                Return
            End If

            ' All other exceptions
            MessageBox.Show(String.Format(My.Resources.MessageBox_Application_ThreadException_Text, ex.Message), strProductName, MessageBoxButton.OK, MessageBoxImage.Error)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' C L A S S E S
        '-----------------------------------------------------------------------------------------------------------
        Private Class CSharpImpl
            Shared Function Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
