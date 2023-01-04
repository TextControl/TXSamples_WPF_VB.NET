Imports System.ComponentModel
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.DocumentServer.DataSources
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for MergeAndWaitDialog.xaml
    ''' </summary>
    Partial Public Class MergeAndWaitDialog
        Inherits Window

        '----------------------------------------------------------------------------------------------
        ' M E M B E R S
        '----------------------------------------------------------------------------------------------
        Private ReadOnly m_bwMergeFiles As BackgroundWorker = New BackgroundWorker()
        Private m_lstMergedFiles As IList(Of Byte()) = Nothing
        Private m_bIsMergeProcessCanceled As Boolean = False
        Private m_exException As Exception = Nothing
        Private ReadOnly m_roArgs As Object() = Nothing


        '----------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R
        '----------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' MergeAndWaitDialog Constructor
        ' Creates dialog to create the merged files that should be displayed by the preview tab.
        '
        ' Parameters:
        '				template:		The template that is used to create the merged files.	
        '				maxPreviews:	The maximum number of merged files that should be created.	
        '				textControl:	The TextControl that is used to create the merged files.
        '				mergeSettings:	The merge settings that are used to create the merged files.	
        '				reportingTab:	The RibbonReportingTab instance that contains the 
        '								DataSourceManager that is used to create the merged files.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New(ByVal template As Byte(), ByVal maxPreviews As Integer, ByVal textControl As TextControl, ByVal mergeSettings As MergeSettings, ByVal reportingTab As RibbonReportingTab)
            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.MergeAndWaitDialog_Caption
            Me.m_lblMerging.Content = My.Resources.MergeAndWaitDialog_Merging

            ' Sets the last selected master table of the ReportingTab's MasterTable Menu as datasource.
            SetLastSelectedMasterTable(reportingTab)

            ' Store the arguments that are necessary to merge files.
            m_roArgs = New Object() {template, maxPreviews, textControl, mergeSettings, reportingTab}
        End Sub


        '----------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '----------------------------------------------------------------------------------------------

        '----------------------------------------------------------------------------------------------
        ' MergedFiles Property
        ' Returns the merged files.
        '----------------------------------------------------------------------------------------------
        Friend ReadOnly Property MergedFiles As IList(Of Byte())
            Get
                Return m_lstMergedFiles
            End Get
        End Property

        '----------------------------------------------------------------------------------------------
        ' Exception Property
        ' Returns the corresponding exception if triggered by the merge process.
        '----------------------------------------------------------------------------------------------
        Friend ReadOnly Property Exception As Exception
            Get
                Return m_exException
            End Get
        End Property


        '----------------------------------------------------------------------------------------------
        ' O V E R R I D D E N   M E T H O D S
        '----------------------------------------------------------------------------------------------

        '----------------------------------------------------------------------------------------------
        ' OnClosing Method (overridden)
        ' If the dialog was closed by the user, the dialog is disabled and a DataRowMerged event handler 
        ' is added to the DataSourceManager that cancels the merge process.
        '----------------------------------------------------------------------------------------------
        Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
            If m_bwMergeFiles.IsBusy Then
                IsEnabled = False
                m_bIsMergeProcessCanceled = True
                Dim rtReportingTab = CType(m_roArgs(4), RibbonReportingTab)
                RemoveHandler rtReportingTab.DataSourceManager.DataRowMerged, AddressOf DataRowMerged_Handler
                AddHandler rtReportingTab.DataSourceManager.DataRowMerged, AddressOf DataRowMerged_Handler
                e.Cancel = True
            End If
            MyBase.OnClosing(e)
        End Sub

        '----------------------------------------------------------------------------------------------
        ' M E T H O D S
        '----------------------------------------------------------------------------------------------

        '----------------------------------------------------------------------------------------------
        ' MergePreview Method
        ' Creates a list of merged files by calling the RibbonReportingTab's DataSourceManager.Merge
        ' method.
        '
        ' Parameters:
        '				template:		The template that is used to create the merged files.
        '				maxPreviews:	The maximum number of merged files that should be created.		
        '				textControl:	The TextControl that is used to create the merged files.
        '               mergeSettings:	The settings that are used to to create the merged files.
        '				reportingTab:	The RibbonReportingTab instance that contains the 
        '								DataSourceManager that is used to create the merged files.
        '----------------------------------------------------------------------------------------------
        Private Sub MergePreview(ByVal template As Byte(), ByVal maxPreviews As Integer, ByVal textControl As TextControl, ByVal msMergeSettings As MergeSettings, ByVal reportingTab As RibbonReportingTab)
            Try
                m_lstMergedFiles = reportingTab.DataSourceManager.Merge(template, maxPreviews, textControl, msMergeSettings)
            Catch e As Exception
                ' Store the exception if thrown.
                m_exException = e
            End Try
        End Sub

        '----------------------------------------------------------------------------------------------
        ' SetLastSelectedMasterTable Method
        ' Sets the last selected master table of the ReportingTab's MasterTable Menu as datasource.
        '
        ' Parameters:
        '				reportingTab:	The RibbonReportingTab instances that contains the master table.
        '----------------------------------------------------------------------------------------------
        Private Sub SetLastSelectedMasterTable(ByVal reportingTab As RibbonReportingTab)
            Dim rmbtnTXITEM_SelectMasterTable As RibbonMenuButton = TryCast(reportingTab.FindName(RibbonReportingTab.RibbonItem.TXITEM_SelectMasterTable.ToString()), RibbonMenuButton)
            For Each dropDownItem As Control In rmbtnTXITEM_SelectMasterTable.Items
                Dim rtbnTable As RibbonToggleButton = TryCast(dropDownItem, RibbonToggleButton)
                If rtbnTable IsNot Nothing AndAlso rtbnTable.IsChecked.Value Then
                    reportingTab.DataSourceManager.MasterDataTableInfo = TryCast(rtbnTable.Tag, DataTableInfo)
                    Exit For
                End If
            Next
        End Sub


        '----------------------------------------------------------------------------------------------
        ' E V E N T H A N D L E R 
        '----------------------------------------------------------------------------------------------

        '----------------------------------------------------------------------------------------------
        ' Window_Loaded Handler
        ' Starts the merge process with a background operation.
        '----------------------------------------------------------------------------------------------
        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Start the merge process with a background operation.
            AddHandler m_bwMergeFiles.DoWork, AddressOf BackgroundWorker_DoWork
            AddHandler m_bwMergeFiles.RunWorkerCompleted, AddressOf BackgroundWorker_RunWorkerCompleted
            m_bwMergeFiles.RunWorkerAsync(m_roArgs)
        End Sub

        '----------------------------------------------------------------------------------------------
        ' BackgroundWorker_DoWork Handler
        ' Start creating a list of merged preview files when the background worker starts its background
        ' operation.
        '----------------------------------------------------------------------------------------------
        Private Sub BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
            Dim roArgs As Object() = TryCast(e.Argument, Object())
            MergePreview(CType(roArgs(0), Byte()), roArgs(1), CType(roArgs(2), TextControl), CType(roArgs(3), MergeSettings), CType(roArgs(4), RibbonReportingTab))
        End Sub

        '----------------------------------------------------------------------------------------------
        ' BackgroundWorker_RunWorkerCompleted Handler
        ' Closes the dialog when the background operation was completed. Furthermore the list of merged 
        ' files is reset to null if no merged file  was created or the merge process was canceled. 
        '----------------------------------------------------------------------------------------------
        Private Sub BackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)
            If m_bIsMergeProcessCanceled OrElse m_lstMergedFiles IsNot Nothing AndAlso m_lstMergedFiles.Count = 0 Then
                m_lstMergedFiles = Nothing
            End If
            Close()
        End Sub

        '----------------------------------------------------------------------------------------------
        ' DataRowMerged_Handler Handler
        ' That handler is only added when the user tries to cancel the merge process by closing the 
        ' dialog. In this case the merge process is stopped after the next data row is merged.
        '----------------------------------------------------------------------------------------------
        Private Sub DataRowMerged_Handler(ByVal sender As Object, ByVal e As DocumentServer.MailMerge.DataRowMergedEventArgs)
            RemoveHandler TryCast(sender, DataSourceManager).DataRowMerged, AddressOf DataRowMerged_Handler
            e.Cancel = True ' Cancel the merge process.
        End Sub
    End Class
End Namespace
