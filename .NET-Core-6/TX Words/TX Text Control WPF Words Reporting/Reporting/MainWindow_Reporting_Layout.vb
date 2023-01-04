'-----------------------------------------------------------------------------------------------------------
' MainWindow_Reporting_Layout.vb File
'
' Description:
'		Sets the layout of the added application menu's sample template button, the RibbonReportingTab's  
'		'Database Sample' button,  Merge' group and the reporting "Result" tab.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Windows.Controls.Ribbon
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    Partial Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        ' RibbonReportingTab
        Private m_rmiSampleDatabaseButton As RibbonMenuItem = Nothing
        Private ReadOnly m_rgMerge As RibbonGroup = New RibbonGroup()
        Private m_rbtnMergeAndExport As RibbonButton = Nothing

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Application Menu
        '-----------------------------------------------------------------------------------------------------------
        '-----------------------------------------------------------------------------------------------------------
        ' SetOpenSampleTemplateButtonDesign Method
        ' Sets the image source of the 'Open Sample Template' menu button.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetOpenSampleTemplateButtonDesign()
            Me.m_rmbtnOpenSampleTemplate.ImageSource = GetLargeIcon("OpenSampleTemplate_Large.svg")
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' RibbonReportingTab
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' AddSampleDatabaseButton Method
        ' Creates a ribbon button that loads the sampled database when clicked. That button is added to the drop 
        ' down menu of the 'Select Data Source' button.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AddSampleDatabaseButton()
            ' Create the ribbon button
            m_rmiSampleDatabaseButton = New RibbonMenuItem() With {
                .Header = My.Resources.RibbonReportingTab_LoadSampleDatabase,
                .ImageSource = GetSmallIcon("SampleDatabase_Small.svg")
            }

            ' Add tool tips
            m_rmiSampleDatabaseButton.ToolTipTitle = My.Resources.RibbonReportingTab_LoadSampleDatabase_ToolTip_Title
            m_rmiSampleDatabaseButton.ToolTipDescription = My.Resources.RibbonReportingTab_LoadSampleDatabase_ToolTip_Description

            ' Add the handler that loads the sampled database when clicked
            AddHandler m_rmiSampleDatabaseButton.Click, AddressOf SampleDatabaseButton_Click

            ' Add the ribbon button to the drop down menu of the 'Select Data Source' button.
            Dim rsbtnSelectDataSource As RibbonSplitButton = TryCast(Me.m_rtRibbonReportingTab.FindName(RibbonReportingTab.RibbonItem.TXITEM_DataSource.ToString()), RibbonSplitButton)
            rsbtnSelectDataSource.Items.Insert(3, m_rmiSampleDatabaseButton)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' AddMergeGroup Method
        ' Creates a ribbon group with a ribbon button that starts merging files and switches to the 'Result' tab when 
        ' clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AddMergeGroup()
            ' Create the icons for the ribbon group and ribbon button.
            Dim bmpSmallIcon = GetSmallIcon("MergeAndExport_Small.svg")
            Dim bmpLargeIcon = GetLargeIcon("MergeAndExport_Large.svg")

            ' Set ribbon group design
            m_rgMerge.SmallImageSource = bmpSmallIcon
            m_rgMerge.LargeImageSource = bmpLargeIcon
            m_rgMerge.Header = My.Resources.Merge
            m_rgMerge.KeyTip = My.Resources.Merge_KeyTip

            ' Add a ribbon button that starts merging files and switches to the 'Result' tab when clicked.
            AddMergeAndExportButton(m_rgMerge, bmpSmallIcon, bmpLargeIcon)

            ' Add the ribbon group to the ribbon tab
            Me.m_rtRibbonReportingTab.Items.Add(m_rgMerge)

            ' The group's enabled state depends on the IsMergingPossible property value of the 
            ' RibbonReportingTab's DataSourceManager 
            m_rgMerge.IsEnabled = Me.m_rtRibbonReportingTab.DataSourceManager.IsMergingPossible
            AddHandler Me.m_rtRibbonReportingTab.DataSourceManager.IsMergingPossibleChanged, AddressOf Me.DataSourceManager_IsMergingPossibleChanged
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' AddMergeAndExportButton Method
        ' Creates a ribbon button that starts merging files and switches to the 'Result' tab when clicked. That 
        ' ribbon button is added to the specified ribbon group. 
        '
        ' Parameters:
        '      ribbonGroup:    The ribbon group where to add the created ribbon button.
        '      smallIcon:      The bitmap that is used as the ribbon button's small icon.
        '      largeIcon:      The bitmap that is used as the ribbon button's large icon.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub AddMergeAndExportButton(ByVal ribbonGroup As RibbonGroup, ByVal smallIcon As BitmapSource, ByVal largeIcon As BitmapSource)

            ' Create the ribbon button
            m_rbtnMergeAndExport = New RibbonButton() With {
                .Label = My.Resources.MergeAndExport,
                .SmallImageSource = smallIcon,
                .LargeImageSource = largeIcon,
                .KeyTip = My.Resources.MergeAndExport_KeyTip
            }

            ' Add tool tips
            m_rbtnMergeAndExport.ToolTipTitle = My.Resources.MergeAndExport_ToolTip_Title
            m_rbtnMergeAndExport.ToolTipDescription = My.Resources.MergeAndExport_ToolTip_Description

            ' Add the handler that starts merging files and switches to the 'Result' tab when clicked.
            AddHandler m_rbtnMergeAndExport.Click, AddressOf MergeAndExport_Click

            ' Add the ribbon button to the ribbon group
            ribbonGroup.Items.Add(m_rbtnMergeAndExport)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' 'Reporting' ContextualTabGroup and 'Result' RibbonTab
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SetMergeResultsTabDesign Method
        ' Creates groups and sets the design of the reporting 'Result' tab.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SetMergeResultsTabDesign()
            ' Set the images of the 'Result' group items
            Me.m_rgMergeResultsTab_ResultGroup.SmallImageSource = CSharpImpl.Assign(Me.m_rbtnExitMergeResultsTab.SmallImageSource, ResourceProvider.GetSmallIcon(ResourceProvider.FileMenuItem.TXITEM_Exit.ToString(), Me))
            Me.m_rgMergeResultsTab_ResultGroup.LargeImageSource = CSharpImpl.Assign(Me.m_rbtnExitMergeResultsTab.LargeImageSource, ResourceProvider.GetLargeIcon(ResourceProvider.FileMenuItem.TXITEM_Exit.ToString(), Me))

            ' Set the images of the 'Navigate' group items 
            Me.m_rgNavigateGroup.SmallImageSource = CSharpImpl.Assign(Me.m_rbtnNextRecord.SmallImageSource, ResourceProvider.GetSmallIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToNext.ToString(), Me))
            Me.m_rgNavigateGroup.LargeImageSource = CSharpImpl.Assign(Me.m_rbtnNextRecord.LargeImageSource, ResourceProvider.GetLargeIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToNext.ToString(), Me))
            Me.m_rbtnFirstRecord.SmallImageSource = ResourceProvider.GetSmallIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToFirst.ToString(), Me)
            Me.m_rbtnFirstRecord.LargeImageSource = ResourceProvider.GetLargeIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToFirst.ToString(), Me)
            Me.m_rbtnPreviousRecord.SmallImageSource = ResourceProvider.GetSmallIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToPrevious.ToString(), Me)
            Me.m_rbtnPreviousRecord.LargeImageSource = ResourceProvider.GetLargeIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToPrevious.ToString(), Me)
            Me.m_rbtnLastRecord.SmallImageSource = ResourceProvider.GetSmallIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToLast.ToString(), Me)
            Me.m_rbtnLastRecord.LargeImageSource = ResourceProvider.GetLargeIcon(ResourceProvider.GeneralItem.TXITEM_NavigateToLast.ToString(), Me)

            ' Set the images of the 'Export' group items 
            Me.m_rgExportGroup.SmallImageSource = CSharpImpl.Assign(Me.m_rbtnExportMergeResult.SmallImageSource, GetSmallIcon("MergeAndExport_Small.svg"))
            Me.m_rgExportGroup.LargeImageSource = CSharpImpl.Assign(Me.m_rbtnExportMergeResult.LargeImageSource, GetLargeIcon("MergeAndExport_Large.svg"))
        End Sub
    End Class
End Namespace
