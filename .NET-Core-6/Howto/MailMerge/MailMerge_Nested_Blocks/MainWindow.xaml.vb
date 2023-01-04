'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'      Sample project that is related to the 'Howto: Mail Merge - Sample: Mail Merge with Nested 
'	   Repeating Blocks' article inside the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Data
Imports Microsoft.Win32

Class MainWindow
    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_dsData As DataSet
    Private m_mmMailMerge As TXTextControl.DocumentServer.MailMerge


    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_Loaded Handler
    ' Creates load the 'Accruals Report.docx' template.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_Loaded(sender As Object, e As RoutedEventArgs)
        ' Create a new data set and load the XML file
        m_dsData = New DataSet()
        m_dsData.ReadXml("Files\Data.xml")

        Dim lsLoadSettings As TXTextControl.LoadSettings = New TXTextControl.LoadSettings With {
            .ApplicationFieldFormat = TXTextControl.ApplicationFieldFormat.MSWord,
            .LoadSubTextParts = True
        }

        ' Load the 'Template.docx' template
        Me.m_txTextControl.Load("Files\Accruals Report.docx", TXTextControl.StreamType.WordprocessingML, lsLoadSettings)

        ' Initialize a MailMerge instance.
        m_mmMailMerge = New TXTextControl.DocumentServer.MailMerge()
        m_mmMailMerge.TextComponent = Me.m_txTextControl
        AddHandler m_mmMailMerge.BlockRowMerged, AddressOf MailMerge_BlockRowMerged
        ' Set focus to the TextControl.
        Me.m_txTextControl.Focus()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Datasource_Click Handler
    ' Get the reference to the 'Data.xml' sample
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Datasource_Click(sender As Object, e As RoutedEventArgs)
        TryCast(sender, Button).ContextMenu.IsEnabled = True
        TryCast(sender, Button).ContextMenu.PlacementTarget = TryCast(sender, Button)
        TryCast(sender, Button).ContextMenu.Placement = Primitives.PlacementMode.Bottom
        TryCast(sender, Button).ContextMenu.IsOpen = True
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' LoadSampleDatasource_Click Handler
    ' Get the reference to the 'Data.xml' sample
    '-----------------------------------------------------------------------------------------------------------
    Private Sub LoadSampleDatasource_Click(sender As Object, e As RoutedEventArgs)
        ' Update the text box.
        Me.m_tbxLoadedDatabaseFile.Tag = "Files\Data.xml"
        Me.m_tbxLoadedDatabaseFile.Text = "Data.xml"
        Me.m_tmiCreateReport.IsEnabled = True
        Me.m_pthRightArraow.Fill = New SolidColorBrush(Colors.Black)
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' LoadXML_Click Handler
    ' Create and open an OpenFileDialog to get the reference to an XML database.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub LoadXML_Click(sender As Object, e As RoutedEventArgs)
        ' Create and open an OpenFileDialog to load an XML database.
        Dim dlgLoadXML As OpenFileDialog = New OpenFileDialog()
        dlgLoadXML.Filter = "XML Database | *.xml"
        dlgLoadXML.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory

        If dlgLoadXML.ShowDialog(Me) = True Then
            ' Update the text box.
            Me.m_tbxLoadedDatabaseFile.Tag = dlgLoadXML.FileName
            Me.m_tbxLoadedDatabaseFile.Text = dlgLoadXML.SafeFileName
            Me.m_tmiCreateReport.IsEnabled = True
            Me.m_pthRightArraow.Fill = New SolidColorBrush(Colors.Black)
        End If
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' CreateReport_Click Handler
    ' Merge the template with the data source.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub CreateReport_Click(sender As Object, e As RoutedEventArgs)
        Try
            ' Load the XML file.
            Dim dsData As DataSet = New DataSet()
            dsData.ReadXml(TryCast(Me.m_tbxLoadedDatabaseFile.Tag, String), XmlReadMode.Auto)

            ' Add the relations for the main block and its child blocks.
            Dim relCompanyEmployee As DataRelation = New DataRelation("company_employee", dsData.Tables("company").Columns("company_number"), dsData.Tables("employee").Columns("company_number"))

            Dim relEmployeeSick As DataRelation = New DataRelation("employee_sick", dsData.Tables("employee").Columns("employee_number"), dsData.Tables("sick").Columns("employee_number"))

            Dim relEmployeeVacation As DataRelation = New DataRelation("employee_vacation", dsData.Tables("employee").Columns("employee_number"), dsData.Tables("vacation").Columns("employee_number"))

            dsData.Relations.Add(relCompanyEmployee)
            dsData.Relations.Add(relEmployeeSick)
            dsData.Relations.Add(relEmployeeVacation)

            ' Update the progress bar.
            Me.m_pbProgress.Maximum = dsData.Tables("employee").Rows.Count

            ' Merge.
            m_mmMailMerge.Merge(dsData.Tables("company"), True)

            ' Reset the progress bar.
            Me.m_pbProgress.Value = 0
        Catch exc As Exception
            MessageBox.Show(Me, exc.Message)
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' MailMerge_BlockRowMerged Handler
    ' Update the progress bar when the 'employee' merge block is handled.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub MailMerge_BlockRowMerged(ByVal sender As Object, ByVal e As TXTextControl.DocumentServer.MailMerge.BlockRowMergedEventArgs)
        If Equals(e.MergeBlockName, "employee") Then
            Me.m_pbProgress.Value += 1
        End If
    End Sub
End Class
