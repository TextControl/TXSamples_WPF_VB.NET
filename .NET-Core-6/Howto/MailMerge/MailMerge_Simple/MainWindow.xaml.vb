'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'      Sample project that is related to the 'Howto: Mail Merge -> Sample: Simple Mail Merge'
'	   article inside the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Data

Class MainWindow
    '-----------------------------------------------------------------------------------------------------------
    ' M E M B E R   V A R I A B L E S
    '-----------------------------------------------------------------------------------------------------------
    Private m_dsAddresses As DataSet
    Private m_mmMailMerge As TXTextControl.DocumentServer.MailMerge


    '-----------------------------------------------------------------------------------------------------------
    ' H A N D L E R S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' TextControl_Loaded Handler
    ' Creates the addresses data set, adds an item for each database field to the 'Add'item drop down and 
    ' loads the 'Instructions.tx'sample template.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub TextControl_Loaded(sender As Object, e As RoutedEventArgs)
        'Create a new DataSet and load the XML file.
        m_dsAddresses = New DataSet()
        m_dsAddresses.ReadXml("Files\Data.xml")

        'Create a new ToolStripMenuItem for each database field.
        For Each dataColumn As DataColumn In m_dsAddresses.Tables(0).Columns
            Dim mnuItem As MenuItem = New MenuItem()
            mnuItem.Header = dataColumn.ColumnName

            AddHandler mnuItem.Click, New RoutedEventHandler(AddressOf DatabaseFieldItem_Click)
            Me.m_miAdd.Items.Add(mnuItem)
        Next

        'Initialize a MailMerge instance.
        m_mmMailMerge = New TXTextControl.DocumentServer.MailMerge()
        m_mmMailMerge.TextComponent = Me.m_txTextControl

        'Load the 'Instructions.tx'sample template.
        Me.m_txTextControl.Selection.Load("Files\Instructions.tx", TXTextControl.StreamType.InternalUnicodeFormat)

        'Set focus to the TextControl.
        Me.m_txTextControl.Focus()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' 'Application Fields'Drop Down
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' ApplicationFields_DropDownOpening Handler
    ' Sets the enabled state of the 'Add'and 'Properties'items when the 'Application Fields'
    ' drop down is opening.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub ApplicationFields_DropDownOpening(sender As Object, e As RoutedEventArgs)
        Me.m_miProperties.IsEnabled = If(Me.m_txTextControl.ApplicationFields.GetItem() Is Nothing, False, True)
        Me.m_miAdd.IsEnabled = Me.m_txTextControl.ApplicationFields.CanAdd
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' DatabaseFieldItem_Click Handler
    ' Creates with the text of the clicked database field item a new TXTextControl.DocumentServer.Fields.MergeField  
    ' and adds it to TextControl.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub DatabaseFieldItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim tmiClickedItem = CType(sender, MenuItem)
        InsertMergeField(tmiClickedItem.Header.ToString())
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Properties_Click Handler
    ' Creates with the text of the clicked database field item a new TXTextControl.DocumentServer.Fields.MergeField  
    ' and adds it to TextControl.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Properties_Click(sender As Object, e As RoutedEventArgs)
        Dim mfMergeField As TXTextControl.DocumentServer.Fields.MergeField = New TXTextControl.DocumentServer.Fields.MergeField(Me.m_txTextControl.ApplicationFields.GetItem())
        mfMergeField.ShowDialog(Me)
    End Sub



    '-----------------------------------------------------------------------------------------------------------
    ' 'Mail Merge'Drop Down
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' MailMerge_DropDownOpening Handler
    ' Sets the enabled state of the 'Merge'item when the 'Mail Merge'drop down is opening.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub MailMerge_DropDownOpening(sender As Object, e As RoutedEventArgs)
        Me.m_miMerge.IsEnabled = Me.m_txTextControl.ApplicationFields.Count > 0
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Merge_Click Handler
    ' Use the MailMerge instance to merge the data into the application fields.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Merge_Click(sender As Object, e As RoutedEventArgs)
        m_mmMailMerge.Merge(m_dsAddresses.Tables(0), True)
    End Sub


    '-----------------------------------------------------------------------------------------------------------
    ' M E T H O D S
    '-----------------------------------------------------------------------------------------------------------

    '-----------------------------------------------------------------------------------------------------------
    ' InsertMergeField Method
    ' Creates with the text of the clicked database field item a new TXTextControl.DocumentServer.Fields.MergeField  
    ' and adds it to TextControl.
    '
    ' Parameters:
    '		name:		The name of the merge field that is created and added to the TextControl.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub InsertMergeField(ByVal name As String)
        'Create a new TXTextControl.DocumentServer.Fields.MergeField
        'and add it to TextControl.
        Dim mfMergeField As TXTextControl.DocumentServer.Fields.MergeField = New TXTextControl.DocumentServer.Fields.MergeField()
        mfMergeField.Name = name
        mfMergeField.Text = "{ " & name & " }"
        mfMergeField.ApplicationField.HighlightMode = TXTextControl.HighlightMode.Activated
        mfMergeField.ApplicationField.DoubledInputPosition = True

        Me.m_txTextControl.ApplicationFields.Add(mfMergeField.ApplicationField)
    End Sub
End Class
