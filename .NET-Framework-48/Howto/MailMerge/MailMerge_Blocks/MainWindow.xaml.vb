'-----------------------------------------------------------------------------------------------------------
' MainWindow.xaml.vb File
'
' Description:
'      Sample project that is related to the 'Howto: Mail Merge -> Sample: Mail Merge with Repeating Blocks' 
'		article inside the 'Windows Presentation Foundation User's Guide'.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Data

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
    ' Creates a new data set and loads the 'Template.docx' template.
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
        Me.m_txTextControl.Load("Files\Template.docx", TXTextControl.StreamType.WordprocessingML, lsLoadSettings)

        ' Initialize a MailMerge instance.
        m_mmMailMerge = New TXTextControl.DocumentServer.MailMerge()
        m_mmMailMerge.TextComponent = Me.m_txTextControl

        ' Set focus to the TextControl.
        Me.m_txTextControl.Focus()
    End Sub

    '-----------------------------------------------------------------------------------------------------------
    ' Merge_Click Handler
    ' Use the MailMerge instance to merge the data.
    '-----------------------------------------------------------------------------------------------------------
    Private Sub Merge_Click(sender As Object, e As RoutedEventArgs)
        m_mmMailMerge.Merge(m_dsData.Tables("orders"), True)
    End Sub
End Class
