Imports System.Text.RegularExpressions

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for MergeSettingsDialog.xaml
    ''' </summary>
    Partial Public Class MergeSettingsDialog
        Inherits Window
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S 
        '-----------------------------------------------------------------------------------------------------------
        Private ReadOnly m_rgxPositivNumber As Regex = New Regex("^[1-9]+\d*$")

        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' MergeSettingsDialog Constructor
        ' Creates a dialog to determine the settings for the following merge process.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New()
            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.MergeSettingsDialog_Caption
            Me.m_lblRecords.Content = My.Resources.MergeSettingsDialog_Records
            Me.m_lblNumberOfRecords.Content = My.Resources.MergeSettingsDialog_NumberOfRecords
            Me.m_chbxMergeAllRecords.Content = My.Resources.MergeSettingsDialog_MergeAllRecords
            Me.m_chbxMergeIntoSingleDocument.Content = My.Resources.MergeSettingsDialog_MergeIntoSingleDocument
            Me.m_lblRemoveEmptyMergeElements.Content = My.Resources.MergeSettingsDialog_RemoveEmptyMergeElements
            Me.m_chbxBlocks.Content = My.Resources.MergeSettingsDialog_Blocks
            Me.m_chbxImages.Content = My.Resources.MergeSettingsDialog_Images
            Me.m_chbxFields.Content = My.Resources.MergeSettingsDialog_Fields
            Me.m_chbxTrailingWhitespace.Content = My.Resources.MergeSettingsDialog_TrailingWhitespace
            Me.m_chbxLines.Content = My.Resources.MergeSettingsDialog_Lines
            Me.m_btnOK.Content = My.Resources.MergeSettingsDialog_OK
            Me.m_btnCancel.Content = My.Resources.MergeSettingsDialog_Cancel
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' P R O P E R T I E S 
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' MaxRecords Property
        ' Returns the number of files that should be merged. If the "Merge all records" check box is checked, the
        ' property returns int:MaxValue
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property MaxRecords As Integer
            Get
                Return If(Me.m_chbxMergeAllRecords.IsChecked.Value, Integer.MaxValue, Integer.Parse(Me.m_tbxNumberOfRecords.Text))
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' MergeIntoSingleFile Property
        ' Returns a value indicating whether all created files should be merged into a single file.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property MergeIntoSingleFile As Boolean
            Get
                Return Me.m_chbxMergeIntoSingleDocument.IsChecked.Value
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveEmptyBlocks Property
        ' Returns a value indicating whether or not the content of empty merge blocks should be removed from the
        ' template.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property RemoveEmptyBlocks As Boolean
            Get
                Return Me.m_chbxBlocks.IsChecked.Value
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveEmptyFields Property
        ' Returns a value indicating whether or not empty fields should be removed from the template.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property RemoveEmptyFields As Boolean
            Get
                Return Me.m_chbxFields.IsChecked.Value
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveEmptyLines Property
        ' Returns a value indicating whether or not text lines which are empty after merging should be removed from
        ' the template.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property RemoveEmptyLines As Boolean
            Get
                Return Me.m_chbxLines.IsChecked.Value
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveEmptyImages Property
        ' Returns a value indicating whether or not images which don't have merge data should be removed from the
        ' template.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property RemoveEmptyImages As Boolean
            Get
                Return Me.m_chbxImages.IsChecked.Value
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' RemoveTrailingWhitespace Property
        ' Returns a value indicating whether trailing whitespace should be removed before saving a document. 
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property RemoveTrailingWhitespace As Boolean
            Get
                Return Me.m_chbxTrailingWhitespace.IsChecked.Value
            End Get
        End Property


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' OK_Click Handler
        ' Closes the dialog with DialogResult.OK when clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            DialogResult = True
            Close()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' MergeAllRecords_CheckedChanged Handler
        ' Enables/Disables the "Number of records:" label and text box when the "Merge all records" check box
        ' was checked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub MergeAllRecords_CheckedChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_lblNumberOfRecords.IsEnabled = CSharpImpl.Assign(Me.m_tbxNumberOfRecords.IsEnabled, Not Me.m_chbxMergeAllRecords.IsChecked.Value)
            Me.m_btnOK.IsEnabled = Me.m_chbxMergeAllRecords.IsChecked.Value OrElse Me.m_tbxNumberOfRecords.Text.Length > 0
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' NumberOfRecords_PreviewTextInput Handler
        ' Validates the text input: Only positive numbers are allowed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub NumberOfRecords_PreviewTextInput(ByVal sender As Object, ByVal e As Input.TextCompositionEventArgs)
            e.Handled = Not m_rgxPositivNumber.IsMatch(e.Text)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' NumberOfRecords_TextChanged Handler
        ' Disables the OK button if no number is set.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub NumberOfRecords_TextChanged(ByVal sender As Object, ByVal e As Controls.TextChangedEventArgs)
            If Me.m_btnOK IsNot Nothing Then
                Me.m_btnOK.IsEnabled = Me.m_tbxNumberOfRecords.Text.Length > 0
            End If
        End Sub

        Private Class CSharpImpl
            Shared Function Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
