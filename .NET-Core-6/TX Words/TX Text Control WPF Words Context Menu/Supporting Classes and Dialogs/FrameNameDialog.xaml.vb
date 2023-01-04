'-----------------------------------------------------------------------------------------------------------
' FrameNameDialog.xaml.vb File
'
' Description:
'     Provides a dialog to to edit the name of a frame.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for FrameNameDialog.xaml
    ''' </summary>
    Partial Public Class FrameNameDialog
        Inherits Window
        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' FrameNameDialog Constructor
        ' Opens the dialog to edit the name of a frame.
        '
        ' Parameters:
        '      frameName:   The current name of the frame.
        '-----------------------------------------------------------------------------------------------------------
        Friend Sub New(ByVal frameName As String)
            Me.InitializeComponent()
            ' Set some texts
            Me.Title = My.Resources.ContextMenu_FrameNameDialog_Caption
            Me.m_lblFrameName.Content = My.Resources.ContextMenu_FrameNameDialog_Label
            Me.m_btnOK.Content = My.Resources.ContextMenu_FrameNameDialog_OK
            Me.m_btnCancel.Content = My.Resources.ContextMenu_FrameNameDialog_Cancel

            Me.m_tbxFrameName.Text = frameName ' Set the text box text.
            Me.m_tbxFrameName.Focus()
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' FrameName Property
        ' Returns the text of the frame name text box.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property FrameName As String
            Get
                Return Me.m_tbxFrameName.Text
            End Get
        End Property


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' OK_Click Handler
        ' Closes the dialog by setting the DialogResult property to true.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            DialogResult = True
        End Sub
    End Class
End Namespace
