'-----------------------------------------------------------------------------------------------------------
' PasswordDialog.vb File
'
' Description:
'     Provides a dialog that sets the password that is used to load a password protected document.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.IO
Imports TXTextControl.WPF

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for PasswordDialog.xaml
    ''' </summary>
    Partial Public Class PasswordDialog
        Inherits Window

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------
        Private m_txTextControl As TextControl
        Private m_lsLoadSettings As LoadSettings


        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' PasswordDialog Constructor
        ' Opens the password dialog to set the password that is used to load a password protected document.
        '
        ' Parameters:
        '      textControl:    The TextControl where to load the password protected document.
        '		loadSettings:	The LoadSettings that are used to load the password protected document.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New(ByVal textControl As TextControl, ByVal loadSettings As LoadSettings)
            m_txTextControl = textControl
            m_lsLoadSettings = loadSettings

            Me.InitializeComponent()

            ' Set some texts.
            Me.Title = My.Resources.PasswordDialog_Caption
            Me.m_lblProtectedDocument.Content = String.Format(My.Resources.PasswordDialog_ProtectedDocument, Path.GetFileName(m_lsLoadSettings.LoadedFile))
            Me.m_lblEnterPassword.Content = My.Resources.PasswordDialog_EnterPassword
            Me.m_lblPassword.Content = My.Resources.PasswordDialog_Password
            Me.m_btnOK.Content = My.Resources.PasswordDialog_OK
            Me.m_btnCancel.Content = My.Resources.PasswordDialog_Cancel
        End Sub



        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' OK_Click Handler
        ' When clicking the OK button the set password from the text box is used to open the password protected 
        ' document.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            If LoadDocument() Then
                DialogResult = True
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Password_TextChanged Handler
        ' Updates the enabled state of the OK button.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Password_TextChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_btnOK.IsEnabled = Me.m_tbxPassword.Password.Length > 0
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' LoadDocument Method
        ' Uses the set password from the text box to load the password protected document.
        ' 
        ' 	Return value:		True if the password protected document could be loaded.
        ' 						Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Private Function LoadDocument() As Boolean
            Try
                ' Set the password from the text box and...
                m_lsLoadSettings.UserPassword = Me.m_tbxPassword.Password
                ' ... try to load the document.
                m_txTextControl.Load(m_lsLoadSettings.LoadedFile, m_lsLoadSettings.LoadedStreamType, m_lsLoadSettings)
                Return True
            Catch ex As FilterException
                If TypeOf ex Is FilterException Then
                    Select Case TryCast(ex, FilterException).Reason
                        Case FilterException.FilterError.InvalidPassword
                            ' The set password is incorrect.
                            ' Inform the user, reset the password and return false.
                            MessageBox.Show(Me, My.Resources.MessageBox_PasswordDialog_IncorrectPassword_Text, My.Resources.MessageBox_PasswordDialog_IncorrectPassword_Caption, MessageBoxButton.OK, MessageBoxImage.Warning)
                            m_lsLoadSettings.UserPassword = String.Empty
                            Return False
                    End Select
                End If
                Throw ex
            End Try
        End Function
    End Class
End Namespace
