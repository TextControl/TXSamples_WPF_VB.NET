'-----------------------------------------------------------------------------------------------------------
' UserAccessDialog.xaml.vb File
'
' Description:
'     Provides a dialog to create a new user account, sign in a user or edit a user account.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    ''' <summary>
    ''' Interaction logic for UserAccessDialog.xaml
    ''' </summary>
    Partial Public Class UserAccessDialog
        Inherits Window
        '-----------------------------------------------------------------------------------------------------------
        ' E N U M S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' DialogBehaviors Enum
        ' Represents three kinds of behavior how the dialog might act.
        '-----------------------------------------------------------------------------------------------------------
        Friend Enum DialogBehaviors
            CreateAccount      ' The dialog is opened to create a new user account.
            SignIn             ' The dialog is opened to sign in a user.
            AccountSettings     ' The dialog is opened to edit the signed in user's account.
        End Enum


        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------

        Private ReadOnly m_uiUserInfo As UserInfo = Nothing
        Private ReadOnly m_strUserName As String = Nothing
        Private ReadOnly m_dbDialogBehavior As DialogBehaviors = DialogBehaviors.CreateAccount
        Private m_bDeletePassword As Boolean = False


        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UserAccessDialog Constructor
        ' Opens the dialog to sign in a user or to edit the signed in user's account settings.
        '
        ' Parameters:
        '      userInfo:   The UserInfo instance of the user to be handled.
        '-----------------------------------------------------------------------------------------------------------
        Friend Sub New(ByVal userInfo As UserInfo)
            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.UserAccessDialog_AccountSettings_Caption
            Me.m_btnDelete.Content = My.Resources.UserAccessDialog_AccountSettings_Delete
            Me.m_btnOK.Content = My.Resources.UserAccessDialog_OK
            Me.m_btnCancel.Content = My.Resources.UserAccessDialog_Cancel

            m_uiUserInfo = userInfo

            ' Set the user name
            Me.m_lblUserName.Content = CSharpImpl.Assign(m_strUserName, m_uiUserInfo.Name)

            ' Check whether to sign in the user or edit its account settings.
            If m_uiUserInfo.IsSignedIn Then
                ' Edit user account settings.
                m_dbDialogBehavior = DialogBehaviors.AccountSettings

                ' Update controls texts
                Title = My.Resources.UserAccessDialog_AccountSettings_Caption

                Me.m_lblPassword.Content = My.Resources.UserAccessDialog_AccountSettings_OldPassword
                Me.m_lblNewPassword.Content = My.Resources.UserAccessDialog_AccountSettings_NewPassword
                Me.m_lblConfirmPassword.Content = My.Resources.UserAccessDialog_AccountSettings_ConfirmPassword

                ' Buttons
                Me.m_btnDelete.Visibility = Visibility.Visible
            Else
                ' Otherwise the user is known but not signed in.
                m_dbDialogBehavior = DialogBehaviors.SignIn

                ' Hide 'New password' and 'Confirm password' controls
                Me.m_lblNewPassword.Visibility = CSharpImpl.Assign(Me.m_tbxNewPassword.Visibility, CSharpImpl.Assign(Me.m_lblConfirmPassword.Visibility, CSharpImpl.Assign(Me.m_tbxConfirmPassword.Visibility, Visibility.Collapsed)))

                ' Update controls texts
                Title = My.Resources.UserAccessDialog_SignIn_Caption
                Me.m_lblPassword.Content = My.Resources.UserAccessDialog_SignIn_Password
            End If
            Me.m_tbxPassword.Focus()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' UserAccessDialog Constructor
        ' Opens the dialog to create a user account.
        '
        ' Parameters:
        '      userName:   The name of the user to create an account for.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New(ByVal userName As String)
            Me.InitializeComponent()

            ' Set some texts
            Me.Title = My.Resources.UserAccessDialog_AccountSettings_Caption
            Me.m_btnDelete.Content = My.Resources.UserAccessDialog_AccountSettings_Delete
            Me.m_btnOK.Content = My.Resources.UserAccessDialog_OK
            Me.m_btnCancel.Content = My.Resources.UserAccessDialog_Cancel

            ' Set the user name
            Me.m_lblUserName.Content = CSharpImpl.Assign(m_strUserName, userName)

            ' A new account should be created.
            m_dbDialogBehavior = DialogBehaviors.CreateAccount

            ' Hide password controls
            Me.m_lblPassword.Visibility = CSharpImpl.Assign(Me.m_tbxPassword.Visibility, Visibility.Collapsed)

            ' Update control texts
            Title = My.Resources.UserAccessDialog_CreateAccount_Caption
            Me.m_lblNewPassword.Content = My.Resources.UserAccessDialog_CreateAccount_NewPassword
            Me.m_lblConfirmPassword.Content = My.Resources.UserAccessDialog_CreateAccount_ConfirmPassword
            Me.m_tbxNewPassword.Focus()
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UserInfo Property
        ' Returns an instance of a UserInfo class that represents the signed in user.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property UserInfo As UserInfo
            Get
                Select Case m_dbDialogBehavior
                    Case DialogBehaviors.SignIn
                        m_uiUserInfo.IsSignedIn = True
                        Return m_uiUserInfo
                    Case Else
                        Return New UserInfo(m_strUserName, Me.m_tbxConfirmPassword.Password)
                End Select
            End Get
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' DeletePassword Property
        ' Returns a value whether the password should be deleted.
        '-----------------------------------------------------------------------------------------------------------
        Friend ReadOnly Property DeletePassword As Boolean
            Get
                Return m_bDeletePassword
            End Get
        End Property


        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Password_TextChanged Handler
        ' Updates the IsEnabled states of the dialog controls when the text of a text box changed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Password_TextChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
            EnableControls()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' Delete_Click Handler
        ' Asks the user whether his user account should be deleted. In that case and if the 
        ' password is correct, the dialog is closed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Delete_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Ask the user whether the current user account should be deleted.
            If MessageBox.Show(Me, My.Resources.MessageBox_UserAccessDialogDeleteAccount_Text, Title, MessageBoxButton.OKCancel, MessageBoxImage.Warning) = MessageBoxResult.OK Then
                ' Validate the password of the current signed in user.
                If Not m_uiUserInfo.ValidatePassword(Me.m_tbxPassword.Password) Then
                    MessageBox.Show(Me, My.Resources.MessageBox_UserAccessDialogIncorrectPassword_Text, Title, MessageBoxButton.OK, MessageBoxImage.Exclamation)
                    Return
                End If
                ' Close the dialog
                m_bDeletePassword = True
                Close()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' OK_Click Handler
        ' Validates the password when the OK button is clicked.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub OK_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Validate the password of the current signed in user or the user to sign in.
            If m_uiUserInfo IsNot Nothing AndAlso Not m_uiUserInfo.ValidatePassword(Me.m_tbxPassword.Password) Then
                ' The password is not correct.
                MessageBox.Show(Me, My.Resources.MessageBox_UserAccessDialogIncorrectPassword_Text, Title, MessageBoxButton.OK, MessageBoxImage.Exclamation)
                Return
            End If

            DialogResult = True
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' EnableControls Method
        ' Updates the IsEnabled states of the dialog controls when the text of a text box changed.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub EnableControls()
            Select Case m_dbDialogBehavior
                Case DialogBehaviors.CreateAccount
                    ' Enable/Disable the confirm password text 
                    If Not CSharpImpl.Assign(Me.m_tbxConfirmPassword.IsEnabled, Me.m_tbxNewPassword.Password.Length > 0) Then
                        Me.m_tbxConfirmPassword.Password = "" ' Reset the confirm password text box if it is disabled.
                    End If
                    ' Enable/Disable the OK button.
                    Me.m_btnOK.IsEnabled = Me.m_tbxConfirmPassword.IsEnabled AndAlso Equals(Me.m_tbxNewPassword.Password, Me.m_tbxConfirmPassword.Password)
                Case DialogBehaviors.SignIn
                    ' Enable/Disable the OK button.
                    Me.m_btnOK.IsEnabled = Me.m_tbxPassword.Password.Length > 0
                Case DialogBehaviors.AccountSettings
                    ' Enable/Disable the confirm password text 
                    If Not CSharpImpl.Assign(Me.m_tbxConfirmPassword.IsEnabled, Me.m_tbxNewPassword.Password.Length > 0) Then
                        Me.m_tbxConfirmPassword.Password = ""  ' Reset the confirm password text box if it is disabled.
                    End If
                    ' Enable/Disable the OK and Delete button.
                    Me.m_btnOK.IsEnabled = CSharpImpl.Assign(Me.m_btnDelete.IsEnabled, Me.m_tbxPassword.Password.Length > 0) AndAlso Me.m_tbxConfirmPassword.IsEnabled AndAlso Equals(Me.m_tbxNewPassword.Password, Me.m_tbxConfirmPassword.Password)
            End Select
        End Sub

        Private Class CSharpImpl
            Shared Function Assign(Of T)(ByRef target As T, value As T) As T
                target = value
                Return value
            End Function
        End Class
    End Class
End Namespace
