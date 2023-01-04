'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_UserAccess.vb File
'
' Description:
'      Handles the user access to the document.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------

Namespace TXTextControl.Words
    Partial Public Class MainWindow
        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R   V A R I A B L E S
        '-----------------------------------------------------------------------------------------------------------

        Private m_uiCurrentUser As UserInfo = Nothing ' Info about the current user.
        Private m_strUserName As String = "" ' Environment.UserName

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' SignIn_Click Handler
        ' Opens a dialog to sign in to the TextControl a user by its account. If no such account is known,
        ' a new account is created. 
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SignIn_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Open the password dialog to sign in or create a user account.
            Dim dlgUserAccessDialog As UserAccessDialog = If(m_uiCurrentUser Is Nothing, New UserAccessDialog(m_strUserName), New UserAccessDialog(m_uiCurrentUser))
            If dlgUserAccessDialog.ShowDialog().Value Then
                ' Get the UserInfo instance that represents the current signed in user.
                m_uiCurrentUser = dlgUserAccessDialog.UserInfo

                ' Give the user access to the TextControl.
                Me.m_txTextControl.UserNames = New String() {m_uiCurrentUser.Name}

                ' Hide the Sign In button.
                Me.m_rmiSignIn.IsEnabled = False
                Me.m_rmiSignIn.Visibility = Visibility.Collapsed

                ' Show the [Current User] button.
                Me.m_rmbtnCurrentUser.IsEnabled = True
                Me.m_rmbtnCurrentUser.Visibility = Visibility.Visible

                ' Save the settings of the current user.
                SaveKnownUserSettings()
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' AccountSettings_Click Handler
        ' Opens a dialog to edit the account settings of the current signed in user.
        '----------------------------------------------------------------------------------------------------------
        Private Sub AccountSettings_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim dlg As UserAccessDialog = New UserAccessDialog(m_uiCurrentUser)
            Dim bDialogResult As Boolean? = dlg.ShowDialog()
            If dlg.DeletePassword Then
                ' The UserAccessDialog's Delete Account button was clicked. 
                ' Set current user to null...
                m_uiCurrentUser = Nothing
                ' ... and sign out.
                SignOut()
                ' Save the settings of the current user.
                SaveKnownUserSettings()
            Else
                If bDialogResult.Value Then
                    ' Replace the user info object of the current user
                    ' with a new instance.
                    m_uiCurrentUser = dlg.UserInfo
                    Me.m_rmbtnCurrentUser.Visibility = Visibility.Visible
                    ' Save the settings of the current user.
                    SaveKnownUserSettings()
                End If
            End If
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SignOut_Click Handler
        ' Signs out the current signed in user.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SignOut_Click(ByVal sender As Object, ByVal e As EventArgs)
            SignOut()
            MessageBox.Show(Me, String.Format(My.Resources.MessageBox_UserAccess_SignOut_Text, m_strUserName), My.Resources.MessageBox_UserAccess_SignOut_Caption, MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' LoadKnownUserSettings Method
        ' Gets the known user from the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub LoadKnownUserSettings()
            m_strUserName = Environment.UserName ' Get the user name of the person who is currently logged on the operation system
            m_uiCurrentUser = My.Settings.Default.KnownUser
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SaveKnownUserSettings Method
        ' Save the known user to the application settings.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SaveKnownUserSettings()

            ' Save the know users to the My.Settings.Default.KnownUsers property
            My.Settings.Default.KnownUser = m_uiCurrentUser
            Call My.Settings.Default.Save()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' SignOut Method
        ' Signs out the current user.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub SignOut()
            'Signs out the current user.
            If m_uiCurrentUser IsNot Nothing Then
                m_uiCurrentUser.IsSignedIn = False
            End If
            ' Reset the TextControl user access.
            Me.m_txTextControl.UserNames = Nothing

            ' Show the Sign In button.
            Me.m_rmiSignIn.IsEnabled = True
            Me.m_rmiSignIn.Visibility = Visibility.Visible

            ' Hide the [Current User] button.
            Me.m_rmbtnCurrentUser.IsEnabled = False
            Me.m_rmbtnCurrentUser.Visibility = Visibility.Collapsed
        End Sub
    End Class
End Namespace
