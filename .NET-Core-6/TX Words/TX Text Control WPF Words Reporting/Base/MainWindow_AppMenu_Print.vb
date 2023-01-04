'-----------------------------------------------------------------------------------------------------------
' MainWindow_AppMenu_Print.vb File
'
' Description:
'     Manages printing a document
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.ComponentModel
Imports System.Printing

Namespace TXTextControl.Words
    Partial Class MainWindow

        '-----------------------------------------------------------------------------------------------------------
        ' H A N D L E R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Print_Click Handler
        ' Invokes the TextControl Print method to open the TextControl print dialog.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub Print_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ' Use the active document name to open the print dialog.
            Me.m_txTextControl.Print(m_strActiveDocumentName, True)
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' PrintQuick_Click Handler
        ' Prints the current document without opening the dialog before.
        '-----------------------------------------------------------------------------------------------------------
        Private Sub PrintQuick_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.m_txTextControl.Print(m_strActiveDocumentName, New PageRange(1, Me.m_txTextControl.Pages), 1, Collation.Collated)
        End Sub

        '------------------------------------------------------------------------------------------------------------
        ' TextControl_PropertyChanged_Print Handler
        ' Update the print button's enabled states.
        '------------------------------------------------------------------------------------------------------------
        Private Sub TextControl_PropertyChanged_Print(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "CanPrint"
                    Me.m_rbtnPrintQAT.IsEnabled = CSharpImpl.Assign(Me.m_rsmiPrint.IsEnabled, CSharpImpl.Assign(Me.m_rbtnPrint.IsEnabled, CSharpImpl.Assign(Me.m_rbtnPrintQuick.IsEnabled, Me.m_txTextControl.CanPrint)))
            End Select
        End Sub
    End Class
End Namespace
