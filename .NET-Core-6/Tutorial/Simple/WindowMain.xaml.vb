'*------------------------------------------------------------------------------------------------
'Ä** program:			TX Text Control Simple Sample
'** description:	A Simple Sample to show the basic functionality of TX Text Control.						
'**
'** copyright:		© Text Control GmbH
'**----------------------------------------------------------------------------------------------*
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation
Imports System.Windows.Shapes


' Interaction logic for Window1.xaml
Partial Public Class WindowMain
    Inherits Window
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub MenuItem_Click_1(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Load()
    End Sub

    Private Sub MenuItem_Click_2(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Save()
    End Sub

    Private Sub MenuItem_Click_3(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.SectionFormatDialog(0)
    End Sub

    Private Sub MenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub MenuItem_Click_4(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Cut()
    End Sub

    Private Sub MenuItem_Click_5(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Copy()
    End Sub

    Private Sub MenuItem_Click_6(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Paste()
    End Sub

    Private Sub MenuItem_Click_7(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.FontDialog()
    End Sub

    Private Sub MenuItem_Click_8(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.ParagraphFormatDialog()
    End Sub

    Private Sub MenuItem_Click_9(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.SectionFormatDialog(2)
    End Sub

    Private Sub textControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        textControl1.Focus()
    End Sub


End Class
