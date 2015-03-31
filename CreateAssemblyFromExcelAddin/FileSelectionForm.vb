Imports CreateAssemblyFromExcelAddin.CreateAssemblyFromExcelAddin
Imports CreateAssemblyFromExcelAddin.CreateAssemblyFromExcelAddin.StandardAddInServer
Imports System.Collections.Generic
Imports Autodesk.DataManagement.Client.Framework
Imports Autodesk.Connectivity.WebServices
Imports System.Windows.Forms
Imports System.Linq

Public Class FileSelectionForm

    Private m_connection As Vault.Currency.Connections.Connection

    Sub New(Connection As Vault.Currency.Connections.Connection)
        ' TODO: Complete member initialization 
        InitializeComponent()
        m_connection = Connection
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles BtnDone.Click
        Dim selectedfileitem As ListBoxFileItem = DirectCast(SearchResultsListBox.SelectedItem, ListBoxFileItem)
        Dim fileIterations As New List(Of Vault.Currency.Entities.FileIteration)()
        fileIterations.Add(selectedfileitem.File)

        If selectedfileitem IsNot Nothing Then
            If FoundList IsNot Nothing Then
                If Not FoundList.Contains(selectedfileitem) Then
                    FoundList.Add(selectedfileitem)
                Else
                    Dim idx As Integer = FoundList.FindIndex(Function(f As ListBoxFileItem) f.File.EntityName = selectedfileitem.File.EntityName)
                    selectedfile = FoundList(idx)
                End If
            End If
            Me.Close()
        Else
            MessageBox.Show("You need to select a file!")
        End If
    End Sub

    Public Sub button2_Click(sender As Object, e As EventArgs) Handles BtnNoFile.Click
        NoMatch = True
    End Sub
End Class