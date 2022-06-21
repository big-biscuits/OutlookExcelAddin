Imports System.Windows.Forms

Public Class ImportWizard
    Private OutlookApp As New OutlookApplication
    Private Sub ImportWizard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim AvailableMailboxes As List(Of String)
        Dim AvailableFolders As List(Of String)

        ' -----------------------------------------------------------------------------------
        AvailableMailboxes = OutlookApp.GetAvailableMailboxes
        AvailableFolders = OutlookApp.GetAvailableFolders


        ' -----------------------------------------------------------------------------------
        ' populate listbox with available mailboxes
        AddItemsToListbox("ListBox1", AvailableMailboxes)
        ' populate listbox with available folders
        AddItemsToListbox("Listbox2", AvailableFolders)




    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim SelectedFolders As List(Of String)
        Dim AvailableFolders As List(Of String)

        ' -----------------------------------------------------------------------------------
        If Me.ListBox1.SelectedItems.Count.Equals(0) Then
            AvailableFolders = OutlookApp.GetAvailableFolders()
        Else
            SelectedFolders = GetSelectedListboxItems("Listbox1")
            AvailableFolders = OutlookApp.GetAvailableFolders(SelectedFolders)
        End If

        AddItemsToListbox("Listbox2", AvailableFolders)
        Me.ErrorMessage.Text = ""

    End Sub


    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Me.ErrorMessage.Text = ""
    End Sub


    Private Function AddItemsToListbox(ControlName As String, Items As List(Of String))
        Dim ctr As ListBox = Me.Controls.Find(ControlName, True).FirstOrDefault

        ctr.Items.Clear()
        For Each Item In Items
            ctr.Items.Add(Item)
        Next

        Return True
    End Function



    Private Function GetSelectedListboxItems(ControlName As String) As List(Of String)
        Dim ctr As ListBox = Me.Controls.Find(ControlName, True).FirstOrDefault
        Dim SelectedItems As New List(Of String)

        For Each Item As String In ctr.SelectedItems
            SelectedItems.Add(Item)
        Next

        Return SelectedItems
    End Function

    Private Sub ImportBtn_Click(sender As Object, e As EventArgs) Handles ImportBtn.Click
        Dim SelectedMailboxes As List(Of String)
        Dim SelectedFolders As List(Of String)
        Dim Count As Integer

        ' -----------------------------------------------------------------------------------
        ' Check if user has selected at least one mailbox and foler
        If Me.ListBox1.SelectedItems.Count > 0 And Me.ListBox2.SelectedItems.Count > 0 Then
            SelectedMailboxes = GetSelectedListboxItems("Listbox1")
            SelectedFolders = GetSelectedListboxItems("Listbox2")
            Count = OutlookApp.ImportEmails(SelectedMailboxes, SelectedFolders)

        ElseIf Me.ListBox1.SelectedItems.Count.Equals(0) Then
            Me.ErrorMessage.Text = "Please select at least one mailbox"

        ElseIf Me.ListBox2.SelectedItems.Count.Equals(0) Then
            Me.ErrorMessage.Text = "Please select at least one folder"

        End If

        ' -----------------------------------------------------------------------------------
        ' Update label with status
        Me.ErrorMessage.Text = "Imported " & Count & " new emails"

    End Sub
End Class