Imports System.Drawing
Imports System.Windows.Forms


Public Class OutlookApplication
    Dim OutlookApp As Microsoft.Office.Interop.Outlook.Application
    Dim OutlookStore As Microsoft.Office.Interop.Outlook.Store
    Dim OutlookRoot As Microsoft.Office.Interop.Outlook.Folder
    Dim OutlookRootFolders As Microsoft.Office.Interop.Outlook.Folders
    Dim OutlookFolder As Microsoft.Office.Interop.Outlook.Folder


    Public Function GetAvailableMailboxes()
        Dim AvailableMailboxes As List(Of String) = New List(Of String)
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        For Each OutlookStore In OutlookApp.Session.Stores

            AvailableMailboxes.Add(OutlookStore.DisplayName.ToLower())

        Next

        AvailableMailboxes.Sort()
        Return AvailableMailboxes

    End Function


    Public Function GetAvailableFolders()
        Dim AvailableFolders As List(Of String) = New List(Of String)
        Dim OutlookFolderName As String
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        For Each OutlookStore In OutlookApp.Session.Stores
            OutlookRoot = OutlookStore.GetRootFolder
            OutlookRootFolders = OutlookRoot.Folders
            For Each OutlookFolder In OutlookRootFolders
                OutlookFolderName = OutlookFolder.Name.ToLower()
                If Not AvailableFolders.Contains(OutlookFolderName) Then AvailableFolders.Add(OutlookFolderName)

            Next
        Next

        AvailableFolders.Sort()
        Return AvailableFolders

    End Function


    Public Function GetAvailableFolders(ByVal SelectedMailBoxes As List(Of String))
        Dim AvailableFolders As List(Of String) = New List(Of String)
        Dim OutlookFolderName As String
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        For Each OutlookStore In OutlookApp.Session.Stores
            For Each Mailbox In SelectedMailBoxes
                If OutlookStore.DisplayName.ToLower.Equals(Mailbox) Then
                    OutlookRoot = OutlookStore.GetRootFolder
                    OutlookRootFolders = OutlookRoot.Folders
                    For Each OutlookFolder In OutlookRootFolders
                        OutlookFolderName = OutlookFolder.Name.ToLower()
                        If Not AvailableFolders.Contains(OutlookFolderName) Then AvailableFolders.Add(OutlookFolderName)

                    Next
                End If
            Next
        Next

        AvailableFolders.Sort()
        Return AvailableFolders

    End Function

    Public Function ImportEmails(Mailbox As List(Of String), Folders As List(Of String)) As Integer
        Dim AvailableFolders As List(Of String) = New List(Of String)
        Dim ItemCount As Integer : ItemCount = 0
        Dim Items(999, 6) As String

        ' --------------------------------------------------------------------------'
        OutlookApp = New Microsoft.Office.Interop.Outlook.Application
        For Each OutlookStore In OutlookApp.Session.Stores
            If Mailbox.Contains(OutlookStore.DisplayName.ToLower()) Then

                OutlookRoot = OutlookStore.GetRootFolder
                OutlookRootFolders = OutlookRoot.Folders
                For Each OutlookFolder In OutlookRootFolders
                    If Folders.Contains(OutlookFolder.Name.ToLower()) Then

                        For Each OutlookMailItem As Object In OutlookFolder.Items

                            If IsNew(OutlookMailItem.EntryID) Then
                                ' ------------------------------------------------------------------------
                                ' OutlookMailItem source
                                Items(ItemCount, 0) = OutlookStore.DisplayName
                                Items(ItemCount, 1) = OutlookFolder.Name
                                ' ------------------------------------------------------------------------
                                ' OutlookMailItem properties
                                Items(ItemCount, 2) = OutlookMailItem.ReceivedTime
                                Items(ItemCount, 3) = OutlookMailItem.SenderName
                                Items(ItemCount, 4) = OutlookMailItem.Subject
                                Items(ItemCount, 5) = OutlookMailItem.Body
                                Items(ItemCount, 6) = OutlookMailItem.EntryID
                                ' Iterator
                                ItemCount = ItemCount + 1

                            End If
                        Next
                    End If
                Next
            End If
        Next

        ' --------------------------------------------------------------------------'
        If ItemCount > 0 Then PasteItems(Items, ItemCount)
        Return ItemCount

    End Function


    Private Function IsNew(EntityId As String) As Boolean
        ' function responcible for checking each
        ' entity if against the entity ids already
        ' imported on the import worksheet
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim LastRow As Integer
        Dim ExistingEntities As Excel.Range

        ' --------------------------------------------------------------------------'
        wb = Globals.ThisAddIn.Application.ActiveWorkbook
        ws = wb.ActiveSheet
        LastRow = ws.Range("A1048576").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        ExistingEntities = ws.Range("G2:G" & LastRow)

        ' --------------------------------------------------------------------------'
        ' Rreturns true if entity id is NOT already 
        ' in the import worksheet
        If ExistingEntities.Find(EntityId) Is Nothing Then Return True
        Return False
    End Function




    Private Function PasteItems(Items(,) As String, ItemCount As Integer) As Boolean
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim NewRow As Integer
        Dim DestinationRange As Excel.Range

        ' -----------------------------------------------------------------------------------
        wb = Globals.ThisAddIn.Application.ActiveWorkbook
        ws = wb.ActiveSheet

        If IsEmptyWorksheet() Then AddHeadings()
        NewRow = ws.Range("A1048576").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row + 1

        ' -----------------------------------------------------------------------------------
        ' Creating the destination for outlook items and 
        ' formatting the destination area
        DestinationRange = ws.Range(ws.Cells(NewRow, 1), ws.Cells(NewRow + ItemCount - 1, 7))
        DestinationRange.Cells.Font.Size = 8
        DestinationRange.Cells.WrapText = True
        DestinationRange.Cells.RowHeight = 112.5
        DestinationRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop
        DestinationRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

        ' -----------------------------------------------------------------------------------
        ' Outputting outlook items
        DestinationRange.Value = Items

        Return True
    End Function


    Private Function AddHeadings() As Boolean
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim DestinationRange As Excel.Range
        Dim ColumnHeaders() As String = {"Mailbox", "Folder", "ReceivedDate", "SenderName", "Subject", "Body", "EntityId"}

        ' -----------------------------------------------------------------------------------
        wb = Globals.ThisAddIn.Application.ActiveWorkbook
        ws = wb.ActiveSheet

        ' -----------------------------------------------------------------------------------
        ' Creating Destination Range and
        ' Formatting as headers
        DestinationRange = ws.Range("A1:G1")
        DestinationRange.RowHeight = 45
        DestinationRange.Cells.Font.Size = 12
        DestinationRange.Cells.Font.Color = Color.FromArgb(51, 51, 51)
        DestinationRange.Cells.Interior.Color = Color.FromArgb(188, 223, 236)
        DestinationRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        DestinationRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        ' Manually setting column width
        ws.Range("A1").ColumnWidth = 40
        ws.Range("B1").ColumnWidth = 20
        ws.Range("C1").ColumnWidth = 20
        ws.Range("D1").ColumnWidth = 20
        ws.Range("E1").ColumnWidth = 30
        ws.Range("F1").ColumnWidth = 60
        ws.Range("G1").ColumnWidth = 30
        ' -----------------------------------------------------------------------------------
        ' Outputting Column Headers
        DestinationRange.Value = ColumnHeaders

        Return True
    End Function


    Private Function IsEmptyWorksheet() As Boolean
        Dim wb As Excel.Workbook
        Dim ws As Excel.Worksheet
        Dim LastRow As Integer

        ' -----------------------------------------------------------------------------------

        wb = Globals.ThisAddIn.Application.ActiveWorkbook
        ws = wb.ActiveSheet
        LastRow = ws.Range("A1048576").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        ' -----------------------------------------------------------------------------------
        If LastRow.Equals(1) Then Return True
        Return False

    End Function



End Class
