Imports System.Windows.Forms
Imports Microsoft.Office
Imports Microsoft.Office.Tools
Imports Microsoft.Office.Tools.Ribbon


Public Class OutlookExtention
    Dim control As ImportWizard
    Private ImportWizard As CustomTaskPane


    Private Sub BtnLaunchImportWizard_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnLaunchImportWizard.Click

        ' --------------------------------------------------------------------------'
        ' Only creates new instance once
        If control Is Nothing Then ImportWizard = CreateNewInstance()

        ' --------------------------------------------------------------------------'
        ' Allows toggle for visible property
        ImportWizard.Visible = Not ImportWizard.Visible

    End Sub

    Private Function CreateNewInstance() As ImportWizard

        ' --------------------------------------------------------------------------'
        ' Creating a new Instance of the Import Wizard
        control = New ImportWizard
        ImportWizard = Globals.ThisAddIn.CustomTaskPanes.Add(control, "Import Wizard")
        ImportWizard.Width = 420
        ImportWizard.Visible = True

        Return ImportWizard
    End Function
End Class
