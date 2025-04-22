Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles FindRefs.Click
        Dim AddIn = Globals.ThisAddIn

        Dim myUserControl1 = New MyUserControl()
        Dim myCustomTaskPane = AddIn.CustomTaskPanes.Add(myUserControl1, "Références")

        myCustomTaskPane.Width = 350
        myUserControl1.ListView1.Items.Clear()
        myUserControl1.ListView1.Columns.Clear()
        myUserControl1.ListView1.Columns.Add("Number", 100)
        myUserControl1.ListView1.Columns.Add("Text", 225)
        myUserControl1.ListView1.View = System.Windows.Forms.View.Details

        myCustomTaskPane.Visible = True



    End Sub

    Private Sub FindRefs_Disposed(sender As Object, e As EventArgs) Handles FindRefs.Disposed

    End Sub
End Class

