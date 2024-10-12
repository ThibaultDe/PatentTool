Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim AddIn = Globals.ThisAddIn
        'Dim NumRefs As Object
        'NumRefs = AddIn.NumRefs()

        Dim myUserControl1 = New MyUserControl
        Dim myCustomTaskPane = AddIn.CustomTaskPanes.Add(myUserControl1, "Références")
        myUserControl1.ListView1.Items.Clear()
        myUserControl1.ListView1.Columns.Clear()
        myUserControl1.ListView1.Columns.Add("Number", 100)
        myUserControl1.ListView1.Columns.Add("Text", 150)
        myUserControl1.ListView1.View = System.Windows.Forms.View.Details

        myCustomTaskPane.Visible = True




        ''--Print--'
        'For Each Key In NumRefs.Keys
        '    Dim RefsArray = NumRefs(Key)
        '    myUserControl1.ListBox1.Items.Add(Key)
        '    For i = 0 To UBound(RefsArray)
        '        If i > 0 Then
        '            myUserControl1.ListBox1.Items.Add("")
        '        End If
        '        myUserControl1.ListBox2.Items.Add(RefsArray(i))
        '    Next i

        'Next Key
    End Sub
End Class

