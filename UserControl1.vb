Imports System.Windows.Controls
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports System.Windows.Input
Imports System.Linq
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Diagnostics
Imports Word = Microsoft.Office.Interop.Word
Imports System.Drawing.Text

Public Class MyUserControl

    Private selectedItem As System.Windows.Forms.ListViewItem = Nothing
    Private selectedCol As Integer = -1 ' Pour stocker l'index de la colonne sélectionnée
    Private selectedIndex As Integer
    Private SelectedRef As String
    Private currentRange As Word.Range

    Private Sub Replace_Click(sender As Object, e As EventArgs) Handles Replace.Click
        Dim DocRange = Globals.ThisAddIn.Application.ActiveDocument.Range
        Dim RevRange As Object
        Dim RevStart As Object

        With DocRange.Find
            .Text = "REVENDICATIONS"        'Cherche le texte en GRAS'
            .Font.Bold = True
            .Forward = False
            .Execute()
        End With

        If DocRange.Find.Found = True Then
            System.Diagnostics.Debug.WriteLine("Revendications Start trouvé ")
            RevStart = DocRange.Start
        Else
            MsgBox("Revendications non trouvées")
        End If

        RevRange = Globals.ThisAddIn.Application.ActiveDocument.Range(RevStart, Globals.ThisAddIn.Application.ActiveDocument.Range.End)

        Dim ind = selectedIndex
        Dim Number = ListView1.Items(ind).SubItems(0).Text
        Debug.Print(ind & " " & Number)
        While Number = ""
            ind = ind - 1
            Number = ListView1.Items(ind).SubItems(0).Text
            Debug.Print(ind & " " & Number)
        End While

        'Dim RepText = SelectedRef + " (" + Number + ")"

        Dim currentRange As Word.Range = RevRange

        With currentRange.Find
            .Text = SelectedRef ' Rechercher le texte sélectionné
            .Font.Italic = False ' Ignorer la mise en forme italique dans la recherche
        End With

        If currentRange.Find.Execute() Then
            Dim foundText As String = currentRange.Text
            Debug.Write("  " & foundText) ' Imprime le texte trouvé

            currentRange.Text = foundText + " (" + Number + ")"
            currentRange.Font.Italic = True ' Mettre en italique le texte remplacé

            'Avancer la plage après l'occurrence remplacée
            currentRange.Start = currentRange.End

            currentRange.Select() ' Sélectionner la plage remplacée
            Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(currentRange)
        Else
            currentRange = Nothing ' Réinitialiser pour recommencer depuis le début
            MessageBox.Show("Aucune autre occurrence trouvée dans la plage sélectionnée.")
        End If

    End Sub

    Private Sub FindRefs_Click_1(sender As Object, e As EventArgs) Handles FindRefs.Click
        ListView1.Items.Clear()

        Dim AddIn = Globals.ThisAddIn
        Dim NumRefs As Object
        NumRefs = AddIn.NumRefs()

        Dim N = NumRefs.Count
        If N = 0 Then
            Return
        End If

        Debug.Print("N=" & N)
        Dim KeyList As New List(Of Integer)()

        For Each Key In NumRefs.Keys
            KeyList.Add(CInt(Key))
        Next

        KeyList.Sort()

        For Each Key In KeyList
            Debug.Print(Key.ToString)
            Dim RefsArray = NumRefs(Key.ToString)
            Dim Num As String
            For i = 0 To UBound(RefsArray)
                If i = 0 Then
                    Num = Key.ToString()
                Else
                    Num = ""
                End If
                Dim lvi As New System.Windows.Forms.ListViewItem(Num)
                lvi.SubItems.Add(RefsArray(i))
                ListView1.Items.Add(lvi)
            Next i
        Next
    End Sub



    '----------- Trouver et surligner l'élément sélectionné ---------------------------------'
    '----------------------------------------------------------------------------------------'
    Private Sub ListView1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseDown
        ' Utiliser HitTest pour déterminer quel élément ou sous-élément a été cliqué
        Dim info As ListViewHitTestInfo = ListView1.HitTest(e.X, e.Y)
        ' Si un sous-élément a été cliqué
        If info.SubItem IsNot Nothing Then
            ' Obtenir l'indice de la colonne cliquée
            selectedItem = info.Item
            selectedCol = info.Item.SubItems.IndexOf(info.SubItem)
            selectedIndex = selectedItem.Index
            ' Afficher un message avec le texte et l'indice de la colonne
            SelectedRef = info.SubItem.Text

            Debug.Print("Colonne " & selectedCol.ToString() & " cliquée : " & info.SubItem.Text)
            ' Forcer le ListView à se redessiner
            ListView1.Invalidate()
        End If


        For Each Item As System.Windows.Forms.ListViewItem In ListView1.Items
            Item.Selected = False
        Next
    End Sub

    Private Sub ListView1_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs) Handles ListView1.DrawColumnHeader
        e.DrawDefault = True ' Dessiner les en-têtes de colonne normalement
    End Sub
    Private Sub ListView1_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs) Handles ListView1.DrawSubItem
        'Si la colonne courante est celle sélectionnée, on dessine avec une couleur de surbrillance
        If e.Item Is selectedItem AndAlso e.ColumnIndex = selectedCol Then
            e.Graphics.FillRectangle(Brushes.LightBlue, e.Bounds) ' Coloration en bleu clair
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.Item.Font, e.Bounds, Color.Black, TextFormatFlags.Left)
        Else
            ' Sinon, dessiner normalement
            e.DrawDefault = True
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Activer OwnerDraw pour les sous-éléments (nécessaire pour dessiner manuellement)
        ListView1.OwnerDraw = True
        ListView1.FullRowSelect = False
    End Sub
    '------------------------------------------------------------------------------------------'

End Class