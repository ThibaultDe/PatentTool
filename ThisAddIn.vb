Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports System.Windows.Input
Imports System.Linq
Imports System.Collections.Generic
Imports System.Drawing
Imports Word = Microsoft.Office.Interop.Word
Imports System.Drawing.Text


Public Class ThisAddIn

    Private WithEvents wordApp As Microsoft.Office.Interop.Word.Application

    Function inArray(myArray, myValue) 'Vérifie si une valeur est dans un Array'
        inArray = False

        If (myArray.Length > 0) Then
            For i = LBound(myArray) To UBound(myArray)
                If myArray(i) = myValue Then 'If value found
                    inArray = True
                    Exit For
                End If
            Next
        End If
    End Function

    Public Function CleanString(ByVal input As String) As String
        ' Remplacer les caractères indésirables par une chaîne vide
        Return input.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbCrLf, " ").Replace("’", "'").Replace(",", "").Replace(".", "").Replace(";", "").Replace("l'", "l ").Replace("(", "").Replace(")", "").Trim()
    End Function

    Public Function ClearDuplicates(ByRef Numrefs As Object) As Object

        For Each Key In Numrefs.Keys
            Dim RefArray = Numrefs(Key)
            Dim TempList As New List(Of Object)

            For i = UBound(RefArray) To LBound(RefArray) Step -1
                Dim Ref1 = RefArray(i).ToString()
                Debug.Print(Ref1)
                Dim IsDuplicate As Boolean = False
                For j = UBound(RefArray) To LBound(RefArray) Step -1
                    If i <> j Then
                        Dim Ref2 = RefArray(j).ToString()
                        If Right(Ref1, Len(Ref2)) = Ref2 Then
                            'Debug.Print("key1", Ref1, "key2", Ref2)
                            IsDuplicate = True
                            Exit For
                        End If
                    End If
                Next j
                ' Ajoute Ref1 à la liste temporaire s'il n'est pas un doublon
                If Not IsDuplicate Then TempList.Add(Ref1)
            Next i

            ' Remplace RefArray par une version filtrée sans doublons
            Numrefs(Key) = TempList.ToArray()
        Next Key

        Return Numrefs
    End Function

    Function GetDescriptionRange() As Range
        '--------------------- Réccupère le Range de la description des figures, entre les termes "Description" et "Revendications" en GRAS -------------------------'
        Dim myRange As Object
        Dim DescriptionStart As Object
        Dim DescriptionEnd As Object
        Dim DescriptionRange As Range

        myRange = Globals.ThisAddIn.Application.ActiveDocument.Range

        With myRange.Find 'Trouve le début de la description détaillée'
            .Text = "Description"
            .Font.Bold = True                   'Cherche le texte en GRAS'
            .Forward = False
            .Execute()
        End With

        If myRange.Find.Found = True Then
            DescriptionStart = myRange.Start
        Else
            MsgBox("Description non trouvée")
            Return (Nothing)
        End If

        myRange = Globals.ThisAddIn.Application.ActiveDocument.Range
        With myRange.Find
            .Text = "REVENDICATIONS"        'Cherche le texte en GRAS'
            .Font.Bold = True
            .Forward = False
            .Execute()
        End With

        If myRange.Find.Found = True Then
            DescriptionEnd = myRange.Start
        Else
            myRange = Globals.ThisAddIn.Application.ActiveDocument.Range
            With myRange.Find
                .Text = "CLAIMS"        'Cherche le texte en GRAS'
                .Font.Bold = True
                .Forward = False
                .Execute()
            End With
            If myRange.Find.Found = True Then
                DescriptionEnd = myRange.Start
            Else
                DescriptionEnd = Globals.ThisAddIn.Application.ActiveDocument.Range.End
            End If
        End If

        DescriptionRange = Globals.ThisAddIn.Application.ActiveDocument.Range(DescriptionStart, DescriptionEnd)
        Return DescriptionRange
    End Function


    Function PrintArray(myArray) 'Debug.Print un Array'
        For i = LBound(myArray) To UBound(myArray)
            Debug.Print(myArray(i))
            Debug.Print(" ")
        Next i
        Debug.Print(" ")
    End Function



    Function NumRefs(Language As String)

        Dim RefDict = CreateObject("Scripting.Dictionary")
        Dim NumDict = CreateObject("Scripting.Dictionary")

        Dim DeterminantsArray As Object
        Dim ExeptionArray As Object

        If Language = "en" Then
            ExeptionArray = New String() {"figure", "fig", "figs", "figures", "and", "or", "about", "approximately", "less", "example", "of", "than", "to", "between", "=", "+", "-", "{", "[", ";", ",", "."}
            DeterminantsArray = New String() {",", ";", "the", "a", "an", "this", "that", "these", "those", "his", "her", "its", "of", "their", "my", "your", "our", "some", "any", "each", "every", "many", "several", "few", "more", "less", "most", "same", "other", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "twenty", "hundred", "thousand", "in", "on", "at", "with", "by", "for", "to", "from", "over", "under", "near", "between", "among", "through", "without", "before", "after", "about", "around", "behind", "above", "below", "next to", "beyond", "beside", "said"}
        ElseIf Language = "fr" Then
            ExeptionArray = New String() {"figure", "fig", "figs", "figures", "et", "ou", "environ", "d'environ", "moins", "exemple", "de", "que", "entre", "à", "=", "+", "-", "{", "[", ";", ",", "."}
            DeterminantsArray = New String() {",", ";", "le", "la", "au", "aux", "l", "les", "ce", "cette", "ces", "son", "sa", "ses", "leur", "leurs", "un", "une", "qu'un", "qu'une", "d'une", "d'un", "du", "des", "et", "chaque", "même", "mêmes", "autre", "autres", "plusieurs", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix", "vingt"}
        End If

        Dim N As Object

        Dim DescriptionRange As Range = GetDescriptionRange() ' Réccupère le range de la description des figures
        If DescriptionRange Is Nothing Then
            Return RefDict 'Si la description n'a pas été trouvée, retourner NumDict tel quel'
        End If


        Dim FullText As String = CleanString(LCase(DescriptionRange.Text)) 'Nettoye pour enlever la ponctuation et les changements de ligne
        Dim wordsArray As String() = Split(FullText) 'Réccupère le texte, et le divise en Array de mots


        For i = 3 To UBound(wordsArray) 'Itère sur l'ensemble des mots du texte
            Dim aWord As String = wordsArray(i)
            If (IsNumeric(aWord)) Then  ' Repère les références numériques
                Dim PotentialNum As String = aWord
                Dim PotentialRef As String = wordsArray(i - 1) 'Récupère le dernier mot de la Ref associée

                If (Not inArray(ExeptionArray, PotentialRef)) And Not (IsNumeric(PotentialRef)) Then   'Si la Ref correspond à un mot clé  => pas un numéro, pas un mot comme "Figure", pas un mot qui annonce une valeur "plus de, moins de, exemple,..."
                    Dim Ref = New String() {" ", " ", " ", " ", " ", " "} 'Taille maximale : 6 mots => Permet d'éviter des bugs avec des Références trop longues

                    Ref(5) = PotentialRef 'Remplis la référence en commencçant par son dernier mot
                    For k = 2 To 6 Step 1 'remonte la référence en regardant les mots précédents, jusqu'à tomber sur un déterminant.
                        Dim RefWord As String = wordsArray(i - k)
                        If (Not (IsNumeric(RefWord)) And Not inArray(DeterminantsArray, RefWord)) Then
                            Ref(6 - k) = RefWord
                        Else
                            Exit For
                        End If
                    Next k

                    Dim refText As String = Trim(Join(Ref))   'Concatène la ref complète'

                    If (Not RefDict.Exists(refText)) Then          'Si cette référence textuelle n'existe pas encore, la crée et lui associe le numéro de référence'
                        Dim numArray = New String() {PotentialNum}
                        RefDict.Add(Key:=refText, Item:=numArray)

                    ElseIf ((RefDict.Exists(refText)) And (Not inArray(RefDict(refText), PotentialNum))) Then 'Si cette référence existe avec un numéro différent, ajoute ce numéro à la référence'
                        Dim numArray = RefDict(refText)
                        N = UBound(numArray) + 1
                        ReDim Preserve numArray(N)
                        numArray(UBound(numArray)) = PotentialNum
                        RefDict(refText) = numArray
                    End If
                End If
            End If
        Next i

        'Retourne pour avoir un tableau des numéros avec les références textuelles associées
        For Each Key In RefDict.Keys
            Dim numArray = RefDict(Key)
            For Each Num In numArray
                If Not (NumDict.Exists(Num)) Then
                    Dim RefArray = New String() {Key}
                    NumDict.Add(Key:=Num, Item:=RefArray)
                ElseIf (Not inArray(NumDict(Num), Key)) Then
                    Dim RefArray = NumDict(Num)
                    N = UBound(RefArray) + 1
                    ReDim Preserve RefArray(N)
                    RefArray(UBound(RefArray)) = Key
                    NumDict(Num) = RefArray
                End If
            Next Num
        Next Key
        NumRefs = NumDict


        ClearDuplicates(NumRefs) 'Retire les doucblons
    End Function


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        wordApp = Me.Application
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Dim myUserControl1 As MyUserControl
    Dim myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

End Class
