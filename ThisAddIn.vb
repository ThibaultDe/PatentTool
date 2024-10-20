﻿Imports System.Diagnostics
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
        Return input.Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbCrLf, " ").Replace("’", "'").Replace(",", "").Replace(".", "").Replace(";", "").Trim()
    End Function

    Public Function ClearDuplicates(ByRef RefDict As Object) As Object
        'Supprime les doublons si deux refs sont contenues l'une dans l'autre
        For Each Key1 In RefDict.Keys
            Debug.Print(Key1)
            For Each Key2 In RefDict.Keys
                If RefDict.Exists(Key1) And (Not (Key1 = Key2)) And (Right(Key1, Len(Key2)) = Key2) Then
                    'Debug.Print "A", "key1", Key1, "key2", Key2'
                    RefDict.Remove(Key1)
                ElseIf (Not (Key1 = Key2)) And (Right(Key2, Len(Key1)) = Key1) Then
                    'Debug.Print "B", "key1", Key1, "key2", Key2'
                    RefDict.Remove(Key2)
                End If

            Next Key2
        Next Key1

        Return RefDict

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
            MsgBox("Revendications non trouvées")
            Return (Nothing)
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


    Function NumRefs()

        Dim RefDict = CreateObject("Scripting.Dictionary")
        Dim NumDict = CreateObject("Scripting.Dictionary")

        Dim ExeptionArray = New String() {"figure", "fig", "figures", "et", "ou", "environ", "d'environ", "moins", "exemple", "de", "entre", "=", "+", "-", "{", "[", ";", ",", "."}
        Dim DeterminantsArray = New String() {",", ";", "le", "la", "au", "l'", "les", "ce", "cette", "ces", "son", "sa", "ses", "leur", "leurs", "un", "une", "qu'un", "qu'une", "d'une", "d'un", "du", "des", "et", "chaque", "plusieurs", "deux", "trois", "quatre", "cinq", "six", "sept", "huit", "neuf", "dix", "vingt"}
        Dim ValNumArray = New String() {"que", "entre", "à", "et"}

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

                If (Not (IsNumeric(PotentialRef)) And (Not inArray(ExeptionArray, PotentialRef))) Then   'Si la Ref correspond à un mot clé  => pas un numéro, pas un mot comme "Figure", pas un mot qui annonce une valeur "plus de, moins de, exemple,..."
                    If Not inArray(ValNumArray, PotentialRef) Then 'et n'existe pas encore, on la crée
                        Dim Ref = New String() {" ", " ", " ", " ", " ", " "} 'Taille maximale : 6 mots => Permet d'éviter des bugs avec des Références trop longues

                        If PotentialRef.StartsWith("l'", StringComparison.OrdinalIgnoreCase) Then '=> Cas particulier pour gérer les mots en "l'"
                            Ref(5) = Right(PotentialRef, (Len(PotentialRef) - 2))
                        Else

                            Ref(5) = PotentialRef 'Remplis la référence en commencçant par son dernier mot
                            For k = 2 To 6 Step 1 'remonte la référence en regardant les mots précédents, jusqu'à tomber sur un déterminant.
                                Dim RefWord As String = wordsArray(i - k)

                                If (Not (IsNumeric(RefWord)) And Not inArray(DeterminantsArray, RefWord)) Then
                                    If RefWord.StartsWith("l'", StringComparison.OrdinalIgnoreCase) Then  '=> Cas particulier pour gérer les mots en "l'"
                                        Ref(6 - k) = Right(RefWord, (Len(RefWord) - 2))
                                        Exit For
                                    Else
                                        Ref(6 - k) = RefWord
                                    End If
                                Else
                                    Exit For
                                End If
                            Next k
                        End If


                        Dim refText As String = Trim(Join(Ref))   'Concatène la ref complète'

                        If (Len(refText) > 0) Then 'A TESTER => EST-CE ENCORE UTILE ?"
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
                End If
            End If
        Next i

        RefDict = ClearDuplicates(RefDict) 'Retire les doucblons

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
    End Function


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        wordApp = Me.Application
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Dim myUserControl1 As MyUserControl
    Dim myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

End Class
