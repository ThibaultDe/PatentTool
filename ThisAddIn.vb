Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Public Class ThisAddIn
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



    Function PrintArray(myArray) 'Debug.Print un Array'
        For i = LBound(myArray) To UBound(myArray)
            Debug.Print(myArray(i))
            Debug.Print(" ")
        Next i
        Debug.Print(" ")
    End Function


    Function NumRefs()

        Dim i As Integer
        Dim myHeadings As Object
        Dim aWord As String
        Dim DescriptionRange As Range

        Dim DescriptionStart As Object
        Dim DescriptionEnd As Object

        Dim PotentialNum As Object
        Dim PotentialRef As String
        Dim RefWord As String
        Dim refText As String
        Dim PrecWord As String
        Dim PrecPrecWord As String

        Dim N As Object
        Dim Nums As Object
        Dim Refs As Object

        Dim RefDict = CreateObject("Scripting.Dictionary")
        Dim NumDict = CreateObject("Scripting.Dictionary")

        '--------------------- Range de la description des figures  -------------------------'
        Dim myRange = Globals.ThisAddIn.Application.ActiveDocument.Range
        With myRange.Find 'Trouve le début de la description détaillée'
            .Text = "Description détaillée"
            .Font.Bold = True                   'Cherche le texte en GRAS'
            .Forward = False
            .Execute()
        End With

        If myRange.Find.Found = True Then
            DescriptionStart = myRange.Start
            'System.Diagnostics.Debug.WriteLine("Description Start trouvé ")
        Else
            MsgBox("Description détaillée non trouvée")
            Return (NumDict)
        End If

        myRange = Globals.ThisAddIn.Application.ActiveDocument.Range
        With myRange.Find
            .Text = "REVENDICATIONS"        'Cherche le texte en GRAS'
            .Font.Bold = True
            .Forward = False
            .Execute()
        End With

        If myRange.Find.Found = True Then
            'System.Diagnostics.Debug.WriteLine("Description End trouvé ")
            DescriptionEnd = myRange.Start
        Else
            MsgBox("Revendications non trouvées")
            Return (NumDict)
        End If


        '--------------------- Recherche des numéros dans la description -------------------------'
        DescriptionRange = Globals.ThisAddIn.Application.ActiveDocument.Range(DescriptionStart, DescriptionEnd)





        Dim ExeptionArray = New String() {"figure", "fig", "figures", "et", "ou", "environ", "d'environ", "moins", "de", "entre", "=", "+", "-", "{", "[", ";", ",", "."}
        Dim DeterminantsArray = New String() {",", ";", "le", "la", "au", "l'", "les", "ce", "cette", "ces", "son", "sa", "ses", "leur", "leurs", "un", "une", "qu'un", "qu'une", "d'une", "d'un", "du", "des", "et", "chaque", "plusieurs", "pluralité de", "succession de", "deux", "trois", "quatre", "cinq", "dix", "vingt"}
        'Dim GroupsArray = New String() {} '"pluralité", "groupe", "succession"}
        Dim ValNumArray = New String() {"que", "entre", "à", "et"}


        With DescriptionRange.Find
            .Text = "’"
            .Replacement.Text = "'"
            .Forward = True
            .Wrap = Word.WdFindWrap.wdFindContinue
            .Format = False ' Ne change pas le formatage
            .Execute(Replace:=Word.WdReplace.wdReplaceAll)
        End With

        Dim FullText As String = LCase(DescriptionRange.Text)
        Dim wordsArray As String() = Split(FullText)

        For i = 0 To UBound(wordsArray)
            wordsArray(i) = Trim(Replace(Replace(Replace(wordsArray(i), ",", ""), ".", ""), ";", ""))
        Next i


        For i = 3 To UBound(wordsArray)

            aWord = wordsArray(i)

            If (IsNumeric(aWord) And (Not DescriptionRange.Words(i).Font.Italic)) Then

                PotentialNum = aWord
                PotentialRef = wordsArray(i - 1)

                If Not inArray(ExeptionArray, PotentialRef) Then   'Si la Ref correspond à un mot clé  
                    If Not inArray(ValNumArray, PotentialRef) Then 'et n'existe pas encore, on la crée
                        Dim Ref = New String() {" ", " ", " ", " ", " ", " "} 'Taille maximale : 6 mots

                        If PotentialRef.StartsWith("l'", StringComparison.OrdinalIgnoreCase) Then
                            Ref(5) = Right(PotentialRef, (Len(PotentialRef) - 2))
                            'Debug.Print(Ref(5))

                        Else
                            Ref(5) = PotentialRef
                            For k = 2 To 6 Step 1
                                RefWord = wordsArray(i - k)
                                If (Not inArray(DeterminantsArray, RefWord) And Not (IsNumeric(RefWord))) Then
                                    If RefWord.StartsWith("l'", StringComparison.OrdinalIgnoreCase) Then
                                        Ref(6 - k) = Right(RefWord, (Len(RefWord) - 2))
                                        Exit For
                                        'ElseIf (RefWord = "de") Then 
                                        '    PrecWord = LCase(Trim(DescriptionRange.Words(i - k - 1).Text))
                                        '    PrecPrecWord = LCase(Trim(DescriptionRange.Words(i - k - 2).Text))
                                        '    If inArray(GroupsArray, PrecWord) And inArray(DeterminantsArray, PrecPrecWord) Then
                                        '        Exit For
                                        '    Else
                                        '        Ref(6 - k) = RefWord
                                        '    End If
                                    Else
                                        Ref(6 - k) = RefWord
                                    End If
                                Else
                                    Exit For
                                End If
                            Next k
                        End If


                        refText = Trim(Join(Ref))   'On concatène la ref complète'

                        If (Len(refText) > 0) Then
                            If (Not RefDict.Exists(refText)) Then
                                Dim numArray = New String() {PotentialNum}
                                RefDict.Add(Key:=refText, Item:=numArray)
                                Debug.Print(refText + " Ajouté avec " + PotentialNum)


                            ElseIf ((RefDict.Exists(refText)) And (Not inArray(RefDict(refText), PotentialNum))) Then
                                Dim numArray = RefDict(refText)
                                N = UBound(numArray) + 1
                                ReDim Preserve numArray(N)
                                numArray(UBound(numArray)) = PotentialNum

                                RefDict(refText) = numArray
                                Debug.Print(PotentialNum & " Ajouté à la ref " & refText)
                            End If

                        End If

                    End If

                End If
            End If

        Next i

        'NumRefs = RefDict

        'On supprime les doublons si deux refs sont contenues l'une dans l'autre
        For Each Key1 In RefDict.Keys
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

        For Each Key In RefDict.Keys
            Debug.Print(Key)
            Dim numArray = RefDict(Key)
            For Each Num In numArray
                Debug.Print(Num)
                If Not (NumDict.Exists(Num)) Then
                    Dim RefArray = New String() {Key}
                    NumDict.Add(Key:=Num, Item:=RefArray)
                    Debug.Print(Key, " added")
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

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Dim myUserControl1 As MyUserControl
    Dim myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

End Class
