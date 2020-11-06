﻿Imports System.Text.RegularExpressions
Public Class StansGroceryForm

    Public Sub StansGroceryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        SplashScreenForm.BackgroundImageLayout = ImageLayout.Stretch
        SplashScreenForm.BackgroundImage = My.Resources.stansDraft2
        SplashScreenForm.Size = Me.Size
        SplashScreenForm.Show()
        Me.Show()


        Dim match As Match
        Dim arr As String() = SplitWords(My.Resources.Grocery)


        Dim foodSizer, foodSizer2, locSizer, locSizer2, catSizer, catSizer2 As Integer





        'Items
        For i = 0 To UBound(arr)
            match = Regex.Match(arr(i), "ITM")
            If match.Success = True Then
                foodSizer += 1
            End If
        Next
        Dim finalArr(foodSizer, 2) As String


        Dim items(foodSizer) As String
        foodSizer2 = 0
        'For p = 0 To UBound(arr)
        '    match = Regex.Match(arr(p), "ITM")
        '    If match.Success = True Then
        '        items(foodSizer2) = arr(p)
        '        foodSizer2 += 1
        '    End If
        'Next

        'Final Array Loader
        For p = 0 To UBound(arr)
            match = Regex.Match(arr(p), "ITM")
            If match.Success = True Then
                finalArr(foodSizer2, 0) = arr(p)
                foodSizer2 += 1
            End If
        Next

        'LOCS
        For l = 0 To UBound(arr)
            match = Regex.Match(arr(l), "LOC")
            If match.Success = True Then
                locSizer += 1
            End If
        Next

        Dim locs(locSizer) As String
        locSizer2 = 0
        'For k = 0 To UBound(arr)
        '    match = Regex.Match(arr(k), "LOC")
        '    If match.Success = True Then
        '        locs(locSizer2) = arr(k)
        '        locSizer2 += 1
        '    End If
        'Next

        'Final Array Loader
        For k = 0 To UBound(arr)
            match = Regex.Match(arr(k), "LOC")
            If match.Success = True Then
                finalArr(locSizer2, 1) = arr(k)
                locSizer2 += 1
            End If
        Next

        'CATS
        For y = 0 To UBound(arr)
            match = Regex.Match(arr(y), "LOC")
            If match.Success = True Then
                catSizer += 1
            End If
        Next

        Dim cats(catSizer) As String
        catSizer2 = 0
        'For r = 0 To UBound(arr)
        '    match = Regex.Match(arr(r), "CAT")
        '    If match.Success = True Then
        '        cats(catSizer2) = arr(r)
        '        catSizer2 += 1
        '    End If
        'Next

        'Final Array Loader
        For r = 0 To UBound(arr)
            match = Regex.Match(arr(r), "CAT")
            If match.Success = True Then
                finalArr(catSizer2, 2) = arr(r)
                catSizer2 += 1
            End If
        Next



        For j = 0 To (UBound(items) - 1)
            For p = 0 To 2

                finalArr(j, p) = Regex.Replace(finalArr(j, p), "Ω", "/")
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "ITM", String.Empty)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "LOC", String.Empty)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "CAT", String.Empty)
                DisplayListBox.Items.Add(finalArr(j, p))
            Next



            'items(j) = Regex.Replace(items(j), "Ω", "/")
            'items(j) = Regex.Replace(items(j), "ITM", String.Empty)
            'locs(j) = Regex.Replace(locs(j), "LOC", String.Empty)
            'cats(j) = Regex.Replace(locs(j), "CAT", String.Empty)
        Next
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        SplashScreenForm.Hide()

    End Sub

    Private Function SplitWords(ByVal s As String) As String()
        Dim initialSplit As String = Regex.Replace(s, "/", "Ω")
        Return Regex.Split(initialSplit, "\p{P}|\p{Sc}")
        'best pattern so far \W+(?!ITM)\w+\b
    End Function


End Class
