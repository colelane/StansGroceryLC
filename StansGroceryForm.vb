Imports System.Text.RegularExpressions
Public Class StansGroceryForm
    Dim sizer, foodSizer, locSizer, catSizer As Integer
    Dim finalarr2(255, 2) As String

    Public Sub StansGroceryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        SplashScreenForm.BackgroundImageLayout = ImageLayout.Stretch
        SplashScreenForm.BackgroundImage = My.Resources.stansDraft2
        SplashScreenForm.Size = Me.Size
        SplashScreenForm.Show()
        Me.Show()

        Dim match As Match
        Dim initialArray As String() = SplitWords(My.Resources.Grocery)

        'Alphabetizer section.
        Dim initialArrStr As String = String.Join("", initialArray)
        Dim alphabetizer() As String
        alphabetizer = Regex.Split(initialArrStr, vbLf)
        Array.Sort(alphabetizer)
        Dim sortedStr As String = String.Join("", alphabetizer)
        Dim arr() As String

        'Uses a zero width positive lookbehind assertion to split the array back 
        'into Single lines to prepare for matching
        arr = Regex.Split(sortedStr, "(?=ITM)|(?=LOC)|(?=CAT)")


        For i = 0 To UBound(arr)
            match = Regex.Match(arr(i), "ITM")
            If match.Success = True Then
                sizer += 1
            End If
        Next
        Dim finalArr(sizer - 1, 2) As String

        foodSizer = 0
        For p = 0 To UBound(arr)
            match = Regex.Match(arr(p), "ITM")
            If match.Success = True Then
                finalArr(foodSizer, 0) = arr(p)
                foodSizer += 1
            End If
        Next

        'LOC
        locSizer = 0
        For k = 0 To UBound(arr)
            match = Regex.Match(arr(k), "LOC")
            If match.Success = True Then
                finalArr(locSizer, 1) = arr(k)
                locSizer += 1
            End If
        Next

        'Categories
        catSizer = 0
        For r = 0 To UBound(arr)
            match = Regex.Match(arr(r), "CAT")
            If match.Success = True Then
                finalArr(catSizer, 2) = arr(r)
                catSizer += 1
            End If
        Next




        For j = 0 To (sizer - 1)
            For p = 0 To 2

                finalArr(j, p) = Regex.Replace(finalArr(j, p), "Ω", "/")
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "ITM", String.Empty)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "LOC", String.Empty)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "CAT", String.Empty)
                DisplayListBox.Items.Add(finalArr(j, p))
            Next
        Next
        Array.Copy(finalArr, finalarr2, finalArr.Length)

        Console.Read()
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        SplashScreenForm.Hide()

    End Sub

    Private Function SplitWords(ByVal s As String) As String()
        Dim initialSplit As String = Regex.Replace(s, "/", "Ω")
        Return Regex.Split(initialSplit, "\p{P}|\p{Sc}")

        'actual best pattern \p{P}|\p{Sc}
    End Function


End Class
