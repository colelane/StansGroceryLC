Option Compare Text
Imports System.Text.RegularExpressions
Public Class StansGroceryForm
    Dim sizer, foodSizer, locSizer, catSizer As Integer
    Dim finalarr2(255, 2) As String

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        DisplayListBox.Items.Clear()
        For a = 0 To 255
            For b = 0 To 2
                If ComboBox1.Text = finalarr2(a, b) Then
                    DisplayListBox.Items.Add("You will find " & finalarr2(a, b) & " on aisle " & finalarr2(a, b + 1) & " with the " & finalarr2(a, b + 2))
                End If
            Next
        Next
    End Sub

    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        DisplayListBox.Items.Clear()
        For a = 0 To 255
            For b = 0 To 2
                If TextBox1.Text = finalarr2(a, b) Then
                    DisplayListBox.Items.Add("You will find " & finalarr2(a, b) & " on aisle " & finalarr2(a, b + 1) & " with the " & finalarr2(a, b + 2))
                End If
            Next
        Next
        If DisplayListBox.Items.Count = 0 Then
            DisplayListBox.Items.Add("Item could not be found")
        End If
    End Sub

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

        'Food
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
            Next
            ComboBox1.Items.Add(finalArr(j, 0))
        Next
        Array.Copy(finalArr, finalarr2, finalArr.Length)
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
