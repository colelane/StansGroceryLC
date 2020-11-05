Imports System.Text.RegularExpressions
Public Class StansGroceryForm
    'Dim i As Integer
    Private Sub StansGroceryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()
        SplashScreenForm.BackgroundImageLayout = ImageLayout.Stretch
        SplashScreenForm.BackgroundImage = My.Resources.stansDraft2
        SplashScreenForm.Size = Me.Size
        SplashScreenForm.Show()
        Me.Show()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        SplashScreenForm.Hide()

    End Sub

    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click

        Dim match As Match
        Dim arr As String() = SplitWords(My.Resources.Grocery)



        Dim sizer, sizer2 As Integer

        For i = 0 To UBound(arr)
            match = Regex.Match(arr(i), "ITM")
            If match.Success = True Then
                sizer += 1
            End If
        Next

        Dim storeItems(sizer) As String
        sizer2 = 0
        For p = 0 To UBound(arr)
            match = Regex.Match(arr(p), "ITM")
            If match.Success = True Then
                storeItems(sizer2) = arr(p)
                sizer2 += 1
            End If
        Next

        For j = 0 To (UBound(storeItems) - 1)
            storeItems(j) = Regex.Replace(storeItems(j), "Ω", "/")
            storeItems(j) = Regex.Replace(storeItems(j), "ITM", String.Empty)
            DisplayListBox.Items.Add(storeItems(j))
        Next
        'i += 1
    End Sub
    Private Function SplitWords(ByVal s As String) As String()
        Dim initialSplit As String = Regex.Replace(s, "/", "Ω")
        Return Regex.Split(initialSplit, "\p{Po}|\p{Sc}")
        'best pattern so far \W+(?!ITM)\w+\b
    End Function

End Class
