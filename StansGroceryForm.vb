Option Compare Text
Option Strict On
Option Explicit On
'Lane Coleman
'RCET 0265
'Fall 2020
'Stans Grocery
'https://github.com/colelane/StansGroceryLC.git

Imports System.Text.RegularExpressions
Public Class StansGroceryForm
    Dim finalarr2(255, 2), sortedLocs(16), sortedCats(23) As String ' Array should be food$() - TJR

    Public Sub StansGroceryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'This group of text handles the initial splash screen.
        Timer1.Start()
        SplashScreenForm.BackgroundImageLayout = ImageLayout.Stretch
        SplashScreenForm.BackgroundImage = My.Resources.stansDraft2
        SplashScreenForm.Size = Me.Size
        SplashScreenForm.Show()
        Me.Show()

        Dim sizer, foodSizer, locSizer, catSizer As Integer
        Dim match As Match
        '/ counts as punction in the below regex.split, so it must be changed prior to the split to preserve it.  Change it back later.
        Dim initialSplit As String = Regex.Replace(My.Resources.Grocery, "/", "Ω")
        'Below: uses a regex.split function with a pattern to set punctuation and money symbols as delimiters.
        Dim initialArray As String() = Regex.Split(initialSplit, "\p{P}|\p{Sc}")

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

        'This section sizes the array columns and loads them with information from the first arr.
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
        Console.Read()
        'Categories
        catSizer = 0
        For r = 0 To UBound(arr)
            match = Regex.Match(arr(r), "CAT")
            If match.Success = True Then
                finalArr(catSizer, 2) = arr(r)
                catSizer += 1
            End If
        Next

        'This puts the /'s back where they belong, and removes the tags.
        For j = 0 To (sizer - 1)
            For p = 0 To 2
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "Ω", "/")
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "ITM", String.Empty)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "LOC", String.Empty).PadLeft(2)
                finalArr(j, p) = Regex.Replace(finalArr(j, p), "CAT", String.Empty)
            Next
            DisplayListBox.Items.Add(finalArr(j, 0))
        Next
        'moved the array's information to a global array, this array needs global access

        finalarr2 = finalArr
        LocSorter()
        CatSorter()

        FilterComboBox.SelectedItem = "Show All"
    End Sub
    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click, FileMenuSearch.Click, ContextMenuSearch.Click
        'Handles the search button clicks.  Checks the search text box and loads the display list box based on matched to the searched text.
        Dim goodData As Boolean
        Dim searchMatch As Match
        goodData = False
        DisplayListBox.Items.Clear()
        DisplayLabel.Text = String.Empty
        'Too many items come up when only one character is used.
        If SearchTextBox.TextLength = 1 Then
            DisplayLabel.Text = "Please be more specific."
            Exit Sub
        ElseIf SearchTextBox.Text = "zzz" Then
            Me.Close()
        End If
        'Matches only occur on the front end of words in the strings, instead of pulling matches out of the center of words.
        For a = 0 To UBound(finalarr2) - 1
            searchMatch = Regex.Match(finalarr2(a, 0), "\b" & SearchTextBox.Text, RegexOptions.IgnoreCase)
            If searchMatch.Success = True Then
                DisplayListBox.Items.Add(finalarr2(a, 0))
                goodData = True
            End If
        Next

        If goodData = False Then
            DisplayLabel.Text = $"Sorry, no matches for {SearchTextBox.Text}"
        End If

        DisplayListBox.Items.Remove("  ")
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub DisplayListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DisplayListBox.SelectedIndexChanged
        'This sub loads the display label with information, based on the users choice.
        For a = 0 To 255
            For b = 0 To 2
                If DisplayListBox.SelectedItem.ToString = finalarr2(a, b) Then
                    DisplayLabel.Text = "You will find " & finalarr2(a, b) & " on aisle " &
                        finalarr2(a, b + 1) & " with the " & finalarr2(a, b + 2)
                End If
            Next
        Next
    End Sub

    Private Sub RadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles FilterByAisleButton.CheckedChanged, FilterByCategoryButton.CheckedChanged
        'This sub watches for the radio buttons to change, then storts the filter combox box based on the users selection.
        If FilterByAisleButton.Checked = True Then
            FilterComboBox.Items.Clear()
            FilterComboBox.Items.Add("Show All")
            FilterComboBox.Items.Add("Choose Aisle...")
            FilterComboBox.SelectedItem = "Choose Aisle..."
            For t = 0 To UBound(sortedLocs)
                FilterComboBox.Items.Add(sortedLocs(t))
            Next
        Else
            FilterComboBox.Items.Clear()
            FilterComboBox.Items.Add("Show All")
            FilterComboBox.Items.Add("Choose Category...")
            FilterComboBox.SelectedItem = "Choose Category..."
            For l = 0 To UBound(sortedLocs)
                FilterComboBox.Items.Add(sortedCats(l))
            Next
        End If
        FilterComboBox.Items.Remove("  ")
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub FilterComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FilterComboBox.SelectedIndexChanged
        'This sub watches for the users selection within the filter combo box, adjusts the display list box accordingly.
        FilterComboBox.SelectedItem.ToString()
        DisplayListBox.Items.Clear()

        For a = 0 To 255
            If FilterComboBox.SelectedItem.ToString() = "Show All" Then
                DisplayListBox.Items.Add(finalarr2(a, 0))
            End If
            For b = 0 To 2
                If FilterComboBox.SelectedItem.ToString() = finalarr2(a, b) Then
                    DisplayListBox.Items.Add(finalarr2(a, 0))
                End If
            Next
        Next
        DisplayListBox.Items.Remove("  ")
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Timer sub, for the splash screen.
        Timer1.Stop()
        SplashScreenForm.Hide()
    End Sub

    Sub LocSorter()
        Dim locs(UBound(finalarr2)) As String

        For a = 0 To UBound(finalarr2)
            locs(a) = finalarr2(a, 1)
        Next
        Dim preDedupe As String = String.Join(",", locs)
        Dim dedupe As String = DeDupeinator(preDedupe)
        sortedLocs = Regex.Split(dedupe, ",")

        Array.Sort(sortedLocs)
        Console.Read()
    End Sub

    Private Sub AboutTopMenuItem_Click(sender As Object, e As EventArgs) Handles AboutTopMenuItem.Click
        CanadianAboutForm.BackgroundImageLayout = ImageLayout.Stretch
        CanadianAboutForm.BackgroundImage = My.Resources.canada
        CanadianAboutForm.Size = Me.Size
        CanadianAboutForm.Show()
    End Sub

    Sub CatSorter()
        Dim cats(UBound(finalarr2)) As String

        For a = 0 To UBound(finalarr2)
            cats(a) = finalarr2(a, 2)
        Next
        Dim preDedupe As String = String.Join(",", cats)
        Dim dedupe As String = DeDupeinator(preDedupe)
        sortedCats = Regex.Split(dedupe, ",")
        Array.Sort(sortedCats)
        Console.Read()
    End Sub
    Function DeDupeinator(ByVal sInput As String, Optional ByVal sDelimiter As String = ",") As String
        'This function is used for the location and category sorters to remove duplicates and return the individual instances once each.
        Dim varSection As String
        Dim sTemp As String

        For Each varSection In Split(sInput, sDelimiter)
            If InStr(1, sDelimiter & sTemp & sDelimiter, sDelimiter & varSection & sDelimiter, vbTextCompare) = 0 Then
                sTemp = sTemp & sDelimiter & varSection
            End If
        Next varSection

        DeDupeinator = Mid(sTemp, Len(sDelimiter) + 1)

    End Function

End Class
