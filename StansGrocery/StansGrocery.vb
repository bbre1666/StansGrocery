Option Strict On
Option Explicit On
Option Compare Text
Imports System.Security.Cryptography
Imports System.Threading
Imports System.Timers

'RCET0265
'BadenBrenner
'Spring 2023    
'Stans Grocery 
'https://github.com/bbre1666/StansGrocery.git


Public Class StansGrocery


    Dim RawDataArray() As String 'raw data array to pull data from file 
    Dim Food(0, 2) As String 'formated array to get the right frmat to display  
    Dim SearchValue As String() 'varable to hold the search item 



    Private Sub StansGroceryForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Hide() 'hide stans form on startup
        Splashscrean.Show() 'show the splash screan
        Thread.Sleep(3000) ' wait whale displaying splash screan
        Splashscrean.Close() 'close splash screan
        Me.Show() 'show stans form

        FindNewFileToolStripMenuItem_Click() 'call sub to open fine 

    End Sub



    Private Sub FindNewFileToolStripMenuItem_Click()


        EnableEverything(False) ' call sub to disable everething to prevent error imputs whale calling the file 

        Dim fileFound As Boolean = False 'set file found state

        Do Until fileFound

            If OpenFileDialog.ShowDialog <> Windows.Forms.DialogResult.Cancel Then 'open file dialog to find the file 
                Try
                    RawDataArray = Split(My.Computer.FileSystem.ReadAllText(OpenFileDialog.FileName), vbCrLf)
                    ReDim Food(RawDataArray.Length - 1, 2)
                    EnableEverything(True)
                    fileFound = True 'set file found state to true so loop will end 
                Catch ex As Exception
                    If MessageBox.Show("File Not Found, do you want to try again? (Choosing no will close the file)", "Try Again?", MessageBoxButtons.YesNo) = DialogResult.No Then
                        Splashscrean.Close()
                        Me.Close()
                        Exit Sub
                    End If
                End Try
            Else
                If MessageBox.Show("File Not Found, do you want to try again? (Choosing no will close the file)", "Try Again?", MessageBoxButtons.YesNo) = DialogResult.No Then
                    Splashscrean.Close()
                    Me.Close()
                    Exit Sub
                End If
            End If
        Loop

        FillListBox() ' calls sub to load file data t the list box 



    End Sub


    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        'action when search button is clicked 

        'stops program if "zzz" is searched 
        If SearchTextBox.Text.ToLower = "zzz" Then
            Splashscrean.Close()
            Me.Close()
        End If

        'Clears clears list box for new data 
        DisplayListBox.Items.Clear()


        Dim somethingFound As Boolean = False 'setting fond varable state 

        'search the array
        For i As Integer = 0 To RawDataArray.Length - 1

            'If the search character is in the data then add item, also compare method "text" is not case sensitive dur to the option compare being set to text
            If InStr(1, Food(i, 0), SearchTextBox.Text, CompareMethod.Text) > 0 Then

                DisplayListBox.Items.Add(Food(i, 0))
                somethingFound = True

            End If

        Next

        'sets the items in alfibitacle oder 
        DisplayListBox.Sorted = True

        'stores the search item 
        SearchValue = DisplayListBox.Items.OfType(Of String).ToArray

        'sets the radio buttons to aisle
        AisleRadioButton.Checked = True
    End Sub



    Private Sub FilterComboBox_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles FilterComboBox.SelectionChangeCommitted

        SearchValue = DisplayListBox.Items.OfType(Of String).ToArray

        DisplayListBox.Items.Clear() ' clears the list box


        If CStr(FilterComboBox.Items(FilterComboBox.SelectedIndex)) = "View All" Then 'convert to string to serch using the chosen filter
            DisplayListBox.Items.AddRange(SearchValue)

            If AisleRadioButton.Checked Then
                PartiallyFillComboBox(1) '1 refferences the location of the aisle data in the array
            Else
                PartiallyFillComboBox(2) '2 refferences the location of the catigoary data in the array
            End If
        Else
            For i As Integer = 0 To SearchValue.Length - 1 'filter by aisle or catogory
                For n As Integer = 0 To RawDataArray.Length - 1
                    If Food(n, 0) = SearchValue(i) And Food(n, 1) = CStr(FilterComboBox.Items(FilterComboBox.SelectedIndex)) Then
                        DisplayListBox.Items.Add(SearchValue(i))
                    ElseIf Food(n, 0) = SearchValue(i) And Food(n, 2) = CStr(FilterComboBox.Items(FilterComboBox.SelectedIndex)) Then
                        DisplayListBox.Items.Add(SearchValue(i))
                    End If
                Next
            Next
        End If

        DisplayListBox.Sorted = True  'sets the items in alfibitacle oder 

    End Sub

    Sub EnableEverything(enableEverything As Boolean) ' sub enalbles or disables user imputs to prevent erros 

        SearchGroupBox.Enabled = enableEverything
        MainMenuStrip.Enabled = enableEverything

    End Sub
    Sub FillListBox() 'fills the list box with data set 
        DisplayListBox.Items.Clear() ' clers list box for new data 
        Dim temp() As String
        For i As Integer = 0 To RawDataArray.Length - 1

            'processes the raw data in to the food array

            RawDataArray(i) = RawDataArray(i).Replace("""", "")
            temp = Split(RawDataArray(i), ",")


            For n As Integer = 0 To temp.Length - 1

                If temp(n).Contains("$$ITM") Then

                    temp(n) = temp(n).Replace("$$ITM", "")
                    If temp(n) = "" Then
                        temp(n) = "N/A"
                    End If
                    Food(i, 0) = temp(n) 'item name to array columm 0

                ElseIf temp(n).Contains("##LOC") Then

                    temp(n) = temp(n).Replace("##LOC", "")
                    If temp(n) = "" Then
                        temp(n) = "N/A"
                    End If
                    Food(i, 1) = temp(n) 'location stored to colum 1

                ElseIf temp(n).Contains("%%CAT") Then

                    temp(n) = temp(n).Replace("%%CAT", "")
                    If temp(n) = "" Then
                        temp(n) = "N/A"
                    End If
                    Food(i, 2) = temp(n) 'catogory to colum 2

                Else ' set empitys incase value is missing 

                    Food(i, 0) = "N/A"
                    Food(i, 1) = "N/A"
                    Food(i, 2) = "N/A"

                End If

            Next


            DisplayListBox.Items.Add(Food(i, 0)) 'displays items to list box 

        Next


        DisplayListBox.Sorted = True 'sets the items in alfibitacle oder 

    End Sub
    Sub PartiallyFillComboBox(aisleOrCategory As Integer) 'sets the options in teh combo box depending on the radio button selection

        FilterComboBox.Items.Clear() 'clears for new values 

        For i As Integer = 0 To DisplayListBox.Items.Count - 1
            For n As Integer = 0 To RawDataArray.Length - 1
                '
                If Food(n, 0) Is DisplayListBox.Items(i) And Not FilterComboBox.Items.Contains(Food(n, aisleOrCategory)) Then
                    If aisleOrCategory = 1 And Food(n, aisleOrCategory).Length < 2 And Not FilterComboBox.Items.Contains(" " & Food(n, aisleOrCategory)) Then

                        FilterComboBox.Items.Add(" " & Food(n, aisleOrCategory))
                    ElseIf Not FilterComboBox.Items.Contains(" " & Food(n, aisleOrCategory)) Then

                        FilterComboBox.Items.Add(Food(n, aisleOrCategory))
                    End If
                End If
            Next
        Next


        FilterComboBox.Sorted = True
        FilterComboBox.Sorted = False
        FilterComboBox.Items.Insert(0, "View All")

        FilterComboBox.SelectedIndex = 0

    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        SearchButton_Click(sender, e) 'cals search function
    End Sub

    Private Sub AisleRadioButton_Checked(sender As Object, e As EventArgs) Handles AisleRadioButton.CheckedChanged
        PartiallyFillComboBox(1) 'radio button filtering-aisle 

        FilterComboBox_SelectionChangeCommitted(sender, e)
    End Sub

    Private Sub CategoryRadioButton_Checked(sender As Object, e As EventArgs) Handles CategoryRadioButton.CheckedChanged
        PartiallyFillComboBox(2) 'radio buttom filtering-catogory

        FilterComboBox_SelectionChangeCommitted(sender, e)
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close() 'closes form
    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close() 'closes form
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Me.Hide() 'closes stans form
        AboutForm.Show() 'shows about form
    End Sub

End Class