Public Class AboutForm
    Private Sub AboutForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        logoPictureBox.SizeMode = PictureBoxSizeMode.StretchImage 'sets image size to display logo 


        logoPictureBox.Image = Image.FromFile("C:\Users\Baden Brenner\OneDrive\Documents\gethub\StansGrocery\StansGrocery\Resources\with componets.jpg") 'pulls logo from location
    End Sub

    Private Sub AboutForm_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        StansGrocery.Show() 'shows stans form again

    End Sub
End Class