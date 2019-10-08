Public Class TestCases1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim result As Integer

        '#############################################
        'REMARQUE UN SUPER TESTEUR C'EST EXCEL => IL EST CAPABLE DE CALCULER LA DIFFERENCE DE JOUR ENTRE 2 DATES
        '#############################################

        'jeu de test
        '26/09/2018 au 26/09/2019 => 
        'Number of days to the end of the year = 96
        'Number of days between 2 new year days = 0
        'Number of days since the new year day = 269
        'Difference between the two dates (in days) = 365

        '26/09/2017 au 26/09/2019 => 
        'Number of days to the end of the year = 96
        'Number of days between 2 new year days = 365
        'Number of days since the new year day = 269
        'Difference between the two dates (in days) = 730

        '25/09/2019 au 26/09/2019 => 
        'Difference between the two dates (in days) = 1


        '26/08/2019 au 26/09/2019 => 
        'Number of days to the end of the month = 5
        'Number of days between 2 new month days = 0
        'Number of days since the new month day = 26
        'Difference between the two dates (in days) = 31

        '20/01/2000 au 21/03/2015
        'Number of days to the end of the year = 346
        'Number of days between 2 new year days = 5112 
        'Number of days since the new year day = 79
        'Difference between the two dates (in days) = 5539

        Dim f1 As Form1 = New Form1()

        TextBox1.Text = TextBox1.Text & "Days between 26/09/2018 and 26/09/2019 => expected result = 365 "
        result = f1.go(26, 9, 2018, 26, 9, 2019)
        TextBox1.Text = TextBox1.Text & "=> result of the alogorith = " & result
        If result = 365 Then
            TextBox1.Text = TextBox1.Text & " ====> TEST CASE OK" & vbCrLf
        Else
            TextBox1.Text = TextBox1.Text & " ====> ERROR TEST CASE NOK !!!" & vbCrLf
        End If

        TextBox1.Text = TextBox1.Text & "Days between 26/09/2017 and 26/09/2019 => expected result = 730 "
        result = f1.go(26, 9, 2017, 26, 9, 2019)
        TextBox1.Text = TextBox1.Text & "=> result of the alogorith = " & result
        If result = 730 Then
            TextBox1.Text = TextBox1.Text & " ====> TEST CASE OK" & vbCrLf
        Else
            TextBox1.Text = TextBox1.Text & " ====> ERROR TEST CASE NOK !!!" & vbCrLf
        End If


    End Sub

    Private Sub TestCases1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class