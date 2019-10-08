Public Class Form1

    Dim daysInAMonth() As Integer = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim day1 As Integer
        Dim month1 As Integer
        Dim year1 As Integer

        Dim day2 As Integer
        Dim month2 As Integer
        Dim year2 As Integer

        Dim nbOfDaysBetweenThe2Dates As Integer
        TextBox1.Text = ""


        'day1 = DateTimePicker1.Value.Day
        'month1 = DateTimePicker1.Value.Month
        'year1 = DateTimePicker1.Value.Year

        'day2 = DateTimePicker2.Value.Day
        'month2 = DateTimePicker2.Value.Month
        'year2 = DateTimePicker2.Value.Year

        'TextBox1.Text = DateTimePicker1.Text

        'With DateTimePicker, the problem is solved with 5 code lines
        Dim d1 As DateTime = DateTimePicker1.Value
        Dim d2 As DateTime = DateTimePicker2.Value
        'Dim result As TimeSpan = d2.Subtract(d1)
        'Dim days As Integer = result.TotalDays
        'Label_result.Text = days

        Dim date1 As String = TextBox2.Text
        Dim date2 As String = TextBox3.Text

        If checkFormatTextBoxes(date1) Then
            If checkFormatTextBoxes(date2) Then
                Dim line1() As String = date1.Split("/")
                Dim line2() As String = date2.Split("/")
                day1 = line1(0)
                month1 = line1(1)
                year1 = line1(2)
                day2 = line2(0)
                month2 = line2(1)
                year2 = line2(2)

                If (inputControl(day1, month1, year1, day2, month2, year2)) Then

                    'Update the display
                    Date.TryParse("" & day1 & "/" & month1 & "/" & year1, d1)
                    Date.TryParse("" & day2 & "/" & month2 & "/" & year2, d2)
                    DateTimePicker1.Value = d1
                    DateTimePicker2.Value = d2
                    TextBox2.Text = day1 & "/" & month1 & "/" & year1
                    TextBox3.Text = day2 & "/" & month2 & "/" & year2

                    nbOfDaysBetweenThe2Dates = go(day1, month1, year1, day2, month2, year2)
                End If

            End If
        End If

    End Sub

    Private Function checkFormatTextBoxes(ByVal date1 As String) As Boolean

        Dim str_MessageErreurFormat As String = ""
        Dim cmpt As Integer = 0
        Dim bDateFormatOK As Boolean = True

        Dim aChar() As Char = date1.ToCharArray()

        For i = 0 To date1.Length - 1
           
            If aChar(i) = "0" Or aChar(i) = "1" Or aChar(i) = "2" Or aChar(i) = "3" Or aChar(i) = "4" Or aChar(i) = "5" Or aChar(i) = "6" Or aChar(i) = "7" Or aChar(i) = "8" Or aChar(i) = "9" Or aChar(i) = "/" Then
                If aChar(i) = "/" Then
                    cmpt = cmpt + 1
                    If aChar(date1.Length - 1) = "/" Then
                        str_MessageErreurFormat = "Enter a correct format for the dates => dd/mm/yyyy !!!"
                        bDateFormatOK = False
                    End If
                End If
                'nothing to do
            Else
                str_MessageErreurFormat = "Enter a correct format for the dates => dd/mm/yyyy !!!"
                bDateFormatOK = False
            End If
        Next
        If cmpt <> 2 Then
            str_MessageErreurFormat = "Enter a correct format for the dates => dd/mm/yyyy !!!"
            bDateFormatOK = False
        End If

        TextBox1.Text = str_MessageErreurFormat
        Return bDateFormatOK

    End Function


    Private Function inputControl(ByRef j1 As Integer, ByRef m1 As Integer, ByRef a1 As Integer, ByRef j2 As Integer, ByRef m2 As Integer, ByRef a2 As Integer) As Boolean

        'Variables
        Dim bDateValide1 As Boolean = True
        Dim bDateValide2 As Boolean = True
        Dim str_MessageErreur As String = ""
        Dim aNbJoursMois() As Integer = {0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
        Dim aNomMois() As String = {"", "Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Décembre"}


        ' Vérification
        If ((j1 >= 1) And (j1 <= 31)) Then
            If ((m1 >= 1) And (m1 <= 12)) Then
                If (j1 <= aNbJoursMois(m1)) Then
                    If ((m1 = 2) And (j1 = 29) And (iSBissextile(a1) = False)) Then
                        str_MessageErreur = str_MessageErreur & "L'année " & a1 & " n'était pas bissextile ! (date 1)" & vbCrLf
                    End If
                Else
                    bDateValide1 = False
                    str_MessageErreur = str_MessageErreur & "Il n'y a que " & aNbJoursMois(m1) & " jours dans le mois de " & aNomMois(m1) & " ! (date 1)" & vbCrLf

                End If
            Else
                bDateValide1 = False
                str_MessageErreur = str_MessageErreur & "Il y a 12 mois dans une année ! (date 1)" & vbCrLf
            End If
        Else
            bDateValide1 = False
            str_MessageErreur = str_MessageErreur & "Un mois a 1 jour minimum et 31 jours maximum ! (date 1)" & vbCrLf
        End If
        If ((j2 >= 1) And (j2 <= 31)) Then
            If ((m2 >= 1) And (m2 <= 12)) Then
                If (j2 <= aNbJoursMois(m2)) Then

                    If ((m2 = 2) And (j2 = 29) And (iSBissextile(a2) = False)) Then
                        str_MessageErreur = str_MessageErreur & "L'année " & a2 & " n'était pas bissextile ! (date 2)" & vbCrLf
                    End If
                Else
                    bDateValide2 = False
                    str_MessageErreur = str_MessageErreur & "Il n'y a que " & aNbJoursMois(m2) & " jours dans le mois de " & aNomMois(m2) & " ! (date 2)" & vbCrLf

                End If
            Else
                bDateValide2 = False
                str_MessageErreur = str_MessageErreur & "Il y a 12 mois dans une année ! (date 2)" & vbCrLf
            End If
        Else
            bDateValide2 = False
            str_MessageErreur = str_MessageErreur & "Un mois a 1 jour minimum  et 31 jours maximum ! (date 2)" & vbCrLf
        End If



        If ((a1 >= 1900) And (a1 <= 2100)) Then
            'c bon
        Else
            bDateValide1 = False
            str_MessageErreur = str_MessageErreur & "Année comprise entre 1900 et 2100 ! (date 1)" & vbCrLf
        End If
        If ((a2 >= 1900) And (a2 <= 2100)) Then
            'c bon
        Else
            bDateValide2 = False
            str_MessageErreur = str_MessageErreur & "Année comprise entre 1900 et 2100 ! (date 2)" & vbCrLf
        End If

        Inversion(j1, m1, a1, j2, m2, a2)
        TextBox1.Text = str_MessageErreur

        Return bDateValide1 And bDateValide2

    End Function


    Public Sub Inversion(ByRef ijour1 As Integer, ByRef imois1 As Integer, ByRef iannee1 As Integer, ByRef ijour2 As Integer, ByRef imois2 As Integer, ByRef iannee2 As Integer)

        Dim bInversion As Boolean = False
        Dim iTemp As Integer

        If (iannee1 < iannee2) Then
            'date 1 < date 2
        ElseIf (iannee1 > iannee2) Then
            bInversion = True
        ElseIf (imois1 < imois2) Then
            'date 1 < date 2
        ElseIf (imois1 > imois2) Then
            bInversion = True
        ElseIf (ijour1 < ijour2) Then
            'date 1 < date 2
        ElseIf (ijour1 > ijour2) Then
            bInversion = True
        Else
            MsgBox("Enter different dates")
        End If

        If (bInversion = True) Then
            iTemp = ijour1
            ijour1 = ijour2
            ijour2 = iTemp
            iTemp = imois1
            imois1 = imois2
            imois2 = iTemp
            iTemp = iannee1
            iannee1 = iannee2
            iannee2 = iTemp
        End If
    End Sub




    Public Function go(ByVal day1, ByVal month1, ByVal year1, ByVal day2, ByVal month2, ByVal year2)

        'Test function bissextile
        'TextBox1.Text = "2000 est bissextile ? " & IsBissextile(2000) & vbCrLf
        'TextBox1.Text = TextBox1.Text & "1996 est bissextile ? " & IsBissextile(1996) & vbCrLf
        'TextBox1.Text = TextBox1.Text & "1995 est bissextile ? " & IsBissextile(1995) & vbCrLf
        Dim result As Integer
        result = -1



        If year2 > year1 Then
            result = calculate(day1, month1, year1, day2, month2, year2)
        Else
            If year2 = year1 Then
                If month2 > month1 Then
                    result = calculateMonthsAndDays(day1, month1, day2, month2, year1)
                Else
                    If month2 = month1 Then
                        If day2 > day1 Then
                            printDays(day2 - day1)
                            result = day2 - day1
                        Else
                            printError()
                        End If
                    Else
                        printError()
                    End If
                End If
            Else
                printError()
            End If

        End If

        Return result


    End Function

    'Procedure error
    Public Sub printError()
        TextBox1.ForeColor = Color.Red
        TextBox1.Text = "The first date must be smaller than the second"
    End Sub


    'Procedure calculation of the number of days in total
    Public Function calculate(ByVal day1, ByVal month1, ByVal year1, ByVal day2, ByVal month2, ByVal year2) As Integer

        Dim iNbDaysFromNYD As Integer
        Dim iNbDaysToEOY As Integer
        Dim iNbDaysBetween2NYD As Integer

        iNbDaysToEOY = nbDaysToEOY(day1, month1, year1)
        iNbDaysBetween2NYD = nbDaysBetween2NYD(year1, year2)
        iNbDaysFromNYD = nbDaysFromNYD(day2, month2, year2)
        print(iNbDaysToEOY, iNbDaysBetween2NYD, iNbDaysFromNYD)

        Return (iNbDaysToEOY + iNbDaysBetween2NYD + iNbDaysFromNYD)

    End Function


    'Procedure print
    Public Sub print(ByVal arg1 As Integer, ByVal arg2 As Integer, ByVal arg3 As Integer)
        TextBox1.Text = "Number of days to the end of the year = " & arg1 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Number of days between 2 new year days = " & arg2 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Number of days since the new year day = " & arg3 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Difference between the two dates (in days) = " & (arg1 + arg2 + arg3)
    End Sub

    'calculate the number of days between 2 years given in parameter
    Public Function nbDaysBetween2NYD(ByVal year1 As Integer, ByVal year2 As Integer) As Integer
        Dim numberOfDayResult As Integer

        numberOfDayResult = 0
        For index = year1 + 1 To year2 - 1
            If (iSBissextile(index)) Then
                numberOfDayResult = numberOfDayResult + 366
            Else
                numberOfDayResult = numberOfDayResult + 365
            End If
        Next

        Return numberOfDayResult

    End Function

    'Calculate if the year is bissextile
    Public Function iSBissextile(ByVal year As Integer) As Boolean
        Dim b1 As Boolean
        Dim b2 As Boolean

        'Another test that works
        'If (((Year Mod 4 = 0) And (Year Mod 100 <> 0)) Or (Year Mod 400 = 0)) Then
        'End If

        If year Mod 4 = 0 Then
            If year Mod 100 = 0 Then
                b1 = False
            Else
                b1 = True
            End If
        End If

        If year Mod 400 = 0 Then
            b2 = True
        Else
            b2 = False

        End If

        Return b1 Or b2

    End Function

    'Calculate the number of days since the new year day
    Public Function nbDaysFromNYD(ByVal day As Integer, ByVal month As Integer, ByVal year As Integer) As Integer
        Dim result As Integer

        result = 0
        For index = 0 To month - 2
            result = result + daysInAMonth(index)
        Next

        If iSBissextile(year) Then
            result = result + 1
        End If

        result = result + day - 1 'A VERIFIER POUR LE -1

        Return result

    End Function

    'Calculate the number of days remaining until the end of the year
    Public Function nbDaysToEOY(ByVal day As Integer, ByVal month As Integer, ByVal year As Integer) As Integer
        Dim result As Integer

        result = 0
        For index = month + 1 To 12
            result = result + daysInAMonth(index - 1)
        Next

        If iSBissextile(year) Then
            If month <= 2 Then
                'If day <= 1 Then
                result = result + 1
                'End If
            End If
        End If

        result = result + daysInAMonth(month - 1) - day + 1 'A VERIFIER POUR LE +1


        Return result

    End Function

    'ONLY FOR THE SPECIAL CASE YEAR1 = YEAR2 AND MONTH1 = MONTH2
    'Procedure printDays
    Public Sub printDays(ByVal arg1 As Integer)
        TextBox1.Text = "Difference between the two dates (in days) = " & arg1 & vbCrLf
    End Sub


    'ONLY FOR THE SPECIAL CASE YEAR1 = YEAR2
    'Calculate the number of days remaining until the end of the month
    Public Function nbDaysToEOM(ByVal day As Integer, ByVal month As Integer, ByVal year As Integer) As Integer
        Dim result As Integer

        result = daysInAMonth(month - 1) - day

        If month = 2 Then
            If iSBissextile(year) Then
                result = result + 1
            End If
        End If

        Return result
    End Function

    'ONLY FOR THE SPECIAL CASE YEAR1 = YEAR2
    'calculate the number of days between two new months
    Public Function nbDaysBetween2NM(ByVal month1 As Integer, ByVal month2 As Integer, ByVal year As Integer) As Integer
        Dim numberOfDayResult As Integer

        numberOfDayResult = 0
        For index = month1 + 1 To month2 - 1
            numberOfDayResult = numberOfDayResult + daysInAMonth(index)
            If index = 2 Then
                If iSBissextile(year) Then
                    numberOfDayResult = numberOfDayResult + 1
                End If
            End If
        Next

        Return numberOfDayResult

    End Function

    'ONLY FOR THE SPECIAL CASE YEAR1 = YEAR2
    'Procedure printMonthsAndDays
    Public Sub printMonthsAndDays(ByVal arg1 As Integer, ByVal arg2 As Integer, ByVal arg3 As Integer)
        TextBox1.Text = "Number of days to the end of the month = " & arg1 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Number of days between 2 new month days = " & arg2 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Number of days since the new month day = " & arg3 & vbCrLf
        TextBox1.Text = TextBox1.Text & "Difference between the two dates (in days) = " & (arg1 + arg2 + arg3)
    End Sub

    'ONLY FOR THE SPECIAL CASE YEAR1 = YEAR2
    'Function calculation of the number of days between two dates that have the same year but not the same month
    Public Function calculateMonthsAndDays(ByVal day1 As Integer, ByVal month1 As Integer, ByVal day2 As Integer, ByVal month2 As Integer, ByVal year As Integer)
        Dim iNbDaysToEOM As Integer
        Dim iNbDaysBetween2NM As Integer

        iNbDaysToEOM = nbDaysToEOM(day1, month1, year)
        iNbDaysBetween2NM = nbDaysBetween2NM(month1, month1, year)
        printMonthsAndDays(iNbDaysToEOM, iNbDaysBetween2NM, day2)

        Return iNbDaysToEOM + iNbDaysBetween2NM + day2

    End Function

   
   
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Initialisation
        TextBox2.Text = "01/01/2000"
        TextBox3.Text = "01/01/2010"
    End Sub
End Class
