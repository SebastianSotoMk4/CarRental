'Sebastian Soto
'RCET0265
'Fall 2021
'Car Rental
'https://github.com/SebastianSotoMk4/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        'Closes program when clicked
        Dim answer As Integer

        answer = MsgBox("close?", vbYesNo)
        If answer = 6 Then
            Me.Close()
        ElseIf answer = 7 Then

        End If
    End Sub
    'Sets Defaults when called
    Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub
    Function UserInputCheck() As Integer
        Dim testNumber As Integer
        Dim checkNumber As Integer
        Dim errorCheck As Integer

        Try
            testNumber = CInt(NameTextBox.Text)
            errorCheck = 1
            NameTextBox.Text = ""
        Catch ex As Exception
            If NameTextBox.Text <> "" Then
                checkNumber += 1
            End If
        End Try


        If AddressTextBox.Text <> "" Then
            checkNumber += 1
        ElseIf AddressTextBox.Text = "" Then
            errorCheck = 2
            AddressTextBox.Text = ""
        End If

        Try
            testNumber = CInt(CityTextBox.Text)
            errorCheck = 3
            CityTextBox.Text = ""
        Catch ex As Exception
            If CityTextBox.Text <> "" Then
                checkNumber += 1
            End If
        End Try

        Try
            testNumber = CInt(StateTextBox.Text)
            errorCheck = 4
            StateTextBox.Text = ""
        Catch ex As Exception
            If StateTextBox.Text <> "" Then
                checkNumber += 1
            End If

        End Try

        Try
            testNumber = CInt(ZipCodeTextBox.Text)

            If ZipCodeTextBox.Text <> "" Then
                checkNumber += 1
            End If
        Catch ex As Exception
            errorCheck = 5
            ZipCodeTextBox.Text = ""
        End Try

        Try
            testNumber = CInt(BeginOdometerTextBox.Text)

            If BeginOdometerTextBox.Text <> "" Then
                checkNumber += 1
            End If
        Catch ex As Exception
            errorCheck = 6
            BeginOdometerTextBox.Text = ""
        End Try

        Try
            testNumber = CInt(EndOdometerTextBox.Text)
            If CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text) > 0 Then
                checkNumber += 1
            Else
                errorCheck = 9
                EndOdometerTextBox.Text = ""
            End If
            If EndOdometerTextBox.Text <> "" Then

            End If
        Catch ex As Exception
            errorCheck = 7
            EndOdometerTextBox.Text = ""
        End Try

        Try
            testNumber = CInt(DaysTextBox.Text)
            If CInt(DaysTextBox.Text) - 45 <= -45 Or CInt(DaysTextBox.Text) - 45 >= 1 Then
                errorCheck = 10
            Else

                If DaysTextBox.Text <> "" Then
                    checkNumber += 1
                End If
            End If

        Catch ex As Exception
            errorCheck = 8
        End Try

        If checkNumber = 8 Then
            errorCheck = 0

        End If
        Return errorCheck

    End Function
    Sub AddMiles()
        Dim totalMiles As Integer
        Dim mileConvert As Double
        Dim milesCharge As Double
        totalMiles = CInt(EndOdometerTextBox.Text) - CInt(BeginOdometerTextBox.Text)
        If KilometersradioButton.Checked = True Then
            For u = 0 To totalMiles
                mileConvert += 0.621371
            Next
            mileConvert = Math.Ceiling(mileConvert)
            totalMiles = CInt(mileConvert)
        End If


        totalMiles -= 200
        If totalMiles >= 0 Then
            For i = 1 To totalMiles
                If i > 500 Then
                    milesCharge += 0.1
                ElseIf i <= 500 Then
                    milesCharge += 0.12
                End If
            Next
        End If
        MileageChargeTextBox.Text = milesCharge.ToString("c")
    End Sub
    Sub AddDays()
        Dim days As Integer
        Dim dayCharge As Integer = 15
        days = CInt(DaysTextBox.Text)
        dayCharge = days * dayCharge
        DayChargeTextBox.Text = dayCharge.ToString("c")
    End Sub
    Sub AddDistance()
        Dim totalDistance As Integer
        totalDistance = CInt(BeginOdometerTextBox.Text) + CInt(EndOdometerTextBox.Text)
        TotalMilesTextBox.Text = ($"{totalDistance}  mi")
    End Sub
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaults()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub


    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click



        Select Case UserInputCheck()
            Case = 0
                AddDays()
                AddMiles()
                AddDistance()
            Case = 1
                MsgBox("Name")
            Case = 2
                MsgBox("Address")
            Case = 3
                MsgBox("city")
            Case = 4
                MsgBox("State")
            Case = 5
                MsgBox("zip code")
            Case = 6
                MsgBox("Beginning Odometer")
            Case = 7
                MsgBox("Enging Odomter")
            Case = 8
                MsgBox("Number oF days")
            Case = 9
                MsgBox("Miles we inputed wrong")
            Case = 10
                MsgBox("invalid amout of days")


        End Select


    End Sub
End Class

'kindof a bug with Beginging odometer being 0