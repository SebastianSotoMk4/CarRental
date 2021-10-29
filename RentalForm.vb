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
        MsgBox("close?", vbYesNo)
        If vbYesNo = 4 Then
            Me.Close()
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
        DayChargeTextBox.Text = CStr(dayCharge)

    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaults()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        AddDays()
        AddMiles()
    End Sub
End Class