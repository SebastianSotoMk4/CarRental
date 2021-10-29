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
        Dim beginingMiles As Integer
        Dim endMiles As Integer
        Dim totalMiles As Integer
        Dim amoutOwed As Double
        beginingMiles = CInt(BeginOdometerTextBox.Text)
        endMiles = CInt(EndOdometerTextBox.Text)

        Select Case endMiles - beginingMiles
            Case < 200
            Case > 201
                totalMiles = 200 - (endMiles - beginingMiles)
                amoutOwed = 0.12 * totalMiles
        End Select


    End Sub
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaults()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        SetDefaults()
    End Sub
End Class