'Sebastian Soto
'RCET0265
'Fall 2021
'Car Rental
'https://github.com/SebastianSotoMk4/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim customersHelped As Double
    Dim maxMiles As Double
    Dim totoalDayCharge As Double

    'This sub will set Form to default values when called
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
        SummaryButton.Enabled = False 'SummaryButton disabled on clear even when there is summary data avalable. also what about the sumary menu item? - TJR
        MilesradioButton.Checked = True
        KilometersradioButton.Checked = False
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    'The UserInputCheck() Function checks for valid user input, this function returns a error code that will indicate what 
    'Textbox is invalid and must be fixed.
    Function UserInputCheck() As Integer
        Dim testNumber As Integer
        Dim checkNumber As Integer
        Dim errorCheck As Integer
        Try
            If NameTextBox.Text = "" Then
                errorCheck = 1
                NameTextBox.Text = ""
                TabIndex = 0
            Else
                testNumber = CInt(NameTextBox.Text) 'Why test if this is a number? - TJR
                errorCheck = 1
                NameTextBox.Text = ""
                TabIndex = 0
            End If

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
            If CityTextBox.Text = "" Then
                errorCheck = 3
                CityTextBox.Text = ""
                TabIndex = 2
            Else
                testNumber = CInt(CityTextBox.Text)'Why test if this is a number? - TJR
                errorCheck = 3
                CityTextBox.Text = ""
                TabIndex = 2
            End If
        Catch ex As Exception
            If CityTextBox.Text <> "" Then
                checkNumber += 1
            End If
        End Try

        Try
            If StateTextBox.Text = "" Then
                errorCheck = 4
                StateTextBox.Text = ""
                TabIndex = 3
            Else
                testNumber = CInt(StateTextBox.Text)'Why test if this is a number? - TJR
                errorCheck = 4
                StateTextBox.Text = ""
                TabIndex = 3
            End If
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
        ElseIf checkNumber = 0 Then
            errorCheck = 11

        End If
        Return errorCheck

    End Function

    'the AddDays() function calculates the charge for amount of days and returns the amount discounted to be used for 
    'the MinusDiscount TextBox.
    Function AddDays() As Double
        Dim days As Integer
        Dim dayCharge As Double = 15
        Dim discountReturnDays As Double
        days = CInt(DaysTextBox.Text)
        dayCharge = days * dayCharge
        discountReturnDays = dayCharge
        Select Case Discounts()
            Case = 0

            Case = 1
                discountReturnDays *= 0.05
            Case = 2
                discountReturnDays *= 0.03
            Case = 3
                discountReturnDays *= 0.08
        End Select
        DayChargeTextBox.Text = dayCharge.ToString("c")
        Return discountReturnDays
    End Function

    'The AddMiles() Function calculates Cost for miles and also handels converting from KM to MI
    'This function returns the amount discounted.
    Function AddMiles() As Double
        Dim totalMiles As Double
        Dim mileConvert As Double
        Dim milesCharge As Double
        Dim discountReturnMiles As Double
        totalMiles = CDbl(EndOdometerTextBox.Text) - CDbl(BeginOdometerTextBox.Text)
        If KilometersradioButton.Checked = True Then
            For u = 0 To totalMiles
                mileConvert += 0.621371
            Next
            mileConvert = Math.Ceiling(mileConvert)
            totalMiles = CDbl(mileConvert)
        End If
        Me.maxMiles += totalMiles
        TotalMilesTextBox.Text = ($"{totalMiles}  mi")
        totalMiles -= 200
        If totalMiles >= 0 Then
            For i = 1 To totalMiles
        If i > 500 Then 'miles form 500 to 700 calculating at 0.12. Should offset to 300 here - TJR
                    milesCharge += 0.1
        ElseIf i <= 500 Then 'miles form 500 to 700 calculating at 0.12. Should offset to 300 here - TJR
                    milesCharge += 0.12
                End If
            Next
        End If
        discountReturnMiles = milesCharge
        Select Case Discounts()
            Case = 0

            Case = 1
                discountReturnMiles *= 0.05
            Case = 2
                discountReturnMiles *= 0.03
            Case = 3
                discountReturnMiles *= 0.08
        End Select
        MileageChargeTextBox.Text = milesCharge.ToString("c")
        Return discountReturnMiles
    End Function

    'the Discounts() function checks for discounts to be acounted for and returns a number 
    'based on what discounts were aplied  
    Function Discounts() As Double
        Dim totalDiscount As Double
        If Seniorcheckbox.Checked = False And AAAcheckbox.Checked = False Then
            totalDiscount = 0
        End If
        If AAAcheckbox.Checked = True Then
            totalDiscount = 1
        End If
        If Seniorcheckbox.Checked = True Then
            totalDiscount = 2
        End If
        If Seniorcheckbox.Checked = True And AAAcheckbox.Checked = True Then
            totalDiscount = 3
        End If
        Return totalDiscount
    End Function

    'This function calculates the Total amout customer owes and returns the value.
    Function AmountOwed() As Double
        Dim totalOwed As Double
        totalOwed = CDbl(MileageChargeTextBox.Text) + CDbl(DayChargeTextBox.Text)
        If AAAcheckbox.Checked = True Or Seniorcheckbox.Checked = True Then
            totalOwed = CDbl(MileageChargeTextBox.Text) + CDbl(DayChargeTextBox.Text) - CDbl(TotalDiscountTextBox.Text)
        End If
        Me.totoalDayCharge += totalOwed

        Return totalOwed
    End Function

    'Sets defaults when programs loads
    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaults()
    End Sub

    'sets defaults when clear is clicked
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click
        SetDefaults()
    End Sub

    'activates the calculations also handels error codes
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click
        Dim discountStorage As Double
        Select Case UserInputCheck()
            Case = 0
                If AAAcheckbox.Checked = True Or Seniorcheckbox.Checked = True Then
                    discountStorage = AddDays() + AddMiles()
                Else
                    AddDays()
                    AddMiles()
                End If
                TotalDiscountTextBox.Text = discountStorage.ToString("c")
                TotalChargeTextBox.Text = AmountOwed().ToString("c")

            Case = 1
                MsgBox("Invalid Name entered.")
            Case = 2
                MsgBox("Invalid Address entered.")
            Case = 3
                MsgBox("Invalid city entered.")
            Case = 4
                MsgBox("Invalid State entered.")
            Case = 5
                MsgBox("Invalid zip code entered.")
            Case = 6
                MsgBox("Invalid Beginning Odometer entered.")
            Case = 7
                MsgBox("Invalid Enging Odomter entered.")
            Case = 8
                MsgBox("Invalid Number oF days entered.")
            Case = 9
                MsgBox("Invalid Miles we inputed wrong entered.") 'Nearly english words - TJR
            Case = 10
                MsgBox("Invalid amout of days entered.")
            Case = 11
                MsgBox("All Feilds empty.")
        End Select
        SummaryButton.Enabled = True 'what about the sumary menu item? - TJR
        Me.customersHelped += 1
    End Sub

    'Asks user if they want to close when the exit button is called
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click
        'Closes program when clicked
        Dim answer As Integer

        answer = MsgBox("Are you sure about Exiting?", vbYesNo)
        If answer = 6 Then
            Me.Close()
        ElseIf answer = 7 Then

        End If
    End Sub

    'Opens a Pop up with a sumary of the Miles, money made, and total customers
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click
        MsgBox($" Amount made to day is ${totoalDayCharge}.{vbNewLine}{vbNewLine} Miles driven by customers {maxMiles} Mi.{vbNewLine}{vbNewLine} Amount of customers helped {customersHelped}.")
    End Sub
End Class
