Option Strict On
Option Infer Off
Option Explicit On
'Program: Subs and Functions Practice
'Purpose: Chapter 6 of Visual Basic Book
'Name:Mason Merritt
'Date: Feb 18th,2022
Public Class Form1
    Private Sub btnListBoxSub_Click(sender As Object, e As EventArgs) Handles btnListBoxSub.Click

        'Call to Sub Procedure to Fill List Box with numbers 1 - 20

        FillListBox()
    End Sub

    Private Sub FillListBox()

        'This Independent Sub Procedure will build the ListBox for the Program to call

        lstSubFill.Items.Clear()

        Dim count As Integer = 1

        While count <= 20
            lstSubFill.Items.Add(count)
            count += 1
        End While
    End Sub

    Private Sub btnSalaryDisplay_Click(sender As Object, e As EventArgs) Handles btnSalaryDisplay.Click

        'Call to Sub Procedure to Fill List Box with Salary Increments step 5,000 from the User input to a max salary of 200,000

        Dim usersalary As Integer
        Integer.TryParse(txtUserSalaryInput.Text, usersalary)

        DisplaySalary(usersalary)
    End Sub

    Private Sub DisplaySalary(ByVal userInput As Integer)

        'This Independent Sub Procedure will build the ListBox with Salary Increments step 5,000 from the User input to a max salary of 200,000

        lstSalaryDisplay.Items.Clear()

        While userInput <= 200000
            lstSalaryDisplay.Items.Add(userInput)
            userInput += 5000
        End While
    End Sub

    Private Sub btnIncomeCalc_Click(sender As Object, e As EventArgs) Handles btnIncomeCalc.Click

        'Call to Sub Procedure to Display a User's net Income after deducting Expenses from Income

        Dim userIncome As Double
        Dim userExpenses As Double
        Dim userNetTotal As Double

        Double.TryParse(txtUserIncome.Text, userIncome)
        Double.TryParse(txtUserExpenses.Text, userExpenses)

        NetIncomeCalculator(userIncome, userExpenses, userNetTotal)
    End Sub

    Private Sub NetIncomeCalculator(ByVal incomeTotal As Double, ByVal expenseTotal As Double, ByRef monthlyNet As Double)

        'Sub Procedure to calculate Net Income

        txtUserNetIncome.Text = ""

        monthlyNet = incomeTotal - expenseTotal

        txtUserNetIncome.Text = "Your monthly net income is " & monthlyNet.ToString("C2")

    End Sub

    Private Sub btnHistoryCalc_Click(sender As Object, e As EventArgs) Handles btnHistoryCalc.Click

        'Build a Program using Independent Sub Procedures that will take in a poiint total and Selecta class. Based upon score display their Letter grade in the label
        'Only Call to the Sub Procedure from here.

        Select Case True
            Case radHistory101.Checked
                Display101()
            Case radHistory201.Checked
                Display201()
        End Select
    End Sub

    Private Sub Display101()

        'Sub Procedure to grade for History 101

        Dim historyScore As Integer

        Integer.TryParse(txtHistoryPoints.Text, historyScore)

        Select Case historyScore
            Case Is >= 90
                lblHistoryGrade.Text = "A"
            Case Is >= 80
                lblHistoryGrade.Text = "B"
            Case Is >= 70
                lblHistoryGrade.Text = "C"
            Case Is > 60
                lblHistoryGrade.Text = "D"
            Case Else
                lblHistoryGrade.Text = "F"
        End Select
    End Sub

    Private Sub Display201()

        'Sub Procedure to grade for History 201

        Dim historyScore As Integer

        Integer.TryParse(txtHistoryPoints.Text, historyScore)

        Select Case historyScore
            Case Is >= 75
                lblHistoryGrade.Text = "P"
            Case Else
                lblHistoryGrade.Text = "F"
        End Select
    End Sub

    Private Sub btnPayCalc_Click(sender As Object, e As EventArgs) Handles btnPayCalc.Click

        'Build an App that Uses two Independent Sub Procedures. One will calculate a user pay per week based on Salary and the other will calculate
        'Bi-Weekly pay per week based on Salary.

        Dim userPay As Double

        Double.TryParse(txtInputSalary.Text, userPay)

        Select Case True
            Case radWeekly.Checked
                Weekly(userPay)
                txtPayResults.Text = userPay.ToString("C2")
            Case radBiWeekly.Checked
                biWeekly(userPay)
        End Select
    End Sub

    Private Sub Weekly(ByRef weeklyPay As Double)

        'Excercise 2 asks to allow the button click event to display a result in the txtLabel. To do this I have changed the Weekly Sub Procedure from By Value to 
        'By reference. This will take the userPay variable (MEMORY LOCATION) and pass it to the Sub Procedure. When the Calculation is finished the procedure returns the
        'Value of the Calculation back to the MEMORY LOCATION of userPay and is userPay variable is assigned the value of the calculation from the sub procedure.

        weeklyPay /= 52
    End Sub

    Private Sub biWeekly(ByVal biweeklyPay As Double)

        biweeklyPay /= 26
        txtPayResults.Text = biweeklyPay.ToString("C2")
    End Sub

    Private Sub btnYouCanDoItOne_Click(sender As Object, e As EventArgs) Handles btnYouCanDoItOne.Click

        'Build an App that will take in a Number and then multiply by 2 in a Independent Sub Procedure with the Click event Calling that Procedure

        Dim userNumber As Double

        Double.TryParse(txtYouCanDoItOne.Text, userNumber)

        MultiplyResults(userNumber)
    End Sub

    Private Sub MultiplyResults(ByVal numberMultiply As Double)

        'Build an App that will take in a Number and then multiply by 2 in a Independent Sub Procedure with the Cl;ick event Calling that Procedure

        Dim total As Double

        total = numberMultiply * 2

        txtYouCanDoItResults.Text = "The number multiplied by 2 is " & total.ToString
    End Sub

    Private Sub btnTicketCalc_Click(sender As Object, e As EventArgs) Handles btnTicketCalc.Click

        'Build an App that will take in an amount of tickets and then display the total amount due. A standard ticket costs $62.50 and a VIP ticket costs
        '$102.75. If less than 5 tickets are purchaed there is no Discount. 5 to 9 tickets will receive a 5% discount and anything over 9 will get a 10%
        'discount. Use Independent subprocedures to make the calculations and the click even call the sub procedure

        Dim userTickets As Double
        Dim subtotal As Double
        Dim total As Double
        Dim discount As Double

        Double.TryParse(txtUserTickets.Text, userTickets)

        Select Case True
            Case radStandard.Checked
                subtotal = userTickets * 62.5
            Case radVIP.Checked
                subtotal = userTickets * 102.75
        End Select

        CalcDiscount(userTickets, subtotal, discount)

        Math.Round(subtotal, 2)
        Math.Round(discount, 2)

        total = subtotal - discount

        txtTicketSubtotal.Text = subtotal.ToString("C2")
        txtTicketTotal.Text = total.ToString("C2")
    End Sub

    Private Sub CalcDiscount(ByVal totalTickets As Double, ByVal beforeDiscount As Double, ByRef ticketDiscount As Double)

        'Build an App that will take in an amount of tickets and then display the total amount due. A standard ticket costs $62.50 and a VIP ticket costs
        '$102.75. If less than 5 tickets are purchaed there is no Discount. 5 to 9 tickets will receive a 5% discount and anything over 9 will get a 10%
        'discount. Use Independent subprocedures to make the calculations and the click even call the sub procedure

        Select Case totalTickets
            Case >= 10
                ticketDiscount = beforeDiscount * 0.1
                txtTicketDiscount.Text = ticketDiscount.ToString("C2")
            Case >= 5
                ticketDiscount = beforeDiscount * 0.05
                txtTicketDiscount.Text = ticketDiscount.ToString("C2")
            Case Else
                ticketDiscount = 0
                txtTicketDiscount.Text = ticketDiscount.ToString("C2")
        End Select
    End Sub

    Private Sub btnYouCanDoItTwo_Click(sender As Object, e As EventArgs) Handles btnYouCanDoItTwo.Click

        'Build an App that will take in a Number then pass the variable LOCATION to an indpendent sub procedure and multiply by two and output
        'the results into the label

        Dim userNumber As Double

        Double.TryParse(txtYouCanDoItTwo.Text, userNumber)

        CalcDouble(userNumber)
    End Sub

    Private Sub CalcDouble(ByRef multiplyNumber As Double)

        'Build an App that will take in a Number then pass the variable LOCATION to an indpendent sub procedure and multiply by two and output
        'the results into the label


        Dim total As Double

        total = multiplyNumber * 2

        txtYouCanDoItTwoResults.Text = "The number multiplied by 2 is " & total.ToString
    End Sub

    Private Sub btnNewPayFunction_Click(sender As Object, e As EventArgs) Handles btnNewPayFunction.Click

        'Build an App that takes in a users old salary and increases it by 2% using a Function Procedure display results in the text label

        Dim userSalary As Double

        Double.TryParse(txtNewPayFunction.Text, userSalary)

        txtNewPayResults.Text = "Your new pay is " & GetNewPay(userSalary).ToString("C2")
    End Sub

    Private Function GetNewPay(ByVal oldSalary As Double) As Double

        'Build an App that takes in a users old salary and increases it by 2% using a Function Procedure display results in the text label

        ' Dim newSalary As Double

        'newSalary = oldSalary * 1.02
        'Return newSalary

        'Can also write the function like this

        Return oldSalary * 1.02
    End Function

    Private Sub btnFuncTicketCalc_Click(sender As Object, e As EventArgs) Handles btnFuncTicketCalc.Click

        'Build an App that will take in an amount of tickets and then display the total amount due. A standard ticket costs $62.50 and a VIP ticket costs
        '$102.75. If less than 5 tickets are purchaed there is no Discount. 5 to 9 tickets will receive a 5% discount and anything over 9 will get a 10%
        'discount. Use Function Procedures to make the calculations and the click invoke the function value

        Dim userTickets As Double
        Dim subTotal As Double
        Dim discount As Double
        Dim totalDue As Double

        Double.TryParse(txtFunctionTickets.Text, userTickets)

        Select Case True
            Case radFunctionStandard.Checked
                subTotal = userTickets * 62.5
            Case radFunctionVIP.Checked
                subTotal = userTickets * 102.75
        End Select

        discount = GetDiscount(userTickets, subTotal, discount)

        totalDue = subTotal - discount

        txtFuncSubTotal.Text = subTotal.ToString("C2")
        txtFuncDiscount.Text = GetDiscount(userTickets, subTotal, discount).ToString("C2")
        txtFuncTotalDue.Text = totalDue.ToString("C2")
    End Sub

    Private Function GetDiscount(ByVal totalTickets As Double, ByVal totalBeforeDiscount As Double, ByRef discountApplied As Double) As Double

        'Build an App that will take in an amount of tickets and then display the total amount due. A standard ticket costs $62.50 and a VIP ticket costs
        '$102.75. If less than 5 tickets are purchaed there is no Discount. 5 to 9 tickets will receive a 5% discount and anything over 9 will get a 10%
        'discount. Use Function Procedures to make the calculations and the click invoke the function value

        Select Case totalTickets
            Case >= 10
                discountApplied = totalBeforeDiscount * 0.1
            Case >= 5
                discountApplied = totalBeforeDiscount * 0.05
            Case Else
                discountApplied = 0
        End Select
        Return discountApplied
    End Function

    Private Sub btnYouCanDoItThree_Click(sender As Object, e As EventArgs) Handles btnYouCanDoItThree.Click

        'Build an App that will take in a User Number and multiply by 2 within a Function Procedure and display the results in the text label

        Dim userNumber As Double

        Double.TryParse(txtYouCanDoItThree.Text, userNumber)

        MultiplyNumber(userNumber)

        txtYouCanDoItThreeResults.Text = "The number multipled by 2 is " & MultiplyNumber(userNumber).ToString
    End Sub

    Private Function MultiplyNumber(ByVal number As Double) As Double

        Return number * 2
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Add Info to Combo Boxes

        'Simple Combo Box

        cboSimple.Items.Add("Mason")
        cboSimple.Items.Add("Veronica")
        cboSimple.Items.Add("Gracie")
        cboSimple.Items.Add("Owen")
        cboSimple.Items.Add("Barry")
        cboSimple.Items.Add("Sal")
        'Choose Selected Item for Simple Combo Box

        cboSimple.SelectedIndex = 0

        'Drop Down Combo Box

        cboDropDown.Items.Add("Tokyo")
        cboDropDown.Items.Add("Paris")
        cboDropDown.Items.Add("New York")
        cboDropDown.Items.Add("Hong Kong")
        'Choose Selected Item for Drop Down Combo Box
        cboDropDown.SelectedItem = "Paris"

        'Drop Down List Combo Box

        cboDropDownList.Items.Add("North Carolina")
        cboDropDownList.Items.Add("Mississippi")
        cboDropDownList.Items.Add("California")
        cboDropDownList.Items.Add("Oregon")
        cboDropDownList.Items.Add("Louisiana")
        'Choose Selected Item for Drop Down List Combo Box
        cboDropDownList.Text = "Mississippi"

        'Three different ways of choosing a Selected Item in the Combo Box

        'Adding Items to Mortage Rates Combo Box

        Dim count As Double = 2
        While count <= 7
            cboMortgageRates.Items.Add(count.ToString)
            count += 0.5
        End While
        cboMortgageRates.SelectedIndex = 0

        'Add Info to ListBoxes and ComboBoxes fgor Cerruti Payroll App

        'add Hours for hours ListBox

        Dim hours As Double = 1
        While hours <= 168
            lstHoursWorked.Items.Add(hours)
            lstHours2.Items.Add(hours)
            hours += 0.5
        End While
        lstHoursWorked.SelectedIndex = 39
        lstHours2.SelectedIndex = 39

        'Add PayRates to Payrate ListBox

        Dim payrate As Double = 7.25
        While payrate <= 100
            lstPayRate.Items.Add(payrate)
            lstPayRates2.Items.Add(payrate)
            payrate += 0.25
        End While
        lstPayRate.SelectedIndex = 31
        lstPayRates2.SelectedIndex = 31

        'Add Allowances to Allowance Combo Box

        Dim allow As Double = 1
        While allow <= 10
            cboAllowances.Items.Add(allow)
            cboAllow2.Items.Add(allow)
            allow += 1
            cboAllowances.SelectedIndex = 0
            cboAllow2.SelectedIndex = 0
        End While

        'Excercise 4 requires New Payroll app. Building ListBoxes for Hours and PayRate as Well as Combo Box for Allowances. Used Same Variables as before just added the Listbox
        'And ComboBox names to the loops previously built

        'Donut Shoppe Radio Buttons Selected

        radGlazed.Checked = True
        radNone.Checked = True

        'Add Translations to Combo Box

        cboTranslation.Items.Add("French")
        cboTranslation.Items.Add("Spanish")
        cboTranslation.Items.Add("Italian")
        cboTranslation.SelectedIndex = 0

        'Add Translations to Excercise 10 

        cboLanguage.Items.Add("French")
        cboLanguage.Items.Add("Spanish")
        cboLanguage.Items.Add("Italian")
        cboLanguage.SelectedIndex = 0

        'Excercise 11 Cable Direct Listboxes

        Dim pchannels As Integer
        While pchannels <= 10
            lstPremiumChannels.Items.Add(pchannels)
            pchannels += 1
        End While

        Dim connections As Integer
        While connections <= 100
            lstCableConnections.Items.Add(connections)
            connections += 1
        End While

        'Excercise 12 Wallpaper Warehouse Combo Boxes (Length, width, height display from 8 - 30) (coverage displays 30 - 40 step 0.5 increments)

        Dim length As Integer = 8
        While length <= 30
            cboLength.Items.Add(length)
            length += 1
        End While
        cboLength.Text = "10"

        Dim width As Integer = 8
        While width <= 30
            cboWidth.Items.Add(width)
            width += 1
        End While
        cboWidth.Text = "10"

        Dim height As Integer = 8
        While height <= 30
            cboHeight.Items.Add(height)
            height += 1
        End While
        cboHeight.Text = "8"

        Dim coverage As Double = 30
        While coverage <= 40
            cboCoverage.Items.Add(coverage)
            coverage += 0.5
        End While
        cboCoverage.Text = "37"
    End Sub

    Private Sub btnMortgageCalc_Click(sender As Object, e As EventArgs) Handles btnMortgageCalc.Click

        'Build an App that will calculate the monthly payment on a mortgage based on 15 20 25 and 30 year periods. Allow the rates to be from 2%
        'to 8% using a COMBO BOX and display the results in the text label. Use a independent sub or function to calculate

        Dim principal As Double
        Dim rate As Double

        Double.TryParse(txtMortgagePrincipal.Text, principal)
        Double.TryParse(cboMortgageRates.Text, rate)

        MortgagePayment(principal, rate)
    End Sub

    Private Sub MortgagePayment(ByVal cost As Double, ByVal rate As Double)

        'Build an App that will calculate the monthly payment on a mortgage based on 15 20 25 and 30 year periods. Allow the rates to be from 2%
        'to 8% and display the results in the text label. Use a independent sub or function to calculate

        Dim Years As Integer = 15
        Dim monthlyPayment As Double
        rate /= 100

        MessageBox.Show(rate.ToString)

        While Years <= 30
            monthlyPayment = -Financial.Pmt(rate / 12, 12 * Years, cost)
            txtMortageResults.Text = txtMortageResults.Text & Years.ToString & " years:" & ControlChars.Tab & monthlyPayment.ToString("C2") & vbCrLf
            Years += 5
        End While
    End Sub

    Private Sub cboAllowances_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cboAllowances.KeyPress

        'Only Allow Users to enter the Numbers 1 - 9 and the Backspace key

        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub cboAllowances_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAllowances.SelectedIndexChanged

        'Sub Procedure to handle all the event changes to clear labels when user clicks to change something

        txtGrossPay.Text = ""
        txtFWH.Text = ""
        txtFICA.Text = ""
        txtNetPay.Text = ""
    End Sub

    Private Sub btnPayrollCalc_Click(sender As Object, e As EventArgs) Handles btnPayrollCalc.Click

        'Build an App that will take in a User's Name, Martial Status, Hours Worked, Payrate and Total Allowances allowed. Calculate their Federal
        'Tax Withholdings as well as their FICA tax withholding(7.65%) and display the total results of Gross Pay, FWT, FICA and Net Pay.
        'use an Independent Sub Procedure to calculate for SINGLE and a Function Procedure to Calculate for Married.


        'declare constants

        Const oneAllowance As Double = 77.9
        Const ficaRate As Double = 0.0765

        'declare variables

        Dim name As String = txtPayrollName.Text
        Dim hours As Double
        Dim payrate As Double
        Dim grosspay As Double
        Dim taxablepay As Double
        Dim ficaTax As Double
        Dim fwt As Double
        Dim netPay As Double
        Dim userAllowance As Double

        'TryParse

        Double.TryParse(lstHoursWorked.SelectedItem.ToString, hours)
        Double.TryParse(lstPayRate.SelectedItem.ToString, payrate)
        Double.TryParse(cboAllowances.Text, userAllowance)

        'calculate gross pay

        If hours <= 40 Then
            grosspay = payrate * hours
        Else
            grosspay = (payrate * 40) + ((hours - 40) * payrate * 1.5)
        End If


        'calculate FICA tax

        ficaTax = grosspay * ficaRate

        'calculate taxable wages

        taxablepay = grosspay - (userAllowance * oneAllowance)

        'which Tax bracket is being chosen to calculate FWT

        Select Case True
            Case radSingle.Checked
                Singletaxes(taxablepay, fwt)
            Case Else
                fwt = Marriedtaxes(taxablepay)
        End Select

        'round all numbers

        grosspay = Math.Round(grosspay, 2)
        ficaTax = Math.Round(ficaTax, 2)
        fwt = Math.Round(fwt, 2)

        netPay = grosspay - ficaTax - fwt
        Math.Round(netPay, 2)

        'Display Results of Procedures

        txtGrossPay.Text = grosspay.ToString("C2")
        txtFWH.Text = fwt.ToString("C2")
        txtFICA.Text = ficaTax.ToString("C2")
        txtNetPay.Text = netPay.ToString("C2")

    End Sub

    Private Sub Singletaxes(ByVal payTaxable As Double, ByRef federalTaxes As Double)

        'This sub procedure is to Calculate the Federal Withholding Taxes for a Single Person

        Select Case payTaxable
            Case <= 44
                federalTaxes = 0
            Case <= 224
                federalTaxes = (payTaxable - 44) * 0.1
            Case <= 774
                federalTaxes = (payTaxable - 224) * 0.15 + 18
            Case <= 1812
                federalTaxes = (payTaxable - 774) * 0.25 + 100.5
            Case <= 3730
                federalTaxes = (payTaxable - 1812) * 0.28 + 360
            Case <= 8058
                federalTaxes = (payTaxable - 3730) * 0.33 + 897.04
            Case <= 8090
                federalTaxes = (payTaxable - 8058) * 0.35 + 2325.28
            Case Else
                federalTaxes = (payTaxable - 8090) * 0.396 + 2336.48
        End Select

    End Sub

    Private Function Marriedtaxes(ByVal payTaxable As Double) As Double

        Dim federalTaxes As Double

        Select Case payTaxable
            Case <= 166
                federalTaxes = 0
            Case <= 525
                federalTaxes = (payTaxable - 166) * 0.1
            Case <= 1626
                federalTaxes = (payTaxable - 525) * 0.15 + 35.9
            Case <= 3111
                federalTaxes = (payTaxable - 1626) * 0.25 + 201.05
            Case <= 4654
                federalTaxes = (payTaxable - 3111) * 0.28 + 572.3
            Case <= 8180
                federalTaxes = (payTaxable - 4654) * 0.33 + 1004.34
            Case <= 9218
                federalTaxes = (payTaxable - 8180) * 0.35 + 2167.92
            Case Else
                federalTaxes = (payTaxable - 9218) * 0.396 + 2531.22
        End Select
        Return federalTaxes
    End Function

    Private Sub btnPayrollExit_Click(sender As Object, e As EventArgs) Handles btnPayrollExit.Click

        Dim answer As Integer

        MessageBox.Show("Do you want to Quit?", "Close program", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If answer = vbYes Then
            System.Windows.Forms.Application.ExitThread()
            Me.Close()
        End If
    End Sub

    Private Sub btnHistoryFunc_Click(sender As Object, e As EventArgs) Handles btnHistoryFunc.Click

        'Excercise 1 asks to Build the History App using a Function Procedure. History 101 has a A-F grading scale
        'traditional style. History 201 has a Pass or Fail system with 75 being the lowest passing grade.

        Dim totalPoints As Integer

        Integer.TryParse(txtPointsFunc.Text, totalPoints)

        If rad101Function.Checked Then
            txtHistoryResultsFunction.Text = History101Function(totalPoints)
        Else
            txtHistoryResultsFunction.Text = History201Function(totalPoints)
        End If

    End Sub

    Private Function History101Function(ByVal scoreTotal As Integer) As String

        'Function for Scoring History 101

        Dim score As String

        Select Case scoreTotal
            Case >= 90
                score = "A"
            Case >= 80
                score = "B"
            Case >= 70
                score = "C"
            Case >= 60
                score = "D"
            Case Else
                score = "F"
        End Select
        Return score
    End Function

    Private Function History201Function(ByVal scoreTotal As Integer) As String

        'Function for Scoring History 201

        Dim score As String

        Select Case scoreTotal
            Case >= 75
                score = "P"
            Case Else
                score = "F"
        End Select
        Return score
    End Function

    Private Sub btnGrossFunction_Click(sender As Object, e As EventArgs) Handles btnGrossFunction.Click

        'Excercise 3 asks to build the Gross pay app using a Function Procedure. Take in the UserSalary and then either give a weekly or biweekly total of paychecks

        Dim userSalary As Double

        Double.TryParse(txtGrossFunctionSalary.Text, userSalary)

        Select Case True
            Case radFunctionWeekly.Checked
                txtGrossFunctionResults.Text = WeeklyCheck(userSalary)
            Case Else
                txtGrossFunctionResults.Text = BiWeeklyCheck(userSalary)
        End Select
    End Sub

    Private Function WeeklyCheck(ByVal salary As Double) As String

        'Dim results As String
        Dim check As Double

        check = salary / 52
        check = Math.Round(check, 2)
        Return check.ToString
    End Function

    Private Function BiWeeklyCheck(ByVal salary As Double) As String

        'Dim results As String
        Dim check As Double

        check = salary / 26
        check = Math.Round(check, 2)
        Return check.ToString
    End Function

    Private Sub radFunctionWeekly_CheckedChanged(sender As Object, e As EventArgs) Handles radFunctionWeekly.CheckedChanged

        'Clear Results Label

        txtGrossFunctionResults.Text = ""
    End Sub

    Private Sub btnCalcPayroll2_Click(sender As Object, e As EventArgs) Handles btnCalcPayroll2.Click

        'Excercise 4 is to rebuild the payroll App that will take in a User's Name, Martial Status, Hours Worked, Payrate and Total Allowances allowed. Calculate their Federal
        'Tax Withholdings as well as their FICA tax withholding(7.65%) and display the total results of Gross Pay, FWT, FICA and Net Pay.
        'use an Function Procedure to calculate for SINGLE and a Independent Sub Procedure to Calculate for Married.

        'declare constants

        Const oneAllowance As Double = 77.9
        Const ficaRate As Double = 0.0765

        'delcare variables

        Dim userName As String = txtPayroll2.Text
        Dim hours As Double
        Dim payRate As Double
        Dim userAllowance As Double
        Dim grossPay As Double
        Dim ficaTax As Double
        Dim taxableIncome As Double
        Dim fwt As Double
        Dim netPay As Double

        'Tryparses

        Double.TryParse(lstHours2.SelectedItem.ToString, hours)
        Double.TryParse(lstPayRates2.SelectedItem.ToString, payRate)
        Double.TryParse(cboAllow2.Text, userAllowance)

        'gross pay total

        If hours <= 40 Then
            grossPay = payRate * hours
        Else
            grossPay = (payRate * hours) + ((hours - 40) * payRate * 1.5)
        End If

        'fica tax

        ficaTax = grossPay * ficaRate

        'taxable income

        taxableIncome = grossPay - (userAllowance * oneAllowance)
        taxableIncome = Math.Round(taxableIncome, 2)

        'check martial status and call to proedures. Single will use a function and Married will use a Independent Sub

        Select Case True
            Case radSingle2.Checked
                fwt = Singletax(taxableIncome)
            Case Else
                Marriedtax(taxableIncome, fwt)
        End Select

        'round numbers

        grossPay = Math.Round(grossPay, 2)
        fwt = Math.Round(fwt, 2)
        ficaTax = Math.Round(ficaTax, 2)

        'get net pay total

        netPay = grossPay - ficaTax - fwt
        netPay = Math.Round(netPay, 2)

        'Display Final Results

        txtGrosspay2.Text = grossPay.ToString("C2")
        txtFWT2.Text = fwt.ToString("C2")
        txtfica2.Text = ficaTax.ToString("C2")
        txtNetPay2.Text = netPay.ToString("C2")
    End Sub

    Private Function Singletax(ByVal income2Tax As Double) As Double

        'Excercise 4 is to rebuild the payroll App - This is the Function Procedure for the Single Taxes

        'Declare variable for tax

        Dim federaltaxes As Double

        Select Case income2Tax
            Case <= 44
                federaltaxes = 0
            Case <= 224
                federaltaxes = (income2Tax - 44) * 0.1
            Case <= 774
                federaltaxes = (income2Tax - 224) * 0.15 + 18
            Case <= 1812
                federaltaxes = (income2Tax - 774) * 0.25 + 100.5
            Case <= 3730
                federaltaxes = (income2Tax - 1812) * 0.28 + 360
            Case <= 8058
                federaltaxes = (income2Tax - 3730) * 0.33 + 897.04
            Case <= 8090
                federaltaxes = (income2Tax - 8058) * 0.35 + 2325.28
            Case Else
                federaltaxes = (income2Tax - 8090) * 0.396 + 2336.48
        End Select
        Return federaltaxes
    End Function

    Private Sub Marriedtax(ByVal income2Tax As Double, ByRef federalTaxes As Double)

        'Excercise 4 is to rebuild the payroll App. This is the Independent sub Procedure for Married taxes

        Select Case income2Tax
            Case <= 166
                federalTaxes = 0
            Case <= 525
                federalTaxes = (income2Tax - 166) * 0.1
            Case <= 1626
                federalTaxes = (income2Tax - 525) * 0.15 + 35.9
            Case <= 3111
                federalTaxes = (income2Tax - 1626) * 0.25 + 201.05
            Case <= 4654
                federalTaxes = (income2Tax - 3111) * 0.28 + 572.3
            Case <= 8180
                federalTaxes = (income2Tax - 4654) * 0.33 + 1004.34
            Case <= 9218
                federalTaxes = (income2Tax - 8180) * 0.35 + 2167.92
            Case Else
                federalTaxes = (income2Tax - 9218) * 0.396 + 2531.22
        End Select
    End Sub

    Private Sub cboAllow2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAllow2.SelectedIndexChanged

        'Clear labels for payroll program 2

        txtGrosspay2.Text = ""
        txtFWT2.Text = ""
        txtfica2.Text = ""
        txtNetPay2.Text = ""
    End Sub

    Private Sub btnSeminarsCalc_Click(sender As Object, e As EventArgs) Handles btnSeminarsCalc.Click

        'Excercise 5--Build an App that totals the Amount due based upon the user selection in Checkboxes. Use a Function Procedure

        txtSeminarsResults.Text = Seminarfunction()

    End Sub

    Private Function Seminarfunction() As String

        Dim total As Double

        If chkFinance.Checked Then
            total += 150
        End If

        If chkHealth.Checked Then
            total += 125
        End If

        If chkMarketing.Checked Then
            total += 135
        End If

        Return total.ToString
    End Function

    Private Sub ClearOutput(sender As Object, e As EventArgs)

        'Clear Label for Seminars Excercise 5 and 6

        txtSeminarsResults.Text = ""
        txtSeminarsResults2.Text = ""

    End Sub

    Private Sub btnSeminarsCalc2_Click(sender As Object, e As EventArgs) Handles btnSeminarsCalc2.Click

        'Excercise 6--Build an App that totals the Amount due based upon the user selection in Radio Button. Use a Independent Sub Procedure

        SeminarSub()

    End Sub

    Private Sub SeminarSub()

        'Excercise 6--Build an App that totals the Amount due based upon the user selection in Radio Button. Use a Independent Sub Procedure

        Dim results As String

        Select Case True
            Case radFinance.Checked
                results = "$150.00"
                txtSeminarsResults2.Text = results
            Case radHealth.Checked
                results = "$125.00"
                txtSeminarsResults2.Text = results
            Case radMarketing.Checked
                results = "$135.00"
                txtSeminarsResults2.Text = results
        End Select
    End Sub

    Private Sub btnDonutCalc_Click(sender As Object, e As EventArgs) Handles btnDonutCalc.Click

        'Excercise 7- Build an App that takes an User's order and add a 6% sales tax. Use a Independent Sub Procedure to find the SUBTOTAL and 
        'use a Function Procedure to fin the TAX due
        'The button procedure should call these procedures take their values and find the total due

        'Variables

        Dim subTotal As Double
        Dim tax As Double
        Dim total As Double

        'Call Procedures

        SubTotalPrice(subTotal)
        tax = TaxTotal(subTotal, tax)

        'Calculate total

        total = subTotal + tax
        total = Math.Round(total, 2)

        'Display Results

        txtDonutSubTotal.Text = subTotal.ToString("C2")
        txtDonutTax.Text = tax.ToString("C2")
        txtDonutTotal.Text = total.ToString("C2")
    End Sub

    Private Sub SubTotalPrice(ByRef beforeTax As Double)

        'Excercise 7- Independent Sub Procedure to Calculate SubTotal

        Select Case True
            Case radGlazed.Checked
                beforeTax = 1.05
            Case radSugar.Checked
                beforeTax = 1.05
            Case radChoco.Checked
                beforeTax = 1.25
            Case radFilled.Checked
                beforeTax = 1.5
        End Select

        If radNone.Checked Then
            beforeTax += 0
        ElseIf radRegular.Checked Then
            beforeTax += 1.5
        ElseIf radCappu.Checked Then
            beforeTax += 2.75
        End If

        beforeTax = Math.Round(beforeTax, 2)

    End Sub

    Private Function TaxTotal(ByVal beforeTax As Double, totalTax As Double) As Double

        'Excercise 7- Function Procedure to Calculate Total tax Due

        totalTax = beforeTax * 0.06
        totalTax = Math.Round(totalTax, 2)
        Return totalTax
    End Function

    Private Sub btnMatCalc_Click(sender As Object, e As EventArgs) Handles btnMatCalc.Click

        'Build a Program that will take a Users order and any extra charges and display in the results Label. Use a Function to calculate the rug price before any extra charges
        'Use a Independent Sub procedure to calculate any extra charges then in the button procedure finalize with taxes of 9.275% and display final results

        'Variables

        Dim subTotal As Double
        Dim taxRate As Double = 0.09275
        Dim total As Double
        Dim taxes As Double

        'Call to Function

        subTotal = PriceBeforeExtras()

        'Call to Independent Sub

        PriceWithExtras(subTotal)

        'Caluclate taxes

        taxes = subTotal * taxRate

        'Calculate Total Price

        total = subTotal + taxes

        'Display Results

        txtMatResults.Text = total.ToString("C2")
    End Sub

    Private Function PriceBeforeExtras() As Double

        'Function to Calculate Price before any Extra Charges

        Select Case True
            Case radStandRug.Checked
                PriceBeforeExtras = 99
            Case radDeluxeRug.Checked
                PriceBeforeExtras = 129
            Case radPremiumRug.Checked
                PriceBeforeExtras = 179
        End Select
        Return PriceBeforeExtras
    End Function

    Private Sub PriceWithExtras(ByRef pricewithExtras As Double)

        'Independent Sub to Add any extras to price of the Rug

        If radBlue.Checked Then
            pricewithExtras = pricewithExtras
        ElseIf radRed.Checked Then
            pricewithExtras += 10
        ElseIf radPink.Checked Then
            pricewithExtras += 15
        End If

        If chkFoldable.Checked Then
            pricewithExtras += 25
        End If
    End Sub

    Private Sub btnTranslator_Click(sender As Object, e As EventArgs) Handles btnTranslator.Click

        'Build an App that will allow a User to select an English word then choose between French, Spanish, and Italian from the combo box to translate to.
        ' Use 3 Function Procedures to find the results and Display results in the Label

        'Call to Functions

        If cboTranslation.Text = "French" Then
            txtTranslatorResults.Text = FrenchTranslation()
        ElseIf cboTranslation.Text = "Spanish" Then
            txtTranslatorResults.Text = SpanishTranslation()
        ElseIf cboTranslation.Text = "Italian" Then
            txtTranslatorResults.Text = ItalianTranslation()
        End If
    End Sub

    Private Function FrenchTranslation() As String

        'Program to Translate English words to French

        'Variables

        Dim results As String = ""

        Select Case True
            Case radFather.Checked AndAlso cboTranslation.Text = "French"
                results = "P`ere"
            Case radMother.Checked AndAlso cboTranslation.Text = "French"
                results = "M`ere"
            Case radBrother.Checked AndAlso cboTranslation.Text = "French"
                results = "Fr`ere"
            Case radSister.Checked AndAlso cboTranslation.Text = "French"
                results = "Soeur"
        End Select
        Return results
    End Function

    Private Function SpanishTranslation() As String

        'Program to Translate English words to French

        'Variables

        Dim results As String = ""

        Select Case True
            Case radFather.Checked AndAlso cboTranslation.Text = "Spanish"
                results = "Padre"
            Case radMother.Checked AndAlso cboTranslation.Text = "Spanish"
                results = "Madre"
            Case radBrother.Checked AndAlso cboTranslation.Text = "Spanish"
                results = "Hermano"
            Case radSister.Checked AndAlso cboTranslation.Text = "Spanish"
                results = "Hermana"
        End Select
        Return results
    End Function

    Private Function ItalianTranslation() As String

        'Program to Translate English words to French

        'Variables

        Dim results As String = ""

        Select Case True
            Case radFather.Checked AndAlso cboTranslation.Text = "Italian"
                results = "Padre"
            Case radMother.Checked AndAlso cboTranslation.Text = "Italian"
                results = "Madre"
            Case radBrother.Checked AndAlso cboTranslation.Text = "Italian"
                results = "Fratello"
            Case radSister.Checked AndAlso cboTranslation.Text = "Italian"
                results = "Sorella"
        End Select
        Return results
    End Function

    Private Sub btnTranslatorCalc_Click(sender As Object, e As EventArgs) Handles btnTranslatorCalc.Click

        'Build an App that will allow a User to select an English word then choose between French, Spanish, and Italian from the combo box to translate to.
        ' Use 3 Independent Sub Procedures to find the results and Display results in the Label

        'Call to Independent Subs

        If cboTranslation.Text = "French" Then
            FrenchTranslationSub()
        ElseIf cboTranslation.Text = "Spanish" Then
            SpanishTranslationSub()
        ElseIf cboTranslation.Text = "Italian" Then
            ItalianTranslationSub()
        End If
    End Sub

    Private Sub FrenchTranslationSub()

        'Independent Sub to Translate from English to French

        'Variables

        Dim results As String

        'Case Statement

        Select Case True
            Case radPapa.Checked AndAlso cboLanguage.Text = "French"
                results = "P`ere"
                txtTranslationResults.Text = results
            Case radMama.Checked AndAlso cboLanguage.Text = "French"
                results = "M`ere"
                txtTranslationResults.Text = results
            Case radBro.Checked AndAlso cboLanguage.Text = "French"
                results = "Fr`ere"
                txtTranslationResults.Text = results
            Case radSis.Checked AndAlso cboLanguage.Text = "French"
                results = "Soeur"
                txtTranslationResults.Text = results
        End Select
    End Sub

    Private Sub SpanishTranslationSub()

        'Independent Sub to Translate from English to Spanish

        'Variables

        Dim results As String

        'Case Statement

        Select Case True
            Case radPapa.Checked AndAlso cboLanguage.Text = "Spanish"
                results = "Padre"
                txtTranslationResults.Text = results
            Case radMama.Checked AndAlso cboLanguage.Text = "Spanish"
                results = "Madre"
                txtTranslationResults.Text = results
            Case radBro.Checked AndAlso cboLanguage.Text = "Spanish"
                results = "Hermano"
                txtTranslationResults.Text = results
            Case radSis.Checked AndAlso cboLanguage.Text = "Spanish"
                results = "Hermana"
                txtTranslationResults.Text = results
        End Select
    End Sub

    Private Sub ItalianTranslationSub()

        'Independent Sub to Translate from English to Italian

        'Variables

        Dim results As String

        'Case Statement

        Select Case True
            Case radPapa.Checked AndAlso cboLanguage.Text = "Italian"
                results = "Padre"
                txtTranslationResults.Text = results
            Case radMama.Checked AndAlso cboLanguage.Text = "Italian"
                results = "Madre"
                txtTranslationResults.Text = results
            Case radBro.Checked AndAlso cboLanguage.Text = "Italian"
                results = "Fratello"
                txtTranslationResults.Text = results
            Case radSis.Checked AndAlso cboLanguage.Text = "Italian"
                results = "Sorella"
                txtTranslationResults.Text = results
        End Select
    End Sub

    Private Sub btnCableCalc_Click(sender As Object, e As EventArgs) Handles btnCableCalc.Click

        'Excercise 11-- Build a Program that displays a Users cable bill based upon the seklected features they choose. use 2 Function Procedures
        'One for RESIDENTIAL customers with pricing as follows: $4.50 processing fee, $30 basic cable fee, $5 for each premium channel
        'The BUSINESS customer pricing is as follows: $16.50 processing fee, $80 for the first 4 connections then $4 for each additional connection
        'premium channels cost a flat rate $50 for any number of connections. Use one function for residential and one for business.

        'Call to Functions

        If radResidential.Checked Then
            txtCableResults.Text = ResidentialBilling()
        ElseIf radBusiness.Checked Then
            txtCableResults.Text = BusinessBilling()
        End If
    End Sub

    Private Function ResidentialBilling() As String

        'Function Procedure to Calculate Residential Cable Billing

        'Variable to store Results

        Dim total As Double
        Dim premChannels As Integer

        Integer.TryParse(lstPremiumChannels.SelectedItem.ToString, premChannels)

        total = 4.5 + 30 + (5 * premChannels)

        Return "Your bill is " & total.ToString("C2")
    End Function

    Private Function BusinessBilling() As String

        'Function Procedure to Calculate Business Cable Billing

        'Variable to store Results

        Dim total As Double
        Dim premChannels As Integer
        Dim connections As Integer

        Integer.TryParse(lstPremiumChannels.SelectedItem.ToString, premChannels)
        Integer.TryParse(lstCableConnections.SelectedItem.ToString, connections)

        If connections <= 10 Then
            total = 16.5 + 80 + (50 * premChannels)
        ElseIf connections > 10 Then
            total = 96.5 + ((connections - 10) * 4) + (50 * premChannels)
        End If

        Return "Your bill is " & total.ToString("C2")
    End Function

    Private Sub btnWallpaperCalc_Click(sender As Object, e As EventArgs) Handles btnWallpaperCalc.Click

        'Excercise 12--Wallpaper Warehouse-- Build A program that will take in a users Length, Width and Height then allow them to select a paper size coverage
        'Calculate how many rolls will be needed to cover the room based upon the rolls coverage size. Use an Independent Sub to make the calculation.

        'Call to Sub

        WallpaperCalculation()

    End Sub

    Private Sub WallpaperCalculation()

        'Independent Sub for Excercise 12 Wallpaper Warehouse Calculation

        'Variables

        Dim length As Double
        Dim width As Double
        Dim Height As Double
        Dim coverage As Double
        Dim totalSquareFt As Double
        Dim rollsNeeded As Double

        'TryParse

        Double.TryParse(cboLength.SelectedItem.ToString, length)
        Double.TryParse(cboWidth.SelectedItem.ToString, width)
        Double.TryParse(cboHeight.SelectedItem.ToString, Height)
        Double.TryParse(cboCoverage.SelectedItem.ToString, coverage)

        'Calculation

        totalSquareFt = ((length * 2) * Height) + ((width * 2) * Height)
        rollsNeeded = totalSquareFt / coverage
        rollsNeeded = Math.Ceiling(rollsNeeded)

        txtWallpaperResults.Text = "You'll need " & rollsNeeded.ToString & " rolls of wallpaper to cover the room"
    End Sub
End Class
