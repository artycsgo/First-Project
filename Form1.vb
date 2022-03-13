Public Class Form1
    Dim subTotalPrice, count As Double
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim itemName As String = ""
        Dim price, itemAmount As Double
        If Not txtItemName.Text = "" And Not txtItemPrice.Text = "" And Not txtItemAmount.Text = "" Then
            If IsNumeric(txtItemPrice.Text) = True And IsNumeric(txtItemAmount.Text) Then
                If txtItemPrice.Text > 0 And txtItemAmount.Text > 0 Then
                    addItemInput(price, itemName, itemAmount)
                    addItemCalc(price, itemAmount)
                    subTotalPrice += price
                    count += 1
                    addItemOutput(count, itemName, itemAmount, price)
                    addItemClear()
                ElseIf Not txtItemPrice.Text > 0 And txtItemAmount.Text > 0 Then
                    MsgBox("Please input a number greater than 1 for the price")
                    txtItemPrice.Clear()
                ElseIf txtItemPrice.Text > 0 And Not txtItemAmount.Text > 0 Then
                    MsgBox("Please input a number greater than 1 for the amount")
                    txtItemAmount.Clear()
                ElseIf Not txtItemPrice.Text > 0 And Not txtItemAmount.Text > 0 Then
                    MsgBox("Please input a number greater than 1 for the price and the amount")
                    txtItemPrice.Clear()
                    txtItemAmount.Clear()
                End If
            ElseIf IsNumeric(txtItemPrice.Text) = False And IsNumeric(txtItemAmount.Text) = True Then
                MsgBox("Please enter a number for the price")
                txtItemPrice.Clear()
            ElseIf IsNumeric(txtItemPrice.Text) = True And IsNumeric(txtItemAmount.Text) = False Then
                MsgBox("Please enter a number for the amount")
                txtItemAmount.Clear()
            ElseIf IsNumeric(txtItemPrice.Text) = False And IsNumeric(txtItemAmount.Text) = False Then
                MsgBox("Please enter a number for the price and the amount")
                txtItemPrice.Clear()
                txtItemAmount.Clear()
            End If
        ElseIf txtItemName.Text = "" And Not txtItemPrice.Text = "" And Not txtItemAmount.Text = "" Then
            MsgBox("Please enter a name")
        ElseIf Not txtItemName.Text = "" And txtItemPrice.Text = "" And Not txtItemAmount.Text = "" Then
            MsgBox("Please enter a price")
        ElseIf Not txtItemName.Text = "" And Not txtItemPrice.Text = "" And txtItemAmount.Text = "" Then
            MsgBox("Please enter an amount")
        ElseIf txtItemName.Text = "" And txtItemPrice.Text = "" And Not txtItemAmount.Text = "" Then
            MsgBox("Please enter a name and a price")
        ElseIf txtItemName.Text = "" And Not txtItemPrice.Text = "" And txtItemAmount.Text = "" Then
            MsgBox("Please enter a name and an amount")
        ElseIf Not txtItemName.Text = "" And txtItemPrice.Text = "" And txtItemAmount.Text = "" Then
            MsgBox("Please enter a price and an amount")
        ElseIf txtItemName.Text = "" And txtItemPrice.Text = "" And txtItemAmount.Text = "" Then
            MsgBox("Please enter a name and a price and an amount")
        End If
    End Sub
    Sub addItemInput(ByRef price As Double, ByRef itemName As String, ByRef itemAmount As String)
        price = txtItemPrice.Text
        itemName = txtItemName.Text
        itemAmount = txtItemAmount.Text
    End Sub
    Sub addItemOutput(ByVal count As Double, ByVal itemName As String, ByVal itemAmount As Double, ByVal price As Double)
        lstBox.Items.Add("Item " & count & ":" & vbTab & "Item Name: " & itemName & vbTab & "Item Amount: " & itemAmount & vbTab & "Item Price: " & FormatCurrency(price))
    End Sub
    Sub addItemClear()
        txtItemName.Clear()
        txtItemPrice.Clear()
        txtItemAmount.Clear()
    End Sub
    Private Sub btnCheckout_Click(sender As Object, e As EventArgs) Handles btnCheckout.Click
        Dim taxrate, totalPrice As Double
        Dim prompt As String = ""
        checkoutInput(prompt, taxrate)
        checkoutCalc(subTotalPrice, taxrate, totalPrice)
        checkoutOutput(subTotalPrice, taxrate, totalPrice)
    End Sub
    Sub checkoutInput(ByRef prompt As String, ByRef taxrate As Double)
        prompt = "Please enter your taxrate"
        taxrate = InputBox(prompt)
    End Sub
    Sub checkoutOutput(ByVal subTotalPrice As Double, ByVal taxrate As Double, ByVal totalPrice As Double)
        lstBox.Items.Add("Subtotal: " & FormatCurrency(subTotalPrice))
        lstBox.Items.Add("Taxrate: " & taxrate & "%")
        lstBox.Items.Add("Total: " & FormatCurrency(totalPrice))
    End Sub
    Function checkoutCalc(ByVal subTotalPrice As Double, ByVal taxrate As Double, ByRef totalPrice As Double) As Double
        totalPrice = subTotalPrice * (1 + taxrate / 100)
        Return totalPrice
    End Function
    Function addItemCalc(ByRef price As Double, ByVal itemAmount As Double) As Double
        price = price * itemAmount
        Return price
    End Function
End Class
