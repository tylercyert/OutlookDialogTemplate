Private WithEvents m_Inspectors As Outlook.Inspectors
Private WithEvents m_Inspector As Outlook.Inspector

Private Sub Application_Startup()
    Set m_Inspectors = Application.Inspectors
End Sub

Private Sub m_Inspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    If TypeOf Inspector.CurrentItem Is Outlook.MailItem Then
        ' Handle emails only
        Set m_Inspector = Inspector
    End If
End Sub

Private Sub m_Inspector_Activate()
    If TypeOf m_Inspector.CurrentItem Is MailItem Then
        Set mail = m_Inspector.CurrentItem

        If mail.Subject = "SURGERY PAYMENT" Then
            If mail.BodyFormat = OlBodyFormat.olFormatHTML Then
                ' Replace [PATIENT NAME HERE] with the entered value
                If InStr(mail.HTMLBody, "[PATIENT NAME HERE]") > 0 Then
                    Value = InputBox("Enter Patient First Name:")
                    If Value <> "" Then
                        mail.HTMLBody = Replace(mail.HTMLBody, "[PATIENT NAME HERE]", Value)
                    End If
                End If
                'This is an example of a dialog box selector for two different preset options. 
                ' Replace [DR. NAME] with the entered value (doctor selector)
                If InStr(mail.HTMLBody, "[DR. NAME]") > 0 Then
                    Value = InputBox("Select the Doctor: 1 - Dr. One, 2 - Dr. Two")
                    If Value <> "" Then
                        If Value = "1" Then
                            mail.HTMLBody = Replace(mail.HTMLBody, "[DR. NAME]", "Dr. One")
                        ElseIf Value = "2" Then
                            mail.HTMLBody = Replace(mail.HTMLBody, "[DR. NAME]", " Dr. Two")
                        End If
                    End If
                End If

                ' Replace [SX DATE] with the entered value and calculate payment deadline
                If InStr(mail.HTMLBody, "[SX DATE]") > 0 Then
                    Value = InputBox("Enter the SX DATE (MM/DD/YYYY):")
                    If Value <> "" Then
                        mail.HTMLBody = Replace(mail.HTMLBody, "[SX DATE]", Value)
                        Dim sxDate As Date
                        sxDate = DateValue(Value)
                        Dim paymentDeadline As Date
                        paymentDeadline = DateAdd("d", -5, sxDate) ' Calculate deadline 5 days prior to surgery
                        mail.HTMLBody = Replace(mail.HTMLBody, "[SX PAYMENT DEADLINE]", Format(paymentDeadline, "MM/DD/YYYY"))
                    End If
                End If

                ' Replace [SX AMOUNT] with the entered value (formatted as currency)
                If InStr(mail.HTMLBody, "[SX AMOUNT]") > 0 Then
                    Value = InputBox("Enter the Full SX Amount:")
                    If Value <> "" Then
                        ' Format the input as currency with two decimal places and a dollar sign
                        mail.HTMLBody = Replace(mail.HTMLBody, "[SX AMOUNT]", FormatCurrency(Value, 2, vbUseDefault, vbUseDefault, vbUseDefault))
                    End If
                End If

                ' Replace [PAYMENT DESCRIPTION] with the entered value
                If InStr(mail.HTMLBody, "[PAYMENT DESCRIPTION]") > 0 Then
                    Value = InputBox("Enter the Payment Description:")
                    If Value <> "" Then
                        mail.HTMLBody = Replace(mail.HTMLBody, "[PAYMENT DESCRIPTION]", Value)
                    End If
                End If
            End If
        End If
    End If
End Sub
