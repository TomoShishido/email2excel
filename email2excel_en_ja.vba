Sub email2excel_en()


    Dim ol_obj_df, f, i, j, k, n As Long
    Dim ol_obj, Accounts, acc As Object
    Dim ol_obj_ns, ol_obj_item As Object

    Dim bodywords As String
    Dim arr() As String

    Dim lastRow As Long
    Dim lastTime As Double

    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    lastTime = Cells(Rows.Count, 2).End(xlUp).Value
    'MsgBox lastTime
    k = lastRow


    ' To create an Outlook object (ol_obj)
    Set ol_obj = CreateObject("Outlook.Application")

    'To select email account
    Set Accounts = ol_obj.Session.Accounts
    For Each acc In Accounts
    'put an email address of interest below
     If acc = "shishid@saaipf.com" Then
        Set f = acc.DeliveryStore.GetDefaultFolder(6)
        For i = 1 To f.Items.Count
            Set ol_obj_item = f.Items(i)
              'update incoming emails of interest
              If ol_obj_item.ReceivedTime > lastTime Then
                'put the "subject" of email to be incorporated
                If ol_obj_item.Subject = "Thank you for your purchase of bFaaaP Switch" Then
                    k = k + 1

                    Cells(k, 1) = k
                    Cells(k, 2) = ol_obj_item.ReceivedTime
                    Cells(k, 3) = ol_obj_item.Subject

                    Cells(k, 21) = ol_obj_item.Body

                    bodywords = Cells(k, 21).Value
                    'split the body of emial by CRLF
                    arr = Split(bodywords, vbCrLf)
                    'To process each line
                    For j = LBound(arr) To UBound(arr)
                        If InStr(arr(j), "Name:") <> 0 Then
                            Cells(k, 4) = GiveContent(arr(j), "Name:")
                        End If
                        If InStr(arr(j), "Email:") <> 0 Then
                            Cells(k, 5) = GiveContent(arr(j), "Email:")
                        End If
                        If InStr(arr(j), "Age:") <> 0 Then
                            Cells(k, 6) = GiveContent(arr(j), "Age:")
                        End If
                        If InStr(arr(j), "Sex:") <> 0 Then
                            Cells(k, 7) = GiveContent(arr(j), "Sex:")
                        End If
                        If InStr(arr(j), "Address: ") <> 0 Then
                            Cells(k, 8) = GiveContent(arr(j), "Address:")
                        End If
                        If InStr(arr(j), "Zipcode:") <> 0 Then
                            Cells(k, 9) = GiveContent(arr(j), "Zipcode:")
                        End If
                        If InStr(arr(j), "Country:") <> 0 Then
                            Cells(k, 10) = GiveContent(arr(j), "Country:")
                        End If

                        If InStr(arr(j), "Message:") <> 0 Then
                            Cells(k, 11) = GiveContent(arr(j), "Message:")
                        End If
                    Next j
                End If
            End If
        Next
     End If
    Next
    'To prohibit folding of each cell
    Cells.WrapText = False
End Sub

Sub email2excel_ja()


    Dim ol_obj_df, f, i, j, k, n As Long
    Dim ol_obj, Accounts, acc As Object
    Dim ol_obj_ns, ol_obj_item As Object

    Dim bodywords As String
    Dim arr() As String

    Dim lastRow As Long
    Dim lastTime As Double

    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    lastTime = Cells(Rows.Count, 2).End(xlUp).Value
    'MsgBox lastTime
    k = lastRow


    ' To create an Outlook object (ol_obj)
    Set ol_obj = CreateObject("Outlook.Application")

    'To select email account
    Set Accounts = ol_obj.Session.Accounts
    For Each acc In Accounts
    'put an email address of interest below
     If acc = "shishid@saaipf.com" Then
        Set f = acc.DeliveryStore.GetDefaultFolder(6)
        For i = 1 To f.Items.Count
            Set ol_obj_item = f.Items(i)
              'update incoming emails of interest
              If ol_obj_item.ReceivedTime > lastTime Then
                'put the "subject" of email to be incorporated
                If ol_obj_item.Subject = "bFaaaP Switch???????????????????????????????????????????????????" Then
                    k = k + 1

                    Cells(k, 1) = k
                    Cells(k, 2) = ol_obj_item.ReceivedTime
                    Cells(k, 3) = ol_obj_item.Subject

                    Cells(k, 21) = ol_obj_item.Body

                    bodywords = Cells(k, 21).Value
                    'split the body of emial by CRLF
                    arr = Split(bodywords, vbCrLf)
                    'To process each line
                    For j = LBound(arr) To UBound(arr)
                        If InStr(arr(j), "??????:") <> 0 Then
                            Cells(k, 4) = GiveContent(arr(j), "??????:")
                        End If
                        If InStr(arr(j), "Email:") <> 0 Then
                            Cells(k, 5) = GiveContent(arr(j), "Email:")
                        End If
                        If InStr(arr(j), "??????:") <> 0 Then
                            Cells(k, 6) = GiveContent(arr(j), "??????:")
                        End If
                        If InStr(arr(j), "??????:") <> 0 Then
                            Cells(k, 7) = GiveContent(arr(j), "??????:")
                        End If
                        If InStr(arr(j), "????????????:") <> 0 Then
                            Cells(k, 8) = GiveContent(arr(j), "????????????:")
                        End If
                        If InStr(arr(j), "?????????????????????:") <> 0 Then
                            Cells(k, 9) = GiveContent(arr(j), "?????????????????????:")
                        End If
                        If InStr(arr(j), "?????????????????????:") <> 0 Then
                            Cells(k, 10) = GiveContent(arr(j), "?????????????????????:")
                        End If

                        If InStr(arr(j), "???????????????:") <> 0 Then
                            Cells(k, 11) = GiveContent(arr(j), "???????????????:")
                        End If
                    Next j
                End If
            End If
        Next
     End If
    Next
    'To prohibit folding of each cell
    Cells.WrapText = False
End Sub


Function GiveContent(stringline As String, markerword As String) As String
    Dim processedContent As String
    Dim markerwordcount As Long
    markerwordcount = Len(markerword)

    processedContent = Mid(stringline, InStr(stringline, markerword) + markerwordcount)
    'If the stringline ends with ",", remove it
    If Right(processedContent, 1) = "," Then
        processedContent = Left(processedContent, Len(processedContent) - 1)
    End If


    GiveContent = processedContent

End Function
