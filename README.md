# email2excel

This VBA macro allows freelancers and small buisiness owners to sort emails sent from a server on the cloud to generate an Excel worksheet with indivisual items in the body of incoming emails with a specific Outlook acount.

The following briefly describes how to change and set the email2excelVBA.
1. Replace the sample email address by your Outlook acount of interest.
    # If acc = "shishid@saaipf.com" Then
2. Change the subject of email to your own one.
    # If ol_obj_item.Subject = "Thank you for your purchase of bFaaaP Switch" Then
3. Select and replace each marker word of the email body (sepalated by CRLF) (e.g., "Name:").
    # If InStr(arr(j), "Name:") <> 0 Then
    #     Cells(k, 4) = GiveContent(arr(j), "Name:")
    # End If
4. Run the macro (cell(B,2) should be 0 to incorporate incoming emails in the ascending order).

For detail, please visit the email2excel website including a YouTube video and specific instructions.
https://ui.saaipf.com/email2excelabout/

If you find that this email2excelVBA macro helps your business, please consider donation on the website.

Thank you
Tomo Shishido
shishid@saaipf.com
