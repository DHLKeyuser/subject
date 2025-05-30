VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    On Error Resume Next

    If Item.Class = olMail Then

        Dim htmlBody As String

        Dim plainText As String

        Dim subjectText As String

        Dim i As Integer

        Dim pdfFound As Boolean

        pdfFound = False

        htmlBody = Item.htmlBody

        subjectText = LCase(Item.Subject)

        plainText = StripHTML(htmlBody)

        ' Normalize everything

        plainText = Replace(plainText, "&nbsp;", " ")

        plainText = Replace(plainText, Chr(160), " ")

        plainText = Replace(plainText, vbCrLf, " ")

        plainText = Replace(plainText, vbLf, " ")

        plainText = Replace(plainText, vbCr, " ")

        plainText = Replace(plainText, vbTab, " ")

        Do While InStr(plainText, "  ") > 0

            plainText = Replace(plainText, "  ", " ")

        Loop

        plainText = LCase(Trim(plainText))

        ' === LIST OF MATCH CONDITIONS ===

        Dim trigger As Boolean

        trigger = False

        ' 1. Dear Customs, please provide Export documents.

        If InStr(plainText, "dear customs, please provide export documents") > 0 Then trigger = True

        ' 2. Dear DSV (body) + subject contains "shipping 5l"

        If InStr(plainText, "dear dsv") > 0 And InStr(subjectText, "shipping 5l") > 0 Then trigger = True

        ' 3. Body contains: Would you be so kind to fill in below information ?

        If InStr(plainText, "would you be so kind to fill in below information ?") > 0 Then trigger = True

        ' 4. Body contains: Dear Van den Boogaard

        If InStr(plainText, "dear van den boogaard") > 0 Then trigger = True

        ' 5. Body contains: Dear Tromp, Shipment is ready for pick up.

        If InStr(plainText, "dear tromp, shipment is ready for pick up.") > 0 Then trigger = True

        ' 6. Dear Global DG Solution, please check below and advise.

        If InStr(plainText, "dear global dg solution, please check below and advise.") > 0 Then trigger = True

        ' === END CONDITIONS ===

        If trigger Then

            For i = 1 To Item.Attachments.Count

                If LCase(Right(Item.Attachments(i).FileName, 4)) = ".pdf" Then

                    pdfFound = True

                    Exit For

                End If

            Next i

            If Not pdfFound Then

                If MsgBox("No invoice attached. Send anyway?", vbYesNo + vbExclamation, "Missing PDF") = vbNo Then

                    Cancel = True

                End If

            End If

        End If

    End If

End Sub

Function StripHTML(strHTML As String) As String

    Dim objRegex As Object

    Set objRegex = CreateObject("VBScript.RegExp")

    With objRegex

        .pattern = "<[^>]+>" ' matches HTML tags

        .Global = True

        StripHTML = .Replace(strHTML, "")

    End With

End Function
 
