Attribute VB_Name = "Module5"
#If VBA7 Then

    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#Else

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#End If

Public Sub DG()

   On Error GoTo 0

    ' 1) Get the active mail item

    Dim mail As MailItem

    Set mail = Application.ActiveInspector.CurrentItem

    If mail Is Nothing Then

        MsgBox "No compose window found.", vbExclamation

        Exit Sub

    End If

    ' 2) Setup paths

    Dim exePath As String, tempPdf As String, tempTxt As String

    exePath = Environ("USERPROFILE") & "\Documents\PDFTools\bin64\pdftotext.exe"

    tempPdf = Environ("TEMP") & "\attached.pdf"

    tempTxt = Environ("TEMP") & "\output.txt"

    If Dir(exePath) = "" Then

        MsgBox "pdftotext.exe not found at:" & vbCrLf & exePath, vbExclamation

        Exit Sub

    End If

    ' 3) Save the invoice-named PDF (55###### or 89######) or first PDF

    Dim att As Attachment, invoiceFound As Boolean

    invoiceFound = False

    For Each att In mail.Attachments

        If LCase(Right(att.FileName, 4)) = ".pdf" Then

            If att.FileName Like "54######.pdf" Or att.FileName Like "55######.pdf" Or att.FileName Like "89######.pdf" Then

                att.SaveAsFile tempPdf

                invoiceFound = True: Exit For

            End If

        End If

    Next

    If Not invoiceFound Then

        For Each att In mail.Attachments

            If LCase(Right(att.FileName, 4)) = ".pdf" Then

                att.SaveAsFile tempPdf

                invoiceFound = True: Exit For

            End If

        Next

    End If

    If Not invoiceFound Then

        MsgBox "No PDF attachment found.", vbExclamation

        Exit Sub

    End If

    ' 4) Convert to text (clear stale output first)

    If Dir(tempTxt) <> "" Then

        On Error Resume Next: Kill tempTxt: On Error GoTo 0

    End If

    Shell """" & exePath & """ """ & tempPdf & """ """ & tempTxt & """", vbHide

    DoEvents

    Dim i As Long

    For i = 1 To 20

        If Dir(tempTxt) <> "" Then Exit For

        Sleep 100: DoEvents

    Next

    If Dir(tempTxt) = "" Then

        MsgBox "PDF?text conversion failed.", vbExclamation

        Exit Sub

    End If

    ' 5) Read & normalize text

    Dim fso As Object, textContent As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    textContent = LCase(fso.OpenTextFile(tempTxt, 1).ReadAll)

    textContent = Replace(textContent, vbCrLf, vbLf)

    textContent = Replace(textContent, vbCr, vbLf)

    ' 6) Extract the 8-digit Doc# (starts 55 or 89), but only if it’s not part of a longer number

    Dim documentNumber As String
    
    Dim chunk       As String
    
    Dim prevChar    As String
    
    Dim nextChar    As String
    
    For i = 1 To Len(textContent) - 7

    chunk = Mid$(textContent, i, 8)

    If (Left$(chunk, 2) = "54" Or Left$(chunk, 2) = "55" Or Left$(chunk, 2) = "89") _
       And IsNumeric(chunk) And InStr(chunk, ".") = 0 Then

        ' grab the character just before and just after the 8-digit chunk

        If i > 1 Then

            prevChar = Mid$(textContent, i - 1, 1)

        Else

            prevChar = ""

        End If

        If i + 8 <= Len(textContent) Then

            nextChar = Mid$(textContent, i + 8, 1)

        Else

            nextChar = ""

        End If

        ' only accept it if neither prevChar nor nextChar is a digit

        If (prevChar = "" Or Not (prevChar Like "[0-9]")) _
           And (nextChar = "" Or Not (nextChar Like "[0-9]")) Then

            documentNumber = chunk

            Exit For

        End If

    End If

Next
 

    ' 7) Extract priority

    Dim priority As String

    If InStr(textContent, "routine") > 0 Then

        priority = "Routine"

    ElseIf InStr(textContent, "priority") > 0 Then

        priority = "Priority"

    ElseIf InStr(textContent, "emergency") > 0 Then

        priority = "Emergency"

    End If

    ' 8) Build the “Ship To” block

    Dim lines()      As String

    Dim blockStr    As String

    Dim shipIdx     As Long, fromIdx As Long

    lines = Split(textContent, vbLf)

    For i = 0 To UBound(lines)

        If InStr(lines(i), "ship to") > 0 Then shipIdx = i: Exit For

    Next

    If shipIdx > 0 Then

        For i = shipIdx + 1 To UBound(lines)

            If InStr(lines(i), "ship from") > 0 Then

                fromIdx = i: Exit For

            End If

        Next

        Dim lastLine As Long

        lastLine = IIf(fromIdx > shipIdx, fromIdx - 1, UBound(lines))

        For i = shipIdx + 1 To lastLine

            If Trim(lines(i)) <> "" Then

                blockStr = blockStr & " " & Trim(lines(i))

            End If

        Next

        blockStr = Trim(blockStr)

    End If

    If blockStr = "" Then blockStr = textContent   ' fallback to whole text

    ' 9) Strip periods so "U.S.A." ? "usa"

    Dim cleanBlock As String

    cleanBlock = Replace(blockStr, ".", "")

    ' 10) Country list WITHOUT “netherlands”

    Dim countryList As Variant

    countryList = Array( _
    "afghanistan", "albania", "algeria", "andorra", "angola", "antigua and barbuda", "argentina", "armenia", "australia", "austria", _
    "azerbaijan", "bahamas", "bahrain", "bangladesh", "barbados", "belarus", "belgium", "belize", "benin", "bhutan", _
    "bolivia", "bosnia and herzegovina", "botswana", "brazil", "brunei", "bulgaria", "burkina faso", "burundi", "cabo verde", "cambodia", _
    "cameroon", "canada", "central african republic", "chad", "chile", "china", "colombia", "comoros", "congo", "costa rica", _
    "croatia", "cuba", "cyprus", "czech republic", "democratic republic of the congo", "denmark", "djibouti", "dominica", "dominican republic", "ecuador", _
    "egypt", "el salvador", "equatorial guinea", "eritrea", "estonia", "eswatini", "ethiopia", "fiji", "finland", "france", _
    "gabon", "gambia", "georgia", "germany", "ghana", "greece", "grenada", "guatemala", "guinea", "guinea-bissau", _
    "guyana", "haiti", "honduras", "hungary", "iceland", "india", "indonesia", "iran", "iraq", "ireland", _
    "israel", "italy", "ivory coast", "jamaica", "japan", "jordan", "kazakhstan", "kenya", "kiribati", "kuwait", _
    "kyrgyzstan", "laos", "latvia", "lebanon", "lesotho", "liberia", "libya", "liechtenstein", "lithuania", "luxembourg", _
    "madagascar", "malawi", "malaysia", "maldives", "mali", "malta", "marshall islands", "mauritania", "mauritius", "mexico", _
    "micronesia", "moldova", "monaco", "mongolia", "montenegro", "morocco", "mozambique", "myanmar", "namibia", "nauru", _
    "nepal", "new zealand", "nicaragua", "niger", "nigeria", "north korea", "north macedonia", "norway", "oman", _
    "pakistan", "palau", "palestine", "panama", "papua new guinea", "paraguay", "peru", "philippines", "poland", "portugal", _
    "qatar", "romania", "russia", "rwanda", "saint lucia", "saint vincent and the grenadines", "samoa", "san marino", "sao tome and principe", _
    "saudi arabia", "senegal", "serbia", "seychelles", "sierra leone", "singapore", "slovakia", "slovenia", "solomon islands", "somalia", _
    "south africa", "south korea", "south sudan", "spain", "sri lanka", "sudan", "suriname", "sweden", "switzerland", "syria", _
    "taiwan", "tajikistan", "tanzania", "thailand", "timor-leste", "togo", "tonga", "trinidad and tobago", "tunisia", "turkey", _
    "turkmenistan", "tuvalu", "uganda", "ukraine", "united arab emirates", "united kingdom", "united states", "uruguay", "uzbekistan", "vanuatu", _
    "vatican city", "venezuela", "vietnam", "yemen", "zambia", "zimbabwe", "u.s.a.", "usa" _
)
    ' 11) Regex-search for any of those countries

    Dim re As Object, matches As Object

    Set re = CreateObject("VBScript.RegExp")

    With re

        .Global = True

        .IgnoreCase = True

        .pattern = "\b(" & Join(countryList, "|") & ")\b"

    End With

    Set matches = re.Execute(cleanBlock)

    Dim country As String

    If matches.Count > 0 Then

        country = LCase(matches(matches.Count - 1).Value)

    Else

        country = "the netherlands"

    End If

    ' 12) Validate

    If documentNumber = "" Or priority = "" Or country = "" Then

        MsgBox "Extraction failed:" & vbCrLf & _
               "Doc#: " & documentNumber & vbCrLf & _
               "Prio: " & priority & vbCrLf & _
               "Ctry: " & country, vbExclamation

        Exit Sub

    End If

' 13) Read shipment data from downloaded file

Dim tempPath As String

Dim shipmentID As String, forwarder As String

Dim shipmentInfo As String

shipmentID = "ID"           ' default fallback

forwarder = "Forwarder"     ' default fallback

tempPath = Environ("USERPROFILE") & "\Downloads\shipment_data_temp.txt"

If Dir(tempPath) <> "" Then

    Set fso = CreateObject("Scripting.FileSystemObject")

    shipmentInfo = fso.OpenTextFile(tempPath, 1).ReadAll

    ' Extract Shipment ID and Forwarder from text

    lines = Split(shipmentInfo, vbLf)

    For i = 0 To UBound(lines)

        If Trim(lines(i)) = "Shipment ID:" Then

            If i + 1 <= UBound(lines) Then shipmentID = Trim(lines(i + 1))

        ElseIf Trim(lines(i)) = "Forwarder:" Then

            If i + 1 <= UBound(lines) Then forwarder = Trim(lines(i + 1))

        End If

    Next





End If

' 14) Assemble subject line using structured info

Dim newSubj As String

newSubj = priority & " - " & documentNumber & " - " & UCase(country) & " - Shipping 5L" & " Dangerous Goods "

' Check for DN block and count how many DNs are listed

Dim isInDNBlock As Boolean

Dim dnMatches As Long

isInDNBlock = False

dnMatches = 0

If shipmentInfo <> "" Then

    lines = Split(shipmentInfo, vbLf)

    For i = 0 To UBound(lines)

        Dim ln As String

        ln = Trim(lines(i))

        If ln = "DN:" Then

            isInDNBlock = True

        ElseIf isInDNBlock Then

            If ln = "" Then

                Exit For ' stop when we hit a blank line

            ElseIf (Len(ln) = 8 And IsNumeric(ln)) Then

                dnMatches = dnMatches + 1

            Else

                Exit For ' stop if it's not a DN

            End If

        End If

    Next

End If

If dnMatches > 1 Then

    newSubj = newSubj & " - Consolidation"

End If

mail.Subject = newSubj
 



End Sub
 










