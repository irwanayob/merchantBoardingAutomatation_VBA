Attribute VB_Name = "MerchantPage"
Sub Test1()

Dim IE As Object
Dim doc As HTMLDocument
Dim firstName As String
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
Set doc = IE.document

'----------------------------------------------------------------------Go to Login Page
IE.navigate "https://map.mcpayment.net:9999"

'-------------------------------------------------------------------Wait for loading
Do While IE.Busy
Application.Wait DateAdd("s", 1, Now)
Loop

'----------------------------------------------------------------Click the Sign in button
Set searchButton = doc.getElementsByClassName("btn btn-inverse pull-right")
'Need to refine this. Poor logic but it works
For Each Button In searchButton
    If InStr(Button.innerHTML, "Sign In") Then
        Button.Click
        Exit For
    End If
Next

'-----------------------------------------------------------------------Wait for loading
While IE.readyState < 4 Or IE.Busy: WScript.Sleep 100: Wend
'(ReadyState:4=Complete,3=Interactive,2=Loaded,1=Loading,0=Uninitialized)

'------------------------------------------------------------------------Go to Merchant Page
IE.navigate "https://map.mcpayment.net:9999/GW/Merchant/"

'-------------------------------------------------------------------------Wait for loading
Do While IE.Busy
Application.Wait DateAdd("s", 1, Now)
Loop
'-------------------------------------------------------------------------Click to collapse test area
doc.getElementById("divExpTitle_divAddMerchant").Click

Do While IE.Busy
Application.Wait DateAdd("s", 1, Now)
Loop

'--------------------------------------------------------------------------Add Section

doc.getElementById("txtFullNameAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AG2").Value
doc.getElementById("txtShortNameAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AH2").Value
doc.getElementById("txtTelCountryCodeAdd").Value = "65"
doc.getElementById("txtTelNoAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AL2").Value
doc.getElementById("txtEmailAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AM2").Value
doc.getElementById("ddlTypeAdd").Value = 1

'   // Acquirer // Drop down with typing
'doc.getElementById("ddAcquirerAdd").Value = 1005

doc.getElementById("ddlStatusAdd").Value = "02"

'doc.getElementById("txtStartDateAdd").Value = Date$ //Start Date // yyyy/mm/dd
'doc.getElementById("txtTerminationDateAdd").Value = //End Date$ // yyyy/mm/dd

doc.getElementById("ddlCreatedByAdd").Value = 1

'---------------------------------------------------Information Merchant Data Section
doc.getElementById("txtBankAccountNameAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AS2").Value
doc.getElementById("txtBankAccountNoAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AV2").Value
doc.getElementById("txtBankNameAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AQ2").Value
doc.getElementById("txtBankCodeAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AT2").Value
doc.getElementById("txtBankBranchCodeAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AU2").Value
doc.getElementById("txtBusinessRegistrationCodeAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AO2").Value
doc.getElementById("txtDescriptiveBillAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AH2").Value

'---------------------------------------------------Authorised Signer Section
'doc.getElementById("txtFirstNameAuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AP2").Value

'firstName = Left(AP2, (Find(" ", A1, 1) - 1))
doc.getElementById("txtFirstNameAuthorisedSignerAdd").Value = firstName
doc.getElementById("txtLastNameAuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AP2").Value


doc.getElementById("txtDateofBirthAuthorisedSignerAdd").Value = "1975/03/01"
doc.getElementById("txtPhoneNumberAuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AL2").Value
doc.getElementById("txtEmailAuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AM2").Value
doc.getElementById("txtIdentificationNumberAuthorisedSignerAdd").Value = "S1234567Z"
doc.getElementById("txtLine1AuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AJ2").Value
'doc.getElementById("txtLine2AuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AS2").Value
doc.getElementById("txtCityAuthorisedSignerAdd").Value = "Singapore"
doc.getElementById("txtRegionCodeAuthorisedSignerAdd").Value = "01"
doc.getElementById("txtPostalCodeAuthorisedSignerAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AT2").Value
doc.getElementById("txtCountryAuthorisedSignerAdd").Value = "Singapore"
doc.getElementById("txtCountryCodeAuthorisedSignerAdd").Value = "702"
doc.getElementById("txtCountryCodeAlphaAuthorisedSignerAdd").Value = "SGP"

'------------------------------------------------------------Significant Owner 1 Section
doc.getElementById("txtFirstNameSignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AP2").Value
doc.getElementById("txtLastNameSignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AP2").Value
doc.getElementById("txtDateofBirthSignificantOwner1DtlAdd").Value = "1975/03/01"
doc.getElementById("txtPhoneNumberSignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AL2").Value
doc.getElementById("txtEmailSignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AM2").Value
doc.getElementById("txtIdentificationNumberSignificantOwner1DtlAdd").Value = "S1234567Z"
doc.getElementById("txtLine1SignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AJ2").Value
'doc.getElementById("txtLine2SignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AS2").Value
doc.getElementById("txtCitySignificantOwner1DtlAdd").Value = "Singapore"
doc.getElementById("txtRegionCodeSignificantOwner1DtlAdd").Value = "01"
doc.getElementById("txtPostalCodeSignificantOwner1DtlAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AT2").Value
doc.getElementById("txtCountrySignificantOwner1DtlAdd").Value = "Singapore"
doc.getElementById("txtCountryCodeSignificantOwner1DtlAdd").Value = "702"
doc.getElementById("txtCountryCodeAlphaSignificantOwner1DtlAdd").Value = "SGP"

'Business Address Section
doc.getElementById("txtLine1BusinessAddressAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AJ2").Value
'doc.getElementById("txtLine2BusinessAddressAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AS2").Value
doc.getElementById("txtCityBusinessAddressAdd").Value = "Singapore"
doc.getElementById("txtRegionCodeBusinessAddressAdd").Value = "01"
doc.getElementById("txtPostalCodeBusinessAddressAdd").Value = ThisWorkbook.Sheets("Sheet1").Range("AT2").Value
doc.getElementById("txtCountryBusinessAddressAdd").Value = "Singapore"
doc.getElementById("txtCountryCodeBusinessAddressAdd").Value = "702"
doc.getElementById("txtCountryCodeAlphaBusinessAddressAdd").Value = "SGP"

'Uncheck fraud check box
doc.getElementById("chkIsFraudEnabledAdd").Click



'Dropdowns
'doc.getElementById("txtFullNameAdd").Value = "Irwan"<---- option value
'Radio Buttons
'doc.getElementById("txtFullNameAdd").Click<-------
'When cell value is different from option value
'If ThisWorbook.Sheets("Sheet1").Range("F2").Value = "Jan" Then
'doc.getElementById("txtFullNameAdd").Value = 1
'Elseif ThisWorbook.Sheets("Sheet1").Range("F2").Value = "Feb" Then
'doc.getElementById("txtFullNameAdd").Value = 2
'End If
'Radio buttons
'If ThisWorbook.Sheets("Sheet1").Range("F2").Value = "Jan" Then
'doc.getElementById("txtFullNameAdd").Click
'Elseif ThisWorbook.Sheets("Sheet1").Range("F2").Value = "Feb" Then
'doc.getElementById("txtFullNameAddqq").Click
'End If

End Sub

