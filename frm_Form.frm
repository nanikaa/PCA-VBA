VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Form 
   Caption         =   "New Data Entry"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18360
   OleObjectBlob   =   "frm_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Submit_Click()
    Dim ws As Worksheet
    Dim summaryWS As Worksheet
    Dim mailVolWS As Worksheet
    
    Dim lastRow As Long
    Dim lastRow_summary As Long
    
    Dim revType As String
    Dim totalAmount As Double
    Dim mailVolume As Double
    Dim dayValue As Integer
    Dim inputDate As Date
    Dim current_pieces As Double
    Dim current_amount As Double
    
    Dim regCol As Integer
    Dim rrrCol As Integer
    Dim demsCol As Integer
    Dim ordCol As Integer
    Dim fgnCol As Integer

    Dim response As VbMsgBoxResult ' Variable to store the user's response

    ' Set reference to the worksheets
    Set ws = ThisWorkbook.Sheets("Database")
    Set summaryWS = ThisWorkbook.Sheets("SUMMARY")
    Set mailVolWS = ThisWorkbook.Sheets("MAIL VOLUME")
    
    ' Check if required fields are filled
    If combo_RegName.Value = "" Then
        MsgBox "Please select a value for 'RegName'.", vbExclamation
        Exit Sub
    End If

    ' Confirm submission
    response = MsgBox("Are you sure you want to submit this form?", vbYesNo + vbQuestion, "Confirm Submission")
    If response = vbNo Then
        Exit Sub ' Exit the subroutine if the user selects 'No'
    End If
    
    ' Initialize the revType variable
    revType = ""
    totalAmount = 0
    mailVolume = 0
    checkCounter = 0
    current_pieces = 0
    current_amount = 0
    inputDate = DateValue(Format(txtBox_Date.Value, "mm/dd/yyyy"))
    
    ' Extract the day value from the date
    dayValue = Day(inputDate)
    
    ' Find the last used row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Assign values to the cells in the new row
    ws.Cells(lastRow, "A").Value = txtBox_Date.Value
    ws.Cells(lastRow, "B").Value = combo_Abbrev.Value
    ws.Cells(lastRow, "C").Value = combo_RegName.Value

    ' Refactored code for setting column indices
    Dim abbreviations As Variant
    Dim baseCol As Integer
    Dim startIndex As Variant

    abbreviations = Array("BIR", "CHED", "CSC", "DAR", "DBM", "DHSUD", "DSWD", "GSIS", "PHIL-H", "HDMF", "NLRC", "DEPED", "SSS", "DOH")
    startIndex = Application.Match(combo_Abbrev.Value, abbreviations, 0)

    If IsError(startIndex) Then
        MsgBox "Abbreviation not recognized.", vbExclamation
        Exit Sub
    Else
        baseCol = 2 + (startIndex - 1) * 14
        regCol = baseCol
        rrrCol = baseCol + 2
        demsCol = baseCol + 4
        ordCol = baseCol + 6
        fgnCol = baseCol + 8
    End If
    
    ' Insert values into worksheets based on selected checkboxes
    InsertValuesIntoWorksheets lastRow, ws, summaryWS, dayValue
    
    ' Calculate mail volume and total amount
    mailVolume = CDbl(ws.Cells(lastRow, "D").Value) + CDbl(ws.Cells(lastRow, "F").Value) + CDbl(ws.Cells(lastRow, "H").Value) + CDbl(ws.Cells(lastRow, "J").Value) + CDbl(ws.Cells(lastRow, "L").Value)
    totalAmount = CDbl(ws.Cells(lastRow, "E").Value) + CDbl(ws.Cells(lastRow, "G").Value) + CDbl(ws.Cells(lastRow, "I").Value) + CDbl(ws.Cells(lastRow, "K").Value) + CDbl(ws.Cells(lastRow, "M").Value)
    
    ' Call the UpdateDailySalesPosting function
    UpdateDailySalesPosting dayValue, regCol, rrrCol, demsCol, ordCol, fgnCol
    
    ' Notify the user that the data has been appended
    MsgBox "Data has been appended successfully!", vbInformation
    
    ' Close the form
    Unload Me
End Sub

Private Sub UpdateDailySalesPosting(dayValue As Integer, regCol As Integer, rrrCol As Integer, demsCol As Integer, ordCol As Integer, fgnCol As Integer)
    Dim dspWS As Worksheet
    Dim rowContainingValue As Long
    
    Set dspWS = ThisWorkbook.Sheets("DAILY SALES POSTING")
    
    rowContainingValue = FindRow(dayValue, dspWS.Columns("A"))
    If rowContainingValue <> -1 Then
        If checkBox_Ordinary.Value = True Then
        ' Insert in DSP Worksheet (ORDINARY)
            dspWS.Cells(rowContainingValue, ordCol).Value = IIf(txtBox_Pieces_Ordinary.Value = "", 0, txtBox_Pieces_Ordinary.Value)
            dspWS.Cells(rowContainingValue, ordCol + 1).Value = IIf(txtBox_Total_Ordinary.Value = "", 0, txtBox_Total_Ordinary.Value)
        End If
    
        If checkBox_Reg.Value = True Then
        ' Insert in DSP Worksheet (REGISTERED)
            dspWS.Cells(rowContainingValue, regCol).Value = IIf(txtBox_Pieces_Reg.Value = "", 0, txtBox_Pieces_Reg.Value)
            dspWS.Cells(rowContainingValue, regCol + 1).Value = IIf(txtBox_Total_Reg.Value = "", 0, txtBox_Total_Reg.Value)
        End If
        
        If checkBox_RegRRR.Value = True Then
        ' Insert in DSP Worksheet (REGISTERED W/ RRR)
            dspWS.Cells(rowContainingValue, rrrCol).Value = IIf(txtBox_Pieces_RegRRR.Value = "", 0, txtBox_Pieces_RegRRR.Value)
            dspWS.Cells(rowContainingValue, rrrCol + 1).Value = IIf(txtBox_Total_RegRRR.Value = "", 0, txtBox_Total_RegRRR.Value)
        End If
        
        If checkBox_DEMS.Value = True Then
        ' Insert in DSP Worksheet (DEMS)
            dspWS.Cells(rowContainingValue, demsCol).Value = IIf(txtBox_Pieces_DEMS.Value = "", 0, txtBox_Pieces_DEMS.Value)
            dspWS.Cells(rowContainingValue, demsCol + 1).Value = IIf(txtBox_Total_DEMS.Value = "", 0, txtBox_Total_DEMS.Value)
        End If
        
        If checkBox_ForeignReg.Value = True Then
        ' Insert in DSP Worksheet (FOREIGN REGISTERED)
            dspWS.Cells(rowContainingValue, fgnCol).Value = IIf(txtBox_Pieces_ForeignReg.Value = "", 0, txtBox_Pieces_ForeignReg.Value)
            dspWS.Cells(rowContainingValue, fgnCol + 1).Value = IIf(txtBox_Total_ForeignReg.Value = "", 0, txtBox_Total_ForeignReg.Value)
        End If
    Else
        MsgBox "Date not found in column A - DAILY SALES POSTING."
    End If
End Sub
Private Sub InsertValuesIntoWorksheets(ByVal lastRow As Long, ByVal ws As Worksheet, ByVal summaryWS As Worksheet, ByVal dayValue As Integer)
    Dim revType As String
    Dim totalAmount As Double
    Dim lastRow_summary As Long
    
    ' Initialize the revType variable
    revType = ""
    totalAmount = 0

    If checkBox_Ordinary.Value = True Then
        revType = "ORDINARY"
        
        ' Insert in Database Sheet
        ws.Cells(lastRow, "J").Value = IIf(txtBox_Pieces_Ordinary.Value = "", 0, txtBox_Pieces_Ordinary.Value)
        ws.Cells(lastRow, "K").Value = IIf(txtBox_Total_Ordinary.Value = "", 0, txtBox_Total_Ordinary.Value)
        
        
        ' Insert in SUMMARY Worksheet
        lastRow_summary = summaryWS.Cells(summaryWS.Rows.Count, "B").End(xlUp).Row + 1
        summaryWS.Cells(lastRow_summary, "B").Value = txtBox_Date.Value
        summaryWS.Cells(lastRow_summary, "H").Value = combo_RegName.Value
        summaryWS.Cells(lastRow_summary, "AB").Value = txtBox_Pieces_Ordinary.Value
        summaryWS.Cells(lastRow_summary, "M").Value = txtBox_Total_Ordinary.Value
        summaryWS.Cells(lastRow_summary, "Z").Value = revType
    End If
    
    If checkBox_Reg.Value = True Then
        revType = "REGISTERED MAILS-MAIL MATTERS-DOMESTIC"
        
        ' Insert in Database Sheet
        ws.Cells(lastRow, "D").Value = IIf(txtBox_Pieces_Reg.Value = "", 0, txtBox_Pieces_Reg.Value)
        ws.Cells(lastRow, "E").Value = IIf(txtBox_Total_Reg.Value = "", 0, txtBox_Total_Reg.Value)
        
        ' Insert in SUMMARY Worksheet
        lastRow_summary = summaryWS.Cells(summaryWS.Rows.Count, "B").End(xlUp).Row + 1
        summaryWS.Cells(lastRow_summary, "B").Value = txtBox_Date.Value
        summaryWS.Cells(lastRow_summary, "H").Value = combo_RegName.Value
        summaryWS.Cells(lastRow_summary, "AB").Value = txtBox_Pieces_Reg.Value
        summaryWS.Cells(lastRow_summary, "M").Value = txtBox_Total_Reg.Value
        summaryWS.Cells(lastRow_summary, "Z").Value = revType
    End If
    
    If checkBox_RegRRR.Value = True Then
        revType = "REGISTERED MAILS WITH RETURN CARDS-MAIL MATTERS-DOMESTIC"
        
        ' Insert in Database Sheet
        ws.Cells(lastRow, "F").Value = IIf(txtBox_Pieces_RegRRR.Value = "", 0, txtBox_Pieces_RegRRR.Value)
        ws.Cells(lastRow, "G").Value = IIf(txtBox_Total_RegRRR.Value = "", 0, txtBox_Total_RegRRR.Value)
                
        ' Insert in SUMMARY Worksheet
        lastRow_summary = summaryWS.Cells(summaryWS.Rows.Count, "B").End(xlUp).Row + 1
        summaryWS.Cells(lastRow_summary, "B").Value = txtBox_Date.Value
        summaryWS.Cells(lastRow_summary, "H").Value = combo_RegName.Value
        summaryWS.Cells(lastRow_summary, "AB").Value = txtBox_Pieces_RegRRR.Value
        summaryWS.Cells(lastRow_summary, "M").Value = txtBox_Total_RegRRR.Value
        summaryWS.Cells(lastRow_summary, "Z").Value = revType
    End If
    
    If checkBox_DEMS.Value = True Then
        revType = "EMS - DOCUMENT-MAIL MATTERS-DOMESTIC"
        
        ' Insert in Database Sheet
        ws.Cells(lastRow, "H").Value = IIf(txtBox_Pieces_DEMS.Value = "", 0, txtBox_Pieces_DEMS.Value)
        ws.Cells(lastRow, "I").Value = IIf(txtBox_Total_DEMS.Value = "", 0, txtBox_Total_DEMS.Value)
        
        ' Insert in SUMMARY Worksheet
        lastRow_summary = summaryWS.Cells(summaryWS.Rows.Count, "B").End(xlUp).Row + 1
        summaryWS.Cells(lastRow_summary, "B").Value = txtBox_Date.Value
        summaryWS.Cells(lastRow_summary, "H").Value = combo_RegName.Value
        summaryWS.Cells(lastRow_summary, "AB").Value = txtBox_Pieces_DEMS.Value
        summaryWS.Cells(lastRow_summary, "M").Value = txtBox_Total_DEMS.Value
        summaryWS.Cells(lastRow_summary, "Z").Value = revType
    End If
    
    If checkBox_ForeignReg.Value = True Then
        revType = "FOREIGN REGISTERED"
        
        ' Insert in Database Sheet
        ws.Cells(lastRow, "L").Value = IIf(txtBox_Pieces_ForeignReg.Value = "", 0, txtBox_Pieces_ForeignReg.Value)
        ws.Cells(lastRow, "M").Value = IIf(txtBox_Total_ForeignReg.Value = "", 0, txtBox_Total_ForeignReg.Value)
        
        ' Insert in SUMMARY Worksheet
        lastRow_summary = summaryWS.Cells(summaryWS.Rows.Count, "B").End(xlUp).Row + 1
        summaryWS.Cells(lastRow_summary, "B").Value = txtBox_Date.Value
        summaryWS.Cells(lastRow_summary, "H").Value = combo_RegName.Value
        summaryWS.Cells(lastRow_summary, "AB").Value = txtBox_Pieces_ForeignReg.Value
        summaryWS.Cells(lastRow_summary, "M").Value = txtBox_Total_ForeignReg.Value
        summaryWS.Cells(lastRow_summary, "Z").Value = revType
    End If
End Sub

Function FindRow(ByVal searchValue As String, ByVal searchRange As Range) As Long
    Dim foundCell As Range
    
    ' Define the search range to start from row 3
    Set searchRange = searchRange.Resize(searchRange.Rows.Count - 3).Offset(3)
    
    ' Use the Find method to search for the value in the adjusted search range
    Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the value was found
    If Not foundCell Is Nothing Then
        FindRow = foundCell.Row ' Corrected line: Return the row number
    Else
        FindRow = -1 ' Return -1 if value not found
    End If
End Function

Private Sub UserForm_Initialize()
    Dim rng_RegName As Range
    Dim rng_Abbrev As Range
    Dim cell_RegName As Range
    Dim cell_Abbrev As Range
    Dim ws As Worksheet
    
    ' Set reference to the "Database" sheet
    Set ws = ThisWorkbook.Sheets("Database")

    
    ' Load data into combo_RegName
    Set rng_RegName = ws.Range("S16", ws.Cells(ws.Rows.Count, "S").End(xlUp))
    For Each cell_RegName In rng_RegName
        If cell_RegName.Value <> "" Then ' Exclude empty cells
            combo_RegName.AddItem cell_RegName.Value
            ' Find corresponding abbreviation in column I and add it to the Tag property
            combo_RegName.List(combo_RegName.ListCount - 1, 1) = ws.Cells(cell_RegName.Row, "R").Value
        End If
    Next cell_RegName
    
    ' Load data into combo_Abbrev
    Set rng_Abbrev = ws.Range("R16", ws.Cells(ws.Rows.Count, "R").End(xlUp))
    For Each cell_Abbrev In rng_Abbrev
        If cell_Abbrev.Value <> "" Then ' Exclude empty cells
            combo_Abbrev.AddItem cell_Abbrev.Value
        End If
    Next cell_Abbrev
    
    ' Set the date textbox to today's date
    txtBox_Date.Value = Date
End Sub

Private Sub combo_RegName_Change()
    ' Update combo_Abbrev when a new item is selected in combo_RegName
    Dim selectedItemIndex As Integer
    Dim foundIndex As Integer
    
    ' Get the index of the selected item in combo_RegName
    selectedItemIndex = combo_RegName.ListIndex
    
    ' If an item is selected
    If selectedItemIndex >= 0 Then
        ' Update the value of combo_Abbrev based on the selected item in combo_RegName
        combo_Abbrev.Value = combo_RegName.List(selectedItemIndex, 1)
    End If
End Sub

Private Sub combo_Abbrev_Change()
    ' Update combo_RegName when a new item is selected in combo_Abbrev
    Dim selectedItemIndex As Integer
    Dim foundIndex As Integer
    
    ' Get the index of the selected item in combo_Abbrev
    selectedItemIndex = combo_Abbrev.ListIndex
    
    ' If an item is selected
    If selectedItemIndex >= 0 Then
        ' Find the index of the corresponding item in combo_RegName
        foundIndex = -1
        For i = 0 To combo_RegName.ListCount - 1
            If combo_RegName.List(i, 1) = combo_Abbrev.Value Then
                foundIndex = i
                Exit For
            End If
        Next i
        
        ' If the corresponding item is found, select it in combo_RegName
        If foundIndex >= 0 Then
            combo_RegName.ListIndex = foundIndex
        Else
            ' If the corresponding item is not found, clear the selection in combo_RegName
            combo_RegName.Value = ""
        End If
    End If
End Sub
' Checkbox validation (ORDINARY)
Private Sub txtBox_Pieces_Ordinary_Change()
    Call CheckOrdinaryValues
End Sub

Private Sub txtBox_Total_Ordinary_Change()
    Call CheckOrdinaryValues
End Sub

Private Sub CheckOrdinaryValues()
    ' Check if both txtBox_Pieces_Ordinary and txtBox_Total_Ordinary have values
    If Len(txtBox_Pieces_Ordinary.Value) > 0 And Len(txtBox_Total_Ordinary.Value) > 0 Then
        ' If both have values, set checkbox_Ordinary to True
        checkBox_Ordinary.Value = True
    Else
        ' Optional: clear or keep the checkbox unchecked if either box is empty
        checkBox_Ordinary.Value = False
    End If
End Sub

' Checkbox validation (REGULAR)
Private Sub txtBox_Pieces_Reg_Change()
    Call CheckRegValues
End Sub
Private Sub txtBox_Total_Reg_Change()
    Call CheckRegValues
End Sub
Private Sub CheckRegValues()
    If Len(txtBox_Pieces_Reg.Value) > 0 And Len(txtBox_Total_Reg.Value) > 0 Then
        checkBox_Reg.Value = True
    Else
        checkBox_Reg.Value = False
    End If
End Sub

' Checkbox validation (REGULAR W RRR)
Private Sub txtBox_Pieces_RegRRR_Change()
    Call CheckRegRRRValues
End Sub

Private Sub txtBox_Total_RegRRR_Change()
    Call CheckRegRRRValues
End Sub

Private Sub CheckRegRRRValues()
    If Len(txtBox_Pieces_RegRRR.Value) > 0 And Len(txtBox_Total_RegRRR.Value) > 0 Then
        checkBox_RegRRR.Value = True
    Else
        checkBox_RegRRR.Value = False
    End If
End Sub

' Checkbox validation (DEMS)
Private Sub txtBox_Pieces_DEMS_Change()
    Call CheckDEMSValues
End Sub
Private Sub txtBox_Total_DEMS_Change()
    Call CheckDEMSValues
End Sub
Private Sub CheckDEMSValues()
    If Len(txtBox_Pieces_DEMS.Value) > 0 And Len(txtBox_Total_DEMS.Value) > 0 Then
        checkBox_DEMS.Value = True
    Else
        checkBox_DEMS.Value = False
    End If
End Sub

' Checkbox validation (FOREIGN REG)
Private Sub txtBox_Pieces_ForeignReg_Change()
    Call CheckForeignRegValues
End Sub
Private Sub txtBox_Total_ForeignReg_Change()
    Call CheckForeignRegValues
End Sub
Private Sub CheckForeignRegValues()
    If Len(txtBox_Pieces_ForeignReg.Value) > 0 And Len(txtBox_Total_ForeignReg.Value) > 0 Then
        checkBox_ForeignReg.Value = True
    Else
        checkBox_ForeignReg.Value = False
    End If
End Sub



