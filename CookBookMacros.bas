Attribute VB_Name = "Module1"
Sub Add_Recipie()
    Dim strInput As String
    Dim rng As Range
    Dim Template As Range
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim copiedSheet As Worksheet
    Dim newRecipie As String
    Dim subAdd As String
    
    
    newRecipie = InputBox("What Recipie Are You Adding", "Recipie Name", "Enter your input text HERE")
    
    'Exit subroutine if user cancels
    If newRecipie = "" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    

    ' Define the worksheet and the named range
    Set ws = ThisWorkbook.Sheets("Catalogue")
    Set rng = ws.Range("Recipies")
    Set Template = ws.Range("Template_Row")
    lastRow = rng.Rows.Count + rng.Row - 1
    
    'make template row visible
    If Template.EntireRow.Hidden = True Then
        Template.EntireRow.Hidden = False
    End If
    
    
    
    'copy template row format to the newly created row and make template row not visible
    Template.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Template.Copy
    Template.Offset(-1, 0).PasteSpecial
    Application.CutCopyMode = False
    Template.EntireRow.Hidden = True
    
    
    'Create and rename new sheet
    Sheets("Template").Visible = True
    Sheets("Template").Copy After:=Worksheets(Worksheets.Count)
    Set copiedSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    copiedSheet.Name = newRecipie
    Sheets("Template").Visible = False
    
    
    
    Sheets(newRecipie).Range("Title").Value = newRecipie

    'Hyperlink the recipie catalogue
    subAdd = "'" & newRecipie & "'!A1"
    ws.Hyperlinks.Add Anchor:=ws.Cells(lastRow, 1), Address:="", SubAddress:=subAdd, TextToDisplay:=newRecipie
    

    Application.ScreenUpdating = True
    
End Sub

Sub Update_Menu()
    Dim ws As Worksheet
    Dim menuSheet As Worksheet
    
    Dim lastRowBreakfast As Long, lastRowLunch As Long, lastRowDinner As Long
'    Dim breakfastChecked As Boolean, lunchChecked As Boolean, dinnerChecked As Boolean

    Dim cbBreakfast As CheckBox
    Dim cbLunch As CheckBox
    Dim cbDinner As CheckBox
    

    ' Set the menu sheet (change the name if different)
    Set menuSheet = ThisWorkbook.Sheets("Menu")

    Application.ScreenUpdating = False
    menuSheet.Visible = True

    ' Clear the previous contents in columns 1, 2, and 3
    menuSheet.Columns(1).ClearContents ' Clear breakfast column
    menuSheet.Columns(2).ClearContents ' Clear lunch column
    menuSheet.Columns(3).ClearContents ' Clear dinner column

    ' Reset row trackers for each column in the menu sheet
    lastRowBreakfast = 1 ' First row in column 1 (for breakfast)
    lastRowLunch = 1 ' First row in column 2 (for lunch)
    lastRowDinner = 1 ' First row in column 3 (for dinner)

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Skip the menu sheet itself
        If ws.Name <> "Catalogue" And ws.Name <> "Meal Planner" And ws.Name <> "Snacks" And ws.Name <> "Menu" And ws.Name <> "Template" Then
            ' Check if the combo boxes exist and retrieve their value
            On Error Resume Next ' To handle missing combo boxes without breaking the code
            Set cbBreakfast = ws.CheckBoxes("cb_breakfast")
            Set cbLunch = ws.CheckBoxes("cb_lunch")
            Set cbDinner = ws.CheckBoxes("cb_dinner")
            On Error GoTo 0 ' Resume normal error handling
                    
        
            ' Add the sheet name to the respective column in the menu sheet if combo box is checked
            If cbBreakfast.Value = xlOn Then
                menuSheet.Cells(lastRowBreakfast, 1).Value = ws.Name
                lastRowBreakfast = lastRowBreakfast + 1
            End If
            If cbLunch.Value = xlOn Then
                menuSheet.Cells(lastRowLunch, 2).Value = ws.Name
                lastRowLunch = lastRowLunch + 1
            End If
            If cbDinner.Value = xlOn Then
                menuSheet.Cells(lastRowDinner, 3).Value = ws.Name
                lastRowDinner = lastRowDinner + 1
            End If

        End If
    Next ws
    menuSheet.Visible = False
    Application.ScreenUpdating = True
End Sub

Sub Make_Grocery_List()
    Dim ws As Worksheet
    Dim groceries As Worksheet
    Dim plan As Range
'    Dim endOfList As Long
    Dim list As String

    Set ws = ThisWorkbook.Sheets("Meal Planner")
'    Set groceries = ThisWorkbook.Sheets("Grocery_List")
    Set plan = ws.Range("Meal_Plan")
    
    For Each cell In plan.Cells
        If Not IsEmpty(cell.Value) Then
            If cell.Row = 7 And cell.Column = 3 Then
                list = Get_Ingredients(cell.Value)
            ElseIf cell.Row = 9 Then
                list = list & ", " & cell.Value
            Else
                list = list & ", " & Get_Ingredients(cell.Value)
            End If
        End If
    Next cell
    
    Handle_Duplicates (list)
End Sub

Function Handle_Duplicates(list As String)
    Dim itemArr() As String
    Dim uniqueItems As Object
    Dim finalMessage As String
    Dim itemCount As String
    Dim i As Long
    
    itemArr = Split(list, ", ")
    
    Set uniqueItems = CreateObject("Scripting.Dictionary")
    
    For i = LBound(itemArr) To UBound(itemArr)
        If uniqueItems.Exists(itemArr(i)) Then
            uniqueItems(itemArr(i)) = uniqueItems(itemArr(i)) + 1
            
        Else
            uniqueItems.Add itemArr(i), 1
        End If
        
    Next i
    
    For Each Item In uniqueItems
        itemCount = uniqueItems(Item)
        If itemCount > 1 Then
            finalMessage = finalMessage & Item & " x" & itemCount & vbCrLf
        Else
            finalMessage = finalMessage & Item & vbCrLf
        End If
    Next Item
    
    
    Send_Email (finalMessage)

End Function

Function Send_Email(bodyText As String)
    Dim OutApp As Object
    Dim OutMail As Object

    ' Create the Outlook application object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Build and send the email
    With OutMail
        .To = "Enter.Email@Here.Com"
        .Subject = "Grocery List"
        .Body = bodyText
        DoEvents
        Application.Wait (Now + TimeValue("0:00:02"))
        '.Display
        .Send ' Use .Display if you want to review the email before sending
    End With
    
    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
End Function


Function Get_Ingredients(recipie As String) As String
    Dim ws As Worksheet
    Dim ing As Range
    Dim list As String
    
    Set ws = ThisWorkbook.Sheets(recipie)
    Set ing = ws.Range("Ingredients")
    
    For Each cell In ing.Cells
        If Not IsEmpty(cell.Value) Then
            If cell.Row = 8 Then
                list = cell.Value
            Else
                list = list & ", " & cell.Value
            End If
        End If
    Next cell
    
    Get_Ingredients = list
End Function

Function Get_Calories(Dish As String) As Variant
    Dim rangeValue As Variant
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(Dish)
    rangeValue = ws.Range("Calories").Value
    
    Get_Calories = rangeValue
End Function

Function Get_Protien(Dish As String) As Variant
    Dim rangeValue As Variant
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(Dish)
    rangeValue = ws.Range("Protein").Value
    
    Get_Protien = rangeValue
End Function

Function Get_Snack_Protien(snack As String) As Variant
    Get_Snack_Protien = Application.WorksheetFunction.VLookup(snack, ThisWorkbook.Sheets("Snacks").Range("Snack_Table"), 3, False)
End Function

Function Get_Total_Protein(rng As Range) As Single
    Dim tot As Single
    For Each cell In rng.Cells
        If cell.Row = 9 Then
            tot = tot + Get_Snack_Protien(cell.Value)
        Else
            tot = tot + Get_Protien(cell.Value)
        End If
    Next cell
    Get_Total_Protein = tot
End Function





