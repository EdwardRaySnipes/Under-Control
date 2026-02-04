Attribute VB_Name = "RecipeManager"
Option Explicit

' Recipe Manager Excel VBA Macros
' Under Control - Recipe Management System
' 
' This module provides macros for importing and exporting recipes
' between Excel and the recipe database

' Import recipes from database to Excel worksheet
Sub ImportRecipesFromDatabase()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Create or clear the Recipes worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Recipes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Recipes"
    Else
        ws.Cells.Clear
    End If
    
    ' Set up headers
    ws.Range("A1").Value = "Recipe ID"
    ws.Range("B1").Value = "Recipe Name"
    ws.Range("C1").Value = "Category"
    ws.Range("D1").Value = "Prep Time (min)"
    ws.Range("E1").Value = "Cook Time (min)"
    ws.Range("F1").Value = "Servings"
    ws.Range("G1").Value = "Ingredients"
    ws.Range("H1").Value = "Instructions"
    
    ' Format headers
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(102, 126, 234)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Add sample data (in a real scenario, this would connect to the SQLite database)
    lastRow = 2
    
    ' Recipe 1: Spaghetti Carbonara
    ws.Cells(lastRow, 1).Value = 1
    ws.Cells(lastRow, 2).Value = "Classic Spaghetti Carbonara"
    ws.Cells(lastRow, 3).Value = "Pasta"
    ws.Cells(lastRow, 4).Value = 10
    ws.Cells(lastRow, 5).Value = 15
    ws.Cells(lastRow, 6).Value = 4
    ws.Cells(lastRow, 7).Value = "400g spaghetti, 200g pancetta, 4 eggs, 100g Parmesan cheese, Black pepper"
    ws.Cells(lastRow, 8).Value = "Cook spaghetti; Fry pancetta; Beat eggs with Parmesan; Toss pasta with pancetta and eggs; Season and serve"
    
    lastRow = lastRow + 1
    
    ' Recipe 2: Chocolate Chip Cookies
    ws.Cells(lastRow, 1).Value = 2
    ws.Cells(lastRow, 2).Value = "Homemade Chocolate Chip Cookies"
    ws.Cells(lastRow, 3).Value = "Dessert"
    ws.Cells(lastRow, 4).Value = 15
    ws.Cells(lastRow, 5).Value = 12
    ws.Cells(lastRow, 6).Value = 24
    ws.Cells(lastRow, 7).Value = "2 1/4 cups flour, 1 cup butter, 3/4 cup sugar, 3/4 cup brown sugar, 2 eggs, 2 cups chocolate chips, 1 tsp vanilla, 1 tsp baking soda, 1/2 tsp salt"
    ws.Cells(lastRow, 8).Value = "Preheat oven to 375°F; Cream butter and sugars; Beat in eggs and vanilla; Mix in dry ingredients; Stir in chocolate chips; Bake 9-11 minutes"
    
    lastRow = lastRow + 1
    
    ' Recipe 3: Vegetable Stir Fry
    ws.Cells(lastRow, 1).Value = 3
    ws.Cells(lastRow, 2).Value = "Vegetable Stir Fry"
    ws.Cells(lastRow, 3).Value = "Main Course"
    ws.Cells(lastRow, 4).Value = 15
    ws.Cells(lastRow, 5).Value = 10
    ws.Cells(lastRow, 6).Value = 4
    ws.Cells(lastRow, 7).Value = "2 cups broccoli, 1 bell pepper, 1 cup snap peas, 2 carrots, 3 cloves garlic, 2 tbsp soy sauce, 1 tbsp sesame oil, 1 tsp ginger"
    ws.Cells(lastRow, 8).Value = "Heat oil in wok; Add garlic and ginger; Stir-fry vegetables 5-7 minutes; Add soy sauce; Serve over rice"
    
    ' Auto-fit columns
    ws.Columns("A:H").AutoFit
    
    MsgBox "Recipes imported successfully! " & (lastRow - 1) & " recipes loaded.", vbInformation, "Import Complete"
End Sub

' Export recipes from Excel to database format (CSV)
Sub ExportRecipesToCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim csvLine As String
    
    ' Get the Recipes worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Recipes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No Recipes worksheet found. Please import recipes first.", vbExclamation, "Export Failed"
        Exit Sub
    End If
    
    ' Get file path from user
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="recipes_export.csv", _
        FileFilter:="CSV Files (*.csv), *.csv", _
        Title:="Export Recipes to CSV")
    
    If filePath = "False" Then Exit Sub ' User cancelled
    
    ' Open file for writing
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write header row
    csvLine = ""
    For i = 1 To 8
        If i > 1 Then csvLine = csvLine & ","
        csvLine = csvLine & """" & ws.Cells(1, i).Value & """"
    Next i
    Print #fileNum, csvLine
    
    ' Write data rows
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        csvLine = ""
        csvLine = csvLine & ws.Cells(i, 1).Value & ","
        csvLine = csvLine & """" & ws.Cells(i, 2).Value & ""","
        csvLine = csvLine & """" & ws.Cells(i, 3).Value & ""","
        csvLine = csvLine & ws.Cells(i, 4).Value & ","
        csvLine = csvLine & ws.Cells(i, 5).Value & ","
        csvLine = csvLine & ws.Cells(i, 6).Value & ","
        csvLine = csvLine & """" & Replace(ws.Cells(i, 7).Value, """", """""") & ""","
        csvLine = csvLine & """" & Replace(ws.Cells(i, 8).Value, """", """""") & """"
        
        Print #fileNum, csvLine
    Next i
    
    ' Close file
    Close #fileNum
    
    MsgBox "Recipes exported successfully to:" & vbCrLf & filePath, vbInformation, "Export Complete"
End Sub

' Add a new recipe to the worksheet
Sub AddNewRecipe()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim recipeName As String
    Dim category As String
    
    ' Get the Recipes worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Recipes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No Recipes worksheet found. Please import recipes first.", vbExclamation, "Add Recipe Failed"
        Exit Sub
    End If
    
    ' Get recipe information from user
    recipeName = InputBox("Enter recipe name:", "Add New Recipe")
    If recipeName = "" Then Exit Sub
    
    category = InputBox("Enter recipe category (e.g., Pasta, Dessert, Main Course):", "Add New Recipe")
    If category = "" Then Exit Sub
    
    ' Find the next empty row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Add new recipe
    ws.Cells(lastRow, 1).Value = lastRow - 1 ' Recipe ID
    ws.Cells(lastRow, 2).Value = recipeName
    ws.Cells(lastRow, 3).Value = category
    ws.Cells(lastRow, 4).Value = 0 ' User should fill in prep time
    ws.Cells(lastRow, 5).Value = 0 ' User should fill in cook time
    ws.Cells(lastRow, 6).Value = 1 ' Default servings
    ws.Cells(lastRow, 7).Value = "Enter ingredients here"
    ws.Cells(lastRow, 8).Value = "Enter instructions here"
    
    ' Select the new row for editing
    ws.Range("A" & lastRow & ":H" & lastRow).Select
    
    MsgBox "New recipe added! Please fill in the remaining details.", vbInformation, "Recipe Added"
End Sub

' Search for recipes by name or category
Sub SearchRecipes()
    Dim ws As Worksheet
    Dim searchTerm As String
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    Dim results As String
    
    ' Get the Recipes worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Recipes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No Recipes worksheet found. Please import recipes first.", vbExclamation, "Search Failed"
        Exit Sub
    End If
    
    ' Get search term from user
    searchTerm = InputBox("Enter recipe name or category to search for:", "Search Recipes")
    If searchTerm = "" Then Exit Sub
    
    searchTerm = LCase(searchTerm)
    found = False
    results = "Search Results for: " & searchTerm & vbCrLf & vbCrLf
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If InStr(1, LCase(ws.Cells(i, 2).Value), searchTerm) > 0 Or _
           InStr(1, LCase(ws.Cells(i, 3).Value), searchTerm) > 0 Then
            found = True
            results = results & ws.Cells(i, 2).Value & " (" & ws.Cells(i, 3).Value & ")" & vbCrLf
        End If
    Next i
    
    If found Then
        MsgBox results, vbInformation, "Search Results"
    Else
        MsgBox "No recipes found matching: " & searchTerm, vbInformation, "No Results"
    End If
End Sub
