# Excel Macros for Recipe Management

This document describes how to use the Excel VBA macros included in the Under Control Recipe Management System.

## Overview

The `RecipeManager.bas` module contains VBA macros that enable you to:
- Import recipes from the database into Excel
- Export recipes from Excel to CSV format
- Add new recipes directly in Excel
- Search for recipes by name or category

## Installation

1. Open Excel
2. Press `Alt + F11` to open the VBA Editor
3. Go to `File > Import File`
4. Select the `RecipeManager.bas` file
5. Close the VBA Editor

## Available Macros

### ImportRecipesFromDatabase

**Description:** Imports recipes from the database into a new or existing "Recipes" worksheet.

**How to use:**
1. Press `Alt + F8` to open the Macro dialog
2. Select `ImportRecipesFromDatabase`
3. Click `Run`

The macro will create a properly formatted worksheet with recipe data including:
- Recipe ID
- Recipe Name
- Category
- Prep Time
- Cook Time
- Servings
- Ingredients
- Instructions

### ExportRecipesToCSV

**Description:** Exports all recipes from the Excel worksheet to a CSV file that can be imported into other systems.

**How to use:**
1. Ensure you have imported recipes first using `ImportRecipesFromDatabase`
2. Press `Alt + F8` to open the Macro dialog
3. Select `ExportRecipesToCSV`
4. Click `Run`
5. Choose a location and filename for the CSV file
6. Click `Save`

### AddNewRecipe

**Description:** Adds a new recipe to the worksheet with a simple dialog interface.

**How to use:**
1. Press `Alt + F8` to open the Macro dialog
2. Select `AddNewRecipe`
3. Click `Run`
4. Enter the recipe name when prompted
5. Enter the recipe category when prompted
6. Fill in the remaining details (prep time, cook time, ingredients, instructions) in the worksheet

### SearchRecipes

**Description:** Search for recipes by name or category.

**How to use:**
1. Press `Alt + F8` to open the Macro dialog
2. Select `SearchRecipes`
3. Click `Run`
4. Enter your search term (recipe name or category)
5. View the results in a message box

## Quick Access Buttons (Optional)

You can add buttons to your worksheet for quick access to these macros:

1. Go to `Developer > Insert > Button (Form Control)`
2. Draw a button on your worksheet
3. Assign a macro to the button when prompted
4. Right-click the button and choose `Edit Text` to label it

Recommended buttons:
- "Import Recipes" → `ImportRecipesFromDatabase`
- "Export to CSV" → `ExportRecipesToCSV`
- "Add Recipe" → `AddNewRecipe`
- "Search" → `SearchRecipes`

## Keyboard Shortcuts

You can assign keyboard shortcuts to frequently used macros:

1. Press `Alt + F8` to open the Macro dialog
2. Select a macro
3. Click `Options`
4. Enter a shortcut key (e.g., `Ctrl + Shift + I` for Import)
5. Click `OK`

## Troubleshooting

### "No Recipes worksheet found" error
- Run the `ImportRecipesFromDatabase` macro first to create the worksheet

### Macros are disabled
- Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`
- Select "Enable all macros" (note: only do this for trusted workbooks)

### Cannot import file
- Ensure the file has a `.bas` extension
- Check that macros are enabled in your Excel settings

## Data Format

The worksheet uses the following column structure:

| Column | Field | Type | Description |
|--------|-------|------|-------------|
| A | Recipe ID | Number | Unique identifier for the recipe |
| B | Recipe Name | Text | Name of the recipe |
| C | Category | Text | Category (Pasta, Dessert, Main Course, etc.) |
| D | Prep Time | Number | Preparation time in minutes |
| E | Cook Time | Number | Cooking time in minutes |
| F | Servings | Number | Number of servings |
| G | Ingredients | Text | Comma-separated list of ingredients |
| H | Instructions | Text | Step-by-step instructions |

## Tips

- Keep your recipes worksheet backed up regularly
- Use consistent category names for better search results
- Export to CSV periodically to maintain a backup
- You can sort and filter the data using Excel's standard tools

## Integration with Database

In a production environment, these macros would connect directly to the SQLite database using:
- ODBC connections
- Third-party SQLite libraries for VBA
- Or external scripts that sync Excel with the database

The current implementation uses sample data for demonstration purposes.
