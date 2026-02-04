# Under-Control

A comprehensive Recipe Management System featuring web interface, Excel integration, and SQLite database.

## Overview

Under Control is a complete solution for managing your recipes with three integrated components:

1. **Web Interface** - Browse and view recipes through HTML pages
2. **Excel Macros** - Import, export, and manage recipes using VBA macros
3. **Recipe Database** - SQLite database for storing and organizing recipes

## Features

### 🌐 Website
- Responsive HTML pages with modern CSS styling
- Recipe browsing and viewing
- Clean, user-friendly interface
- Mobile-friendly design

### 📊 Excel Macros
- Import recipes from database to Excel
- Export recipes to CSV format
- Add new recipes with dialog interface
- Search recipes by name or category
- Fully documented VBA code

### 🗄️ Database
- SQLite database with normalized schema
- Pre-populated with sample recipes
- Support for ingredients, instructions, and categories
- Easy to query and extend

## Quick Start

### View the Website
1. Open `index.html` in your web browser
2. Navigate to the Recipes page to view sample recipes

### Use Excel Macros
1. Open Excel
2. Import the `RecipeManager.bas` file (Alt+F11 → File → Import)
3. Run macros using Alt+F8
4. See `EXCEL_MACROS.md` for detailed instructions

### Query the Database
```bash
# View all recipes
sqlite3 recipes.db "SELECT * FROM recipes;"

# Get recipe with ingredients
sqlite3 recipes.db "SELECT * FROM recipe_details;"
```

## Project Structure

```
Under-Control/
├── index.html              # Home page
├── recipes.html            # Recipe listing page
├── styles.css              # CSS styling
├── RecipeManager.bas       # Excel VBA macros
├── recipes.db              # SQLite database
├── database_schema.sql     # Database schema and sample data
├── EXCEL_MACROS.md        # Excel macro documentation
├── DATABASE.md            # Database documentation
└── README.md              # This file
```

## Documentation

- **[Excel Macros Guide](EXCEL_MACROS.md)** - Complete guide to using the VBA macros
- **[Database Documentation](DATABASE.md)** - Database schema and usage instructions

## Sample Recipes Included

The system comes with three sample recipes:
1. **Classic Spaghetti Carbonara** - Traditional Italian pasta dish
2. **Homemade Chocolate Chip Cookies** - Classic American cookies
3. **Vegetable Stir Fry** - Healthy Asian-inspired dish

## Technologies Used

- **HTML5/CSS3** - Modern web standards
- **SQLite** - Lightweight database
- **VBA (Visual Basic for Applications)** - Excel macro development
- **SQL** - Database queries and management

## Usage Examples

### Web Interface
Simply open `index.html` in any web browser to access the recipe management interface.

### Excel Integration
```vba
' Import recipes to Excel
Sub ImportRecipes()
    Call ImportRecipesFromDatabase
End Sub

' Search for a recipe
Sub FindRecipe()
    Call SearchRecipes
End Sub
```

### Database Queries
```sql
-- Find all pasta recipes
SELECT recipe_name, prep_time_minutes, cook_time_minutes 
FROM recipes 
WHERE category = 'Pasta';

-- Get recipe with full details
SELECT r.recipe_name, i.ingredient_name, ri.quantity, ri.unit
FROM recipes r
JOIN recipe_ingredients ri ON r.recipe_id = ri.recipe_id
JOIN ingredients i ON ri.ingredient_id = i.ingredient_id
WHERE r.recipe_id = 1;
```

## Contributing

Feel free to extend this system with:
- Additional recipes
- New Excel macros
- Database views and queries
- Enhanced web interface features
- Mobile app integration

## Requirements

- **Web Interface**: Any modern web browser
- **Excel Macros**: Microsoft Excel with macro support enabled
- **Database**: SQLite3 (command-line tool or any SQLite browser)

## License

This is a personal repository for recipe management.

## Author

EdwardRaySnipes

## Getting Started with Development

### Add a New Recipe to Database
```sql
-- Add recipe
INSERT INTO recipes (recipe_name, category, prep_time_minutes, cook_time_minutes, servings)
VALUES ('My New Recipe', 'Appetizer', 15, 20, 6);

-- Add ingredients (get the recipe_id from previous insert)
INSERT INTO recipe_ingredients (recipe_id, ingredient_id, quantity, unit)
VALUES (4, 1, 200, 'grams');

-- Add instructions
INSERT INTO instructions (recipe_id, step_number, instruction_text)
VALUES (4, 1, 'First step of the recipe');
```

### Customize the Website
Edit `styles.css` to change colors, fonts, and layout. The CSS uses CSS Grid and Flexbox for responsive design.

### Extend Excel Macros
The VBA code in `RecipeManager.bas` is well-commented and easy to extend. Add new functions for additional recipe management features.

## Troubleshooting

### Macros not working in Excel
- Enable macros in Excel: File → Options → Trust Center → Macro Settings
- Ensure the file is saved as `.xlsm` (macro-enabled workbook)

### Database not found
- Ensure `recipes.db` is in the same directory as your scripts
- Use absolute paths when accessing the database programmatically

### Website not displaying correctly
- Ensure all files (HTML, CSS) are in the same directory
- Check browser console for errors
- Use a modern browser (Chrome, Firefox, Edge, Safari)

## Future Enhancements

Potential features to add:
- [ ] REST API for recipe access
- [ ] User authentication system
- [ ] Recipe ratings and reviews
- [ ] Meal planning functionality
- [ ] Shopping list generation
- [ ] Nutritional information tracking
- [ ] Photo uploads for recipes
- [ ] Recipe sharing and export features

---

For more information, see the detailed documentation files included in this repository.
