# Recipe Database Documentation

## Overview

The Under Control Recipe Management System uses SQLite as its database backend. SQLite is a lightweight, serverless database that stores all data in a single file (`recipes.db`).

## Database Schema

### Tables

#### 1. recipes
The main table storing recipe information.

| Column | Type | Description |
|--------|------|-------------|
| recipe_id | INTEGER | Primary key, auto-increment |
| recipe_name | TEXT | Name of the recipe (required) |
| category | TEXT | Recipe category (required) |
| prep_time_minutes | INTEGER | Preparation time in minutes |
| cook_time_minutes | INTEGER | Cooking time in minutes |
| servings | INTEGER | Number of servings |
| created_date | DATETIME | Timestamp when recipe was created |
| modified_date | DATETIME | Timestamp when recipe was last modified |

#### 2. ingredients
Master list of all ingredients.

| Column | Type | Description |
|--------|------|-------------|
| ingredient_id | INTEGER | Primary key, auto-increment |
| ingredient_name | TEXT | Name of the ingredient (unique) |
| unit_of_measure | TEXT | Default unit of measure |

#### 3. recipe_ingredients
Junction table linking recipes to ingredients (many-to-many relationship).

| Column | Type | Description |
|--------|------|-------------|
| recipe_ingredient_id | INTEGER | Primary key, auto-increment |
| recipe_id | INTEGER | Foreign key to recipes table |
| ingredient_id | INTEGER | Foreign key to ingredients table |
| quantity | REAL | Amount of ingredient needed |
| unit | TEXT | Unit of measure for this recipe |

#### 4. instructions
Step-by-step cooking instructions for each recipe.

| Column | Type | Description |
|--------|------|-------------|
| instruction_id | INTEGER | Primary key, auto-increment |
| recipe_id | INTEGER | Foreign key to recipes table |
| step_number | INTEGER | Order of the instruction step |
| instruction_text | TEXT | The instruction text |

#### 5. categories
Lookup table for recipe categories.

| Column | Type | Description |
|--------|------|-------------|
| category_id | INTEGER | Primary key, auto-increment |
| category_name | TEXT | Name of the category (unique) |
| description | TEXT | Category description |

### Views

#### recipe_details
A simplified view for querying basic recipe information.

```sql
SELECT * FROM recipe_details;
```

Returns: recipe_id, recipe_name, category, prep_time_minutes, cook_time_minutes, servings, created_date, modified_date

## Sample Queries

### Get all recipes
```sql
SELECT recipe_id, recipe_name, category 
FROM recipes 
ORDER BY recipe_name;
```

### Get a recipe with its ingredients
```sql
SELECT 
    r.recipe_name,
    i.ingredient_name,
    ri.quantity,
    ri.unit
FROM recipes r
JOIN recipe_ingredients ri ON r.recipe_id = ri.recipe_id
JOIN ingredients i ON ri.ingredient_id = i.ingredient_id
WHERE r.recipe_id = 1;
```

### Get a recipe with instructions
```sql
SELECT 
    r.recipe_name,
    inst.step_number,
    inst.instruction_text
FROM recipes r
JOIN instructions inst ON r.recipe_id = inst.recipe_id
WHERE r.recipe_id = 1
ORDER BY inst.step_number;
```

### Search recipes by category
```sql
SELECT recipe_name, prep_time_minutes, cook_time_minutes 
FROM recipes 
WHERE category = 'Pasta';
```

### Find recipes by ingredient
```sql
SELECT DISTINCT r.recipe_name, r.category
FROM recipes r
JOIN recipe_ingredients ri ON r.recipe_id = ri.recipe_id
JOIN ingredients i ON ri.ingredient_id = i.ingredient_id
WHERE i.ingredient_name LIKE '%chicken%';
```

### Get all categories
```sql
SELECT category_name, description 
FROM categories 
ORDER BY category_name;
```

## Adding New Data

### Insert a new recipe
```sql
INSERT INTO recipes (recipe_name, category, prep_time_minutes, cook_time_minutes, servings)
VALUES ('New Recipe Name', 'Main Course', 20, 30, 4);
```

### Add an ingredient
```sql
INSERT INTO ingredients (ingredient_name, unit_of_measure)
VALUES ('chicken breast', 'pounds');
```

### Link ingredient to recipe
```sql
INSERT INTO recipe_ingredients (recipe_id, ingredient_id, quantity, unit)
VALUES (1, 5, 2, 'pounds');
```

### Add cooking instructions
```sql
INSERT INTO instructions (recipe_id, step_number, instruction_text)
VALUES (1, 1, 'Preheat the oven to 350°F');
```

## Using the Database

### Command Line
```bash
# Open the database
sqlite3 recipes.db

# Run a query
sqlite3 recipes.db "SELECT * FROM recipes;"

# Import SQL file
sqlite3 recipes.db < database_schema.sql

# Export to CSV
sqlite3 -header -csv recipes.db "SELECT * FROM recipes;" > recipes.csv
```

### Python Integration
```python
import sqlite3

# Connect to database
conn = sqlite3.connect('recipes.db')
cursor = conn.cursor()

# Query recipes
cursor.execute("SELECT * FROM recipes")
recipes = cursor.fetchall()

for recipe in recipes:
    print(recipe)

# Close connection
conn.close()
```

### Excel Integration
The VBA macros in `RecipeManager.bas` provide Excel integration. See `EXCEL_MACROS.md` for details.

## Database Maintenance

### Backup
```bash
# Create a backup
cp recipes.db recipes_backup_$(date +%Y%m%d).db

# Or use SQLite backup
sqlite3 recipes.db ".backup recipes_backup.db"
```

### Optimize
```sql
-- Rebuild the database file to reclaim unused space
VACUUM;

-- Update statistics for query optimizer
ANALYZE;
```

### Check integrity
```sql
-- Check database integrity
PRAGMA integrity_check;
```

## Schema Updates

The schema is defined in `database_schema.sql`. To recreate the database:

```bash
rm recipes.db
sqlite3 recipes.db < database_schema.sql
```

**Warning:** This will delete all existing data.

## Security Considerations

- The database file (`recipes.db`) contains all recipe data
- Set appropriate file permissions to restrict access
- For web applications, use parameterized queries to prevent SQL injection
- Consider encrypting the database file if storing sensitive information

## Sample Data

The database comes pre-populated with three sample recipes:
1. Classic Spaghetti Carbonara (Pasta)
2. Homemade Chocolate Chip Cookies (Dessert)
3. Vegetable Stir Fry (Main Course)

These can be used as templates for adding your own recipes.

## Performance Tips

- Indexes are already created on commonly queried columns
- Use the `recipe_details` view for simple queries
- For complex queries involving multiple joins, consider creating additional views
- Use EXPLAIN QUERY PLAN to analyze query performance

## Further Resources

- [SQLite Documentation](https://www.sqlite.org/docs.html)
- [SQLite Tutorial](https://www.sqlitetutorial.net/)
- [Python SQLite3 Module](https://docs.python.org/3/library/sqlite3.html)
