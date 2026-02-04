-- Recipe Database Schema
-- Under Control - Recipe Management System
-- SQLite Database Schema

-- Create recipes table
CREATE TABLE IF NOT EXISTS recipes (
    recipe_id INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_name TEXT NOT NULL,
    category TEXT NOT NULL,
    prep_time_minutes INTEGER,
    cook_time_minutes INTEGER,
    servings INTEGER,
    created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
    modified_date DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Create ingredients table
CREATE TABLE IF NOT EXISTS ingredients (
    ingredient_id INTEGER PRIMARY KEY AUTOINCREMENT,
    ingredient_name TEXT NOT NULL UNIQUE,
    unit_of_measure TEXT
);

-- Create recipe_ingredients junction table (many-to-many relationship)
CREATE TABLE IF NOT EXISTS recipe_ingredients (
    recipe_ingredient_id INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_id INTEGER NOT NULL,
    ingredient_id INTEGER NOT NULL,
    quantity REAL,
    unit TEXT,
    FOREIGN KEY (recipe_id) REFERENCES recipes(recipe_id) ON DELETE CASCADE,
    FOREIGN KEY (ingredient_id) REFERENCES ingredients(ingredient_id)
);

-- Create instructions table
CREATE TABLE IF NOT EXISTS instructions (
    instruction_id INTEGER PRIMARY KEY AUTOINCREMENT,
    recipe_id INTEGER NOT NULL,
    step_number INTEGER NOT NULL,
    instruction_text TEXT NOT NULL,
    FOREIGN KEY (recipe_id) REFERENCES recipes(recipe_id) ON DELETE CASCADE
);

-- Create categories table for lookup
CREATE TABLE IF NOT EXISTS categories (
    category_id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_name TEXT NOT NULL UNIQUE,
    description TEXT
);

-- Insert sample categories
INSERT INTO categories (category_name, description) VALUES
    ('Pasta', 'Pasta dishes and noodle-based meals'),
    ('Dessert', 'Sweet treats and desserts'),
    ('Main Course', 'Main dishes and entrees'),
    ('Appetizer', 'Starters and appetizers'),
    ('Soup', 'Soups and stews'),
    ('Salad', 'Salads and cold dishes'),
    ('Breakfast', 'Breakfast items'),
    ('Beverage', 'Drinks and beverages');

-- Insert sample recipes
INSERT INTO recipes (recipe_name, category, prep_time_minutes, cook_time_minutes, servings) VALUES
    ('Classic Spaghetti Carbonara', 'Pasta', 10, 15, 4),
    ('Homemade Chocolate Chip Cookies', 'Dessert', 15, 12, 24),
    ('Vegetable Stir Fry', 'Main Course', 15, 10, 4);

-- Insert sample ingredients
INSERT INTO ingredients (ingredient_name, unit_of_measure) VALUES
    ('spaghetti', 'grams'),
    ('pancetta', 'grams'),
    ('eggs', 'count'),
    ('Parmesan cheese', 'grams'),
    ('black pepper', 'grams'),
    ('all-purpose flour', 'cups'),
    ('butter', 'cups'),
    ('granulated sugar', 'cups'),
    ('brown sugar', 'cups'),
    ('chocolate chips', 'cups'),
    ('vanilla extract', 'teaspoons'),
    ('baking soda', 'teaspoons'),
    ('salt', 'teaspoons'),
    ('broccoli florets', 'cups'),
    ('bell pepper', 'count'),
    ('snap peas', 'cups'),
    ('carrots', 'count'),
    ('garlic', 'cloves'),
    ('soy sauce', 'tablespoons'),
    ('sesame oil', 'tablespoons'),
    ('ginger', 'teaspoons');

-- Insert recipe ingredients for Spaghetti Carbonara (recipe_id = 1)
INSERT INTO recipe_ingredients (recipe_id, ingredient_id, quantity, unit) VALUES
    (1, 1, 400, 'grams'),  -- spaghetti
    (1, 2, 200, 'grams'),  -- pancetta
    (1, 3, 4, 'count'),    -- eggs
    (1, 4, 100, 'grams'),  -- Parmesan
    (1, 5, 1, 'to taste'); -- black pepper

-- Insert instructions for Spaghetti Carbonara
INSERT INTO instructions (recipe_id, step_number, instruction_text) VALUES
    (1, 1, 'Cook spaghetti according to package directions in salted boiling water'),
    (1, 2, 'While pasta cooks, fry pancetta in a large skillet until crispy'),
    (1, 3, 'Beat eggs with grated Parmesan cheese in a bowl'),
    (1, 4, 'Drain pasta, reserving 1 cup of pasta water'),
    (1, 5, 'Toss hot pasta with pancetta and remove from heat'),
    (1, 6, 'Quickly stir in egg mixture, adding pasta water to create a creamy sauce'),
    (1, 7, 'Season generously with black pepper and serve immediately');

-- Insert recipe ingredients for Chocolate Chip Cookies (recipe_id = 2)
INSERT INTO recipe_ingredients (recipe_id, ingredient_id, quantity, unit) VALUES
    (2, 6, 2.25, 'cups'),   -- flour
    (2, 7, 1, 'cup'),       -- butter
    (2, 8, 0.75, 'cup'),    -- granulated sugar
    (2, 9, 0.75, 'cup'),    -- brown sugar
    (2, 3, 2, 'count'),     -- eggs
    (2, 10, 2, 'cups'),     -- chocolate chips
    (2, 11, 1, 'teaspoon'), -- vanilla
    (2, 12, 1, 'teaspoon'), -- baking soda
    (2, 13, 0.5, 'teaspoon'); -- salt

-- Insert instructions for Chocolate Chip Cookies
INSERT INTO instructions (recipe_id, step_number, instruction_text) VALUES
    (2, 1, 'Preheat oven to 375°F (190°C)'),
    (2, 2, 'Cream together softened butter and both sugars until fluffy'),
    (2, 3, 'Beat in eggs and vanilla extract'),
    (2, 4, 'In a separate bowl, mix flour, baking soda, and salt'),
    (2, 5, 'Gradually mix dry ingredients into wet ingredients'),
    (2, 6, 'Stir in chocolate chips'),
    (2, 7, 'Drop rounded tablespoons of dough onto ungreased baking sheets'),
    (2, 8, 'Bake for 9-11 minutes until edges are golden brown'),
    (2, 9, 'Cool on baking sheet for 2 minutes, then transfer to wire rack');

-- Insert recipe ingredients for Vegetable Stir Fry (recipe_id = 3)
INSERT INTO recipe_ingredients (recipe_id, ingredient_id, quantity, unit) VALUES
    (3, 14, 2, 'cups'),       -- broccoli
    (3, 15, 1, 'count'),      -- bell pepper
    (3, 16, 1, 'cup'),        -- snap peas
    (3, 17, 2, 'count'),      -- carrots
    (3, 18, 3, 'cloves'),     -- garlic
    (3, 19, 2, 'tablespoons'), -- soy sauce
    (3, 20, 1, 'tablespoon'),  -- sesame oil
    (3, 21, 1, 'teaspoon');    -- ginger

-- Insert instructions for Vegetable Stir Fry
INSERT INTO instructions (recipe_id, step_number, instruction_text) VALUES
    (3, 1, 'Prepare all vegetables by washing and cutting into uniform pieces'),
    (3, 2, 'Heat sesame oil in a large wok or skillet over high heat'),
    (3, 3, 'Add minced garlic and grated ginger, stir-fry for 30 seconds until fragrant'),
    (3, 4, 'Add carrots and broccoli, stir-fry for 3 minutes'),
    (3, 5, 'Add bell pepper and snap peas, continue stir-frying for 3-4 minutes'),
    (3, 6, 'Add soy sauce and toss to coat all vegetables evenly'),
    (3, 7, 'Cook for another minute until vegetables are tender-crisp'),
    (3, 8, 'Serve hot over steamed rice or noodles');

-- Create indexes for better query performance
CREATE INDEX idx_recipes_category ON recipes(category);
CREATE INDEX idx_recipe_ingredients_recipe_id ON recipe_ingredients(recipe_id);
CREATE INDEX idx_recipe_ingredients_ingredient_id ON recipe_ingredients(ingredient_id);
CREATE INDEX idx_instructions_recipe_id ON instructions(recipe_id);

-- Create a view for easy recipe querying with full details
CREATE VIEW recipe_details AS
SELECT 
    r.recipe_id,
    r.recipe_name,
    r.category,
    r.prep_time_minutes,
    r.cook_time_minutes,
    r.servings,
    r.created_date,
    r.modified_date
FROM recipes r;
