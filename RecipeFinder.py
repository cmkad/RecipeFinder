import openpyxl as xl

file_path = "c:/Users/miss_/OneDrive/Desktop/RecipeFinder/recipes.xlsx"
workbook = xl.load_workbook(filename=file_path)

# Get the active worksheet
ws = workbook.active

printed_recipes = set()  # Initialize an empty set to store printed recipe titles

while True:
    # The String(s) we'll search for
    search_strings = input('Enter Ingredient(s) (comma-separated) or "exit" to quit: ')
    
    # Check if the user wants to exit
    if search_strings.lower() == 'exit':
        break
    
    search_strings = search_strings.split(",")
    
    # Initialize a set for each query to store matching recipe titles
    matching_recipes = set()

    # Iterate through rows starting from the second row (assuming data starts from row 2)
    for row in ws.iter_rows(min_row=2):
        cell_value = str(row[3].value)  # Column D, containing ingredients
        
        # Check if any search string is found in the cell's value
        if any(search_string.lower() in cell_value.lower() for search_string in search_strings):
            matching_recipes.add(row[0].value)  # Column A, containing recipe names

    # Calculate the set difference to find new recipes
    new_recipes = matching_recipes - printed_recipes

    # Print the matching recipe names that haven't been printed before
    if new_recipes:
        print("Matching Recipes:")
        for recipe_name in new_recipes:
            print(recipe_name)
        printed_recipes.update(new_recipes)  # Update the set of printed recipe titles
    else:
        print("No new matching recipes found.")
