# Meal Planning Excel Workbook
## Introduction
Welcome to the Meal Planning Excel Workbook! This tool is designed to help you simplify meal planning by allowing you to:
- Save and Organize Recipes: Store your favorite meals, complete with ingredients, macros, and serving sizes.
- Track Your Weekly Macros: Monitor your intake of key nutrients like protein, carbohydrates, and fats over the course of the week.
- Generate a Grocery List: Automatically build a grocery list based on the meals you plan for the week, ensuring you buy exactly what you need.  

Whether you're following a specific diet or just trying to eat healthier, this workbook helps you stay organized and on track with your goals. It's fully customizable and user-friendly, giving you control over your meal planning while saving you time and effort.

## Key Features
- Recipe Database: Store recipes with ingredients, cooking instructions, and macronutrient breakdowns.
- Macro Tracking: Calculate your weekly intake of protein, but can be easily expanded for calories, carbs, and fats to meet your nutritional goals.
- Meal Planner: Plan your meals for each day of the week and instantly see how they impact your macro totals.
- Grocery List Generator: Automatically creates a detailed grocery list based on your planned meals, and emails it to your phone for quick access in the store.

## How to Use
### Catalogue Sheet
The catalog sheet is where you can add new recipes to the workbook by pressing the "Add Recipe" button and entering the name of the dish you are adding. Each recipe is hyperlinked to the appropriate sheet to save time scrolling through all your recipes.  
![image](https://github.com/user-attachments/assets/762006ce-e02e-47c6-bcda-915dfd3a9425) ![image](https://github.com/user-attachments/assets/bfad4462-fde3-47c2-a612-f16d49f9aa08)
 

A new sheet will automatically be created and you can add all information pertaining to the recipe like ingredients, quantities, nutrition facts, and cooking directions. The meals it would be prepared for can be selected in the checkboxes at the top. This is to aid in organizing the selection of meals for the weekly planner.
![image](https://github.com/user-attachments/assets/04d57070-395e-46ae-a551-47562a11010c)  

After adding recipes, the "Update Menu" button can be pressed to update the food available in the "Meal Planner" sheet
![image](https://github.com/user-attachments/assets/fd3a3913-db48-4136-ac00-f14a8813eee8)

### Meal Planner
The meal planner has all the cataloged recipes saved in their respective drop-down menus to be selected for each day. It also calculates the total protein consumed since that was my main focus for nutritional macros.  
![image](https://github.com/user-attachments/assets/a11ed5c6-48d5-415d-acee-b5f8774a7112)  
By clicking the "Make Grocery List" button, the workbook retrieves all ingredients used in the recipes and emails it to the programmed email, making for easy mobile access.

![image](https://github.com/user-attachments/assets/1bcb6636-0dfd-445a-bd70-e3c71b48c5c5)  

## Setting Up Your Email
Setting up your email to work is simple. Go to the Developer tab and click the "Visual Basic" button in the ribbon.  
![image](https://github.com/user-attachments/assets/8f85d58d-9447-4211-a571-89ee35d930f4)  

Navigate to "Module1" and scroll down to the function "Send_Email". Enter the correct email on the ".To" line in replacement of "Enter.Email@Here.Com"  
![image](https://github.com/user-attachments/assets/4bc14695-9214-43b3-96bd-b2062f302ada)




