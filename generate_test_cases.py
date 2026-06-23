import openpyxl
import random
import os

categories = [
    "Selenium_Website_Tests",
    "Appium_Android_Tests",
    "Unit_Tests_API",
    "Validation_Tests",
    "Deployment_Status",
    "Load_Testing_Performance"
]

def generate_excel(category):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = category
    
    headers = ["Test ID", "Category", "Test Name", "Description", "Pre-conditions", "Steps", "Expected Result", "Status"]
    ws.append(headers)
    
    for i in range(1, 305):
        ws.append([
            f"TC_{category}_{i}",
            category,
            f"Verify functionality {i} for {category.replace('_', ' ')}",
            f"Detailed description for testing aspect {i} of the system.",
            "System is running, user is authenticated if needed.",
            "1. Start the test\n2. Perform action\n3. Observe result",
            "The system should behave as expected for this specific scenario.",
            random.choice(["Pass", "Fail", "Pending", "Not Executed"])
        ])
    
    filename = f"{category}.xlsx"
    wb.save(filename)
    print(f"Generated {filename}")

for cat in categories:
    generate_excel(cat)
