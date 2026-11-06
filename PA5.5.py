import csv
import openpyxl
from openpyxl.chart import PieChart, Reference
import matplotlib.pyplot as plt

def askUser():
    total = 0
    # This loop iterates 5 times to prompt the user for five numerical inputs, 
    # adds each number to the running total accumulator, and updates the sum.
    for i in range(5):
        num = float(input(f"Enter number {i+1}: "))
        total += num
    print(f"The total of the five numbers is: {total}")


def askIncome():
    entries = []
    # This loop iterates 5 times to collect the name and annual income 
    # for five separate people, storing them temporarily in a list before appending.
    for i in range(5):
        name = input(f"Enter name of person {i+1}: ")
        income = int(input(f"Enter annual income for {name}: "))
        entries.append([name, income])
    
    # Open final.csv in append mode and write each new entry with a newline character,
    # adding 5 new rows to the existing 4 rows for a total of 9 entries.
    with open('final.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        for entry in entries:
            writer.writerow(entry)


def excelPie():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Incomes"

    # Load existing data from final.csv into the spreadsheet and ensure incomes are cast to int
    ws.append(["Name", "Annual Income"])
    with open('final.csv', mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            if row:
                ws.append([row[0], int(row[1])])

    # Initialize a new PieChart object
    pie = PieChart()

    # Set the title of the pie chart using StudentID and today's date
    pie.title = "jorkee0238 November 5, 2026"

    # Define the data reference range including headers for the pie slices
    data = Reference(ws, min_col=2, min_row=1, max_row=10)

    # Define the categories reference range for the slice labels
    labels = Reference(ws, min_col=1, min_row=2, max_row=10)

    # Add the data series to the pie chart object with titles included
    pie.add_data(data, titles_from_data=True)

    # Set the category labels corresponding to each data slice
    pie.set_categories(labels)

    # Place the completed pie chart object onto the worksheet at cell D1
    ws.add_chart(pie, "D1")

    # Save the populated workbook and chart into a new Excel file named final.xlsx
    wb.save('final.xlsx')


def verticalBar():
    names = []
    incomes = []
    
    # Read the data from final.csv to populate lists for the bar chart
    with open('final.csv', mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            if row:
                names.append(row[0])
                incomes.append(int(row[1]))

    # Generate a vertical bar graph using matplotlib
    plt.figure(figsize=(8, 6))
    plt.bar(names, incomes, color='teal')
    plt.title("jorkee0238 November 5, 2026")
    plt.xlabel("Name")
    plt.ylabel("Annual Income")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()
    
def main():
    print("--- Running askUser() ---")
    askUser()
    
    print("\n--- Running askIncome() ---")
    askIncome()
    
    print("\n--- Running excelPie() ---")
    excelPie()
    print("Excel pie chart created and saved as 'final.xlsx'.")
    
    print("\n--- Running verticalBar() ---")
    verticalBar()


if __name__ == "__main__":
    main()
