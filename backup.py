import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import os, sys
import subprocess
import csv


inputFiles = [file for file in os.listdir('.\Budget\Input')] #Check for input files, concate them into main datafile and remove duplicate entries.
print (f'Found {len(inputFiles)} files to process.')

allYearData = pd.DataFrame()

for file in inputFiles:
    df = pd.read_csv('./Budget/Input/' + file)
    allYearData = pd.concat([allYearData, df])

allYearData = allYearData.drop_duplicates(keep='last')
allYearData.to_csv('./Budget/Output/allYearData.csv', index=False)

count_row = allYearData.shape[0]
if count_row <1:
    print ('No data available to process, please make sure raw data is available in the Input folder.')
    exit
else:
    print ('Data loaded successfully and duplicated entries removed.\nProcessing data, please wait...')


df = pd.read_csv('./Budget/Output/allYearData.csv') #Load the main data file and processing section.
df.rename(columns = {'Transaction Date':'TransactionDate', 'Description 1':'Description1', 'Description 2':'Description2'}, inplace = True) #Remove spaces in column name for easier column management.


with open (os.path.join(sys.path[0], "Calendar.csv")) as csv_file:
    calendar_reader = csv.reader(csv_file, delimiter=",")
    months = {int(rows[0]):rows[1] for rows in calendar_reader}

df['TransactionDate'] = pd.to_datetime(df['TransactionDate']) #Transform transaction date into Month and sort based on Calendar.
df['TransactionDate'] = df['TransactionDate'].dt.month
df = df.sort_values('TransactionDate').reset_index()
df['Month'] = df.TransactionDate.map(months)


with open (os.path.join(sys.path[0], "Category.csv")) as csv_file:
    category_reader = csv.reader(csv_file, delimiter=",")
    categoryByDescription = {rows[0]:rows[1] for rows in category_reader}

df = df.loc[~df['Description1'].str.contains('Transfer')] #Remove exchanges between accounts.
df = df.loc[~df['Description1'].str.contains('MERCI')]
df["Description"] = df.Description1.str.split().str.get(0) #Simplify description and map to category based on Dictionary.
### if len of des < 4? 5?, split and get two str 0:2
pat = '|'.join(r"\b{}\b".format(x) for x in categoryByDescription.keys())
df['CategoryAll'] = df['Description'].str.extract('('+ pat + ')', expand=False).map(categoryByDescription)
df['Category1'] = df.CategoryAll.str.split().str.get(0)
df['Category2'] = df.CategoryAll.str.split().str.get(1)
df['CAD$'] = df['CAD$'].astype(int)
df = df.drop(['CategoryAll', 'Account Number', 'Account Type', 'USD$', 'Cheque Number', 'TransactionDate', 'Description1', 'Description2'], axis = 1) #Remove unnecesary/empty columns.

mainCategories = {'Housing' : 'Rent, Insurance, Other.',
                  'Food': 'Groceries, Take-Out, Coffee, Other.',
                  'Transportation': 'Auto, AutoInsurance, Fuel, Other.',
                  'Personal': 'Gym, Hair, PersonalCare, Other.',
                  'Savings': 'Investing, Other.',
                  'Entertainment': 'Netflix, Games, Movies, Other.',
                  'Utilities': 'Power, Internet, Phone, Other.'}
emptyRows = df.loc[(df['Category2'].isnull()) | (df['Category2'] == '')]  #Check for transactions without a Category mapped, verbose.
emptyRows.to_excel('./Budget/Output/finalData.xlsx', sheet_name='Empty Category', index=False)
count_row = emptyRows.shape[0]
if count_row >= 1:
    print ('Following transactions do not have a category assigned to them:')
    print (emptyRows.head(10))
    for index, row in emptyRows.iterrows():
        q = row.Description
        print (f'\nPlease enter Main and Secondary category for {q} vendor.')
        value1 = input(f'Main Category (Housing, Food, Transportation, Personal, Savings, Entertainment or Utilities):  ')
        df.loc[index, 'Category1'] = value1
        if value1 in mainCategories.keys():
            options = mainCategories[f'{value1}']
            value2 = input(f'Sub Category (Ex: {options}): ')
            df.loc[index, 'Category2'] = value2
        newValue = f'{value1} {value2}'
        categoryByDescription[q] = newValue
else:
    pass


df = df.groupby(['Month', 'Category1', 'Category2', 'Description'], sort=False).sum().reset_index() #Sort and merge database based on Month, Categories and Description. Save to output file.
df.to_excel('./Budget/Output/finalData.xlsx', sheet_name='Data', index=False)
wb = openpyxl.load_workbook('./Budget/Output/finalData.xlsx')
ws = wb.create_sheet('Graphs',0)


monthlyIncome = df[df['Category1'] == 'Income'] #Create top plot for Yearly revenue based on monthly Income/Expenses.
monthlyIncome = monthlyIncome.groupby(['Month'], sort=False).sum().reset_index()
monthlyExpenses = df[df['Category1'] != 'Income']
monthlyExpenses = monthlyExpenses.groupby(['Month'], sort=False).sum().reset_index()
monthlyExpenses['CAD$'] = monthlyExpenses['CAD$'].abs()
plt.figure(figsize=(15.6,2))
plt.title('Monthly Income / Expenses')
plt.plot(monthlyIncome['Month'], monthlyIncome['CAD$'], label='Income', color='g', marker='.', markersize=10, markeredgecolor='k', markeredgewidth=0.5)
plt.plot(monthlyExpenses['Month'], monthlyExpenses['CAD$'], label='Expenses', color='r', marker='.', markersize=10, markeredgecolor='k', markeredgewidth=0.5)
plt.legend()
plt.savefig('./Budget/Temp/tempI.png',dpi=100, bbox_inches='tight')
img = openpyxl.drawing.image.Image('./Budget/Temp/tempI.png')
img.anchor = 'A1'
ws.add_image(img)
wb.save('./Budget/Output/finalData.xlsx')
plt.close()


yearIncome = df[df['Category1'] == 'Income']['CAD$'].sum() #Create main pie chart based on total expenses and main category.
yearExpenses = abs(df[df['Category1'] != 'Income']['CAD$'].sum())
yearExpensesFood = abs(df[df['Category1'] == 'Food']['CAD$'].sum())
yearExpensesPersonal = abs(df[df['Category1'] == 'Personal']['CAD$'].sum())
yearExpensesTransportation = abs(df[df['Category1'] == 'Transportation']['CAD$'].sum())
yearExpensesUtilities = abs(df[df['Category1'] == 'Utilities']['CAD$'].sum())
yearExpensesHousing = abs(df[df['Category1'] == 'Housing']['CAD$'].sum())
yearExpensesEntertainment = abs(df[df['Category1'] == 'Entertainment']['CAD$'].sum())
yearExpensesSavings = abs(df[df['Category1'] == 'Savings']['CAD$'].sum())

pieLabels = ('Housing', 'Food', 'Personal', 'Transportation', 'Utilities', 'Savings', 'Entertainment')
pieValues = (yearExpensesHousing, yearExpensesFood, yearExpensesPersonal, yearExpensesTransportation, yearExpensesUtilities, yearExpensesSavings, yearExpensesEntertainment)
fig, ax1 = plt.subplots()
explodeTemp = []
for i in pieLabels:
    explodeTemp.append(0.05)
explode = tuple(explodeTemp)
colors = ['#ff9999', '#66b3ff', '#99ff99', '#ffcc99', '#c2c2f0', '#dec2f0', '#bfac99']
plt.pie(pieValues, labels=pieLabels, colors=colors, autopct=lambda p : '{:.1f}%\n(${:,.0f})'.format(p,p * sum(pieValues)/100), pctdistance=0.80, explode=explode, startangle=90, counterclock=False)
centre_circle = plt.Circle((0,0),0.60,fc='white')
label = ax1.annotate(f'Total Expenses\n${yearExpenses}', fontsize = 12, xy=(0, -0.10), ha="center")
fig = plt.gcf()
fig.gca().add_artist(centre_circle)
plt.savefig('./Budget/Temp/temp.png', dpi = 70, bbox_inches='tight')
img = openpyxl.drawing.image.Image('./Budget/Temp/temp.png')
img.anchor = 'A12'
ws.add_image(img)
wb.save('./Budget/Output/finalData.xlsx')
plt.close()


def pieChart (category, title, location, total): #Generate piecharts based on name and category
    monthly = df[df['Category1'] == f'{category}']
    monthly = monthly.groupby(['Category2']).sum().reset_index()
    monthly['CAD$'] = monthly['CAD$'].abs()
    monthly = monthly.sort_values(by='CAD$', ascending=False)
    pieLabels = monthly.Category2.values.tolist()
    pieValues = monthly['CAD$'].values.tolist()

    fig, ax1 = plt.subplots()
    explodeTemp = []
    for i in pieLabels:
        explodeTemp.append(0.05)
    explode = tuple(explodeTemp)
    colors = ['#ff9999','#66b3ff','#99ff99','#ffcc99', '#c2c2f0', '#ffcc99']
    plt.pie(pieValues, labels=pieLabels, colors=colors, autopct=lambda p : '{:.1f}%\n(${:,.0f})'.format(p,p * sum(pieValues)/100), pctdistance=0.82, explode=explode, startangle=90, counterclock=False)
    centre_circle = plt.Circle((0,0),0.68,fc='white')
    label = ax1.annotate(f'{title}\n${total}', fontsize = 12, xy=(0, -0.10), ha="center")
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    plt.savefig(f'./Budget/Temp/{category}.png',dpi=70, bbox_inches='tight')
    img = openpyxl.drawing.image.Image(f'./Budget/Temp/{category}.png')
    img.anchor = f'{location}'
    ws.add_image(img)
    wb.save('./Budget/Output/finalData.xlsx')
    plt.close()

pieChart('Housing', 'Housing Expenses', 'F12', yearExpensesHousing)
pieChart('Food', 'Food Expenses', 'K12', yearExpensesFood)
pieChart('Transportation', 'Auto Expenses', 'P12', yearExpensesUtilities)
pieChart('Personal', 'Personal Expenses', 'A26', yearExpensesPersonal)
pieChart('Savings', 'Savings', 'F26', yearExpensesSavings)
pieChart('Entertainment', 'Entertainment', 'K26', yearExpensesEntertainment)
pieChart('Utilities', 'Utilities', 'P26', yearExpensesUtilities)


finalDataLocation = os.getcwd() + '/Budget/Output/finalData.xlsx' #save final document with all data and charts

with open (os.path.join(sys.path[0], "Category.csv"), 'w', newline='') as f: #update Category dictionary
    w = csv.writer(f)
    w.writerows(categoryByDescription.items())
    
openFile = input('Would you like to open the Final Data file? (Yes / No): ')
if openFile == 'y':
    subprocess.Popen([finalDataLocation], shell=True)
else:
    pass