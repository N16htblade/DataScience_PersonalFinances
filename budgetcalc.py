import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import os, sys
import subprocess
import csv
import numpy as np

inputFiles = [file for file in os.listdir('.\Budget\Input') if file.endswith('.csv')] #Check for input files (ignore non .csv), concate into main datafile and remove duplicated entries.
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


df = pd.read_csv('./Budget/Output/allYearData.csv') #Load the main data file.
df.rename(columns = {'Transaction Date':'TransactionDate', 'Description 1':'Description1', 'Description 2':'Description2'}, inplace = True) #Remove spaces in column name.
df = df.loc[~df['Description1'].str.contains('Transfer') & ~df['Description1'].str.contains('MERCI') & ~df['Description1'].str.contains('ONLINE') & ~df['Description1'].str.contains('PAYBACK')]#Remove exchanges between accounts.

with open (os.path.join(sys.path[0], "Calendar.csv")) as csv_file: #Transform transaction date into Month and sort based on Calendar.
    calendar_reader = csv.reader(csv_file, delimiter=",")
    months = {int(rows[0]):rows[1] for rows in calendar_reader}

df['TransactionDate'] = pd.to_datetime(df['TransactionDate']) 
df['TransactionDate'] = df['TransactionDate'].dt.month
df = df.sort_values('TransactionDate').reset_index()
df['Month'] = df.TransactionDate.map(months)


with open (os.path.join(sys.path[0], "Category.csv")) as csv_file:#Simplify description and map to category based on Dictionary.
    category_reader = csv.reader(csv_file, delimiter=",")
    categoryByDescription = {rows[0]:rows[1] for rows in category_reader}

df["Description"] = df.Description1.str.split().str.get(0) 
tempdf=df.loc[(df['Description'].str.len() < 3) | (df['Description'] == 'THE')] #Check for short or common denominator Description and add second string to avoid issues.
for index, row in tempdf.iterrows():
    q = row.Description1
    q = q.split()[:2]
    q = q[0] + ' ' + q[1]
    df.loc[index, 'Description'] = q


mainCategories = {'Housing' : 'Rent, Insurance, Other.', #Check for transactions without a Category mapped, verbose.
                  'Food': 'Groceries, Take-Out, Coffee, Other.',
                  'Transportation': 'Auto, AutoInsurance, Fuel, Other.',
                  'Personal': 'Gym, Hair, PersonalCare, Other.',
                  'Banking': 'Investing, Fee, Taxes, Other.',
                  'Entertainment': 'Netflix, Games, Movies, Other.',
                  'Utilities': 'Power, Internet, Phone, Other.'}
descriptionList = list(set(df["Description"]))
for i in descriptionList:
    if i in categoryByDescription.keys():
        pass
    else:
        print (f'\nFollowing transactions does not have a category assigned: {i}')
        print (f'Please enter Main and Secondary category for {i} vendor.')
        value1 = input(f'Main Category (Housing, Food, Transportation, Personal, Banking, Entertainment or Utilities):  ')
        if value1 in mainCategories.keys():
            options = mainCategories[f'{value1}']
            value2 = input(f'Sub Category (Ex: {options}): ')
        newValue = f'{value1} {value2}'
        categoryByDescription[i] = newValue

with open (os.path.join(sys.path[0], "Category.csv"), 'w', newline='') as f: #update Category dictionary
    w = csv.writer(f)
    w.writerows(categoryByDescription.items())


pat = '|'.join(r"\b{}\b".format(x) for x in categoryByDescription.keys()) #Process Descriptions and map to Categories.
df['CategoryAll'] = df['Description'].str.extract('('+ pat + ')', expand=False).map(categoryByDescription)
df['Category1'] = df.CategoryAll.str.split().str.get(0)
df['Category2'] = df.CategoryAll.str.split().str.get(1)
df['CAD$'] = df['CAD$'].astype(int)
df = df.drop(['CategoryAll', 'Account Number', 'TransactionDate', 'Account Type', 'USD$', 'Cheque Number', 'Description1', 'Description2', 'index'], axis = 1) #Remove unnecesary/empty columns.
df = df.groupby(['Month', 'Category1', 'Category2', 'Description'], sort=False).sum().reset_index() #Sort and merge database on Month, Categories and Description. Save to output file.
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


monthly = pd.DataFrame() #Following section generates data and the Monthly bar plot.
monthly['Month'] = ('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec')

def monthlyBars (category, monthly):
    monthlyx = df[df['Category1'] == f'{category}']
    monthlyx = monthlyx.groupby(['Month'], sort=False).sum().reset_index()
    cols = ['Month']
    monthly = monthly.merge(monthlyx, how='left')
    monthly.fillna(0, inplace=True)
    monthly['CAD$'] = monthly['CAD$'].abs()
    monthlyExpence = list(monthly['CAD$'])
    return monthlyExpence

housing = monthlyBars ('Housing', monthly)
utilities = monthlyBars ('Utilities', monthly)
food = monthlyBars ('Food', monthly)
transportation = monthlyBars ('Transportation', monthly)
personal = monthlyBars ('Personal', monthly)
banking = monthlyBars ('Banking', monthly)
entertainment = monthlyBars ('Entertainment', monthly)
dftemp = pd.DataFrame({'Housing' : housing,'Utilities' : utilities, 'Food' : food,'Transportation' : transportation,'Personal' : personal,'Banking' : banking,'Entertainment' : entertainment})

fig, ax = plt.subplots()
fig.set_size_inches(7,9)
#colors = ['#33bbff', '#F65314' ,'#FBBB00', '#7CBB00', '#ff99ff', '#9999ff', '#ffcc99']
dftemp.plot.barh(stacked=True, ax=ax)#, color=colors)
ax.set_title("Monthly Expenses by Group")
ax.legend(loc='lower right')
ax.set_yticklabels(monthly['Month'])
plt.savefig('./Budget/Temp/tempIEM.png',dpi=100, bbox_inches='tight')
img = openpyxl.drawing.image.Image('./Budget/Temp/tempIEM.png')
img.anchor = 'U1'
ws.add_image(img)
wb.save('./Budget/Output/finalData.xlsx')
plt.close()

## create method for main calc
##  + other method for second calc
##   -> into pie chart

yearIncome = df[df['Category1'] == 'Income']['CAD$'].sum() #Create main pie chart based on total expenses and main category.
yearExpenses = abs(df[df['Category1'] != 'Income']['CAD$'].sum())
yearExpensesFood = abs(df[df['Category1'] == 'Food']['CAD$'].sum())
yearExpensesPersonal = abs(df[df['Category1'] == 'Personal']['CAD$'].sum())
yearExpensesTransportation = abs(df[df['Category1'] == 'Transportation']['CAD$'].sum())
yearExpensesUtilities = abs(df[df['Category1'] == 'Utilities']['CAD$'].sum())
yearExpensesHousing = abs(df[df['Category1'] == 'Housing']['CAD$'].sum())
yearExpensesEntertainment = abs(df[df['Category1'] == 'Entertainment']['CAD$'].sum())
yearExpensesBanking = abs(df[df['Category1'] == 'Banking']['CAD$'].sum())

pieLabels = ('Banking', 'Transportation', 'Housing', 'Food', 'Utilities', 'Personal', 'Entertainment')
pieValues = (yearExpensesBanking, yearExpensesTransportation, yearExpensesHousing, yearExpensesFood, yearExpensesUtilities, yearExpensesPersonal, yearExpensesEntertainment)
fig, ax1 = plt.subplots()
explodeTemp = []
for i in pieLabels:
    explodeTemp.append(0.02)
explode = tuple(explodeTemp)
colors = ['#33bbff', '#F65314' ,'#FBBB00', '#7CBB00', '#9999ff', '#ff99ff', '#ffcc99']
plt.pie(pieValues, labels=pieLabels, colors=colors, autopct=lambda p : '{:.1f}%\n${:,.0f}'.format(p,p * sum(pieValues)/100), pctdistance=0.80, explode=explode, startangle=90, counterclock=False)
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


def donutChart (category, title, location, total): #Generate donutcharts based on category, name and location
    yearly = df[df['Category1'] == f'{category}']
    yearly = yearly.groupby(['Category2']).sum().reset_index()
    yearly['CAD$'] = yearly['CAD$'].abs()
    yearly = yearly.sort_values(by='CAD$', ascending=False)
    pieLabels = yearly.Category2.values.tolist()
    pieValues = yearly['CAD$'].values.tolist()

    fig, ax1 = plt.subplots()
    explodeTemp = []
    for i in pieLabels:
        explodeTemp.append(0.02)
    explode = tuple(explodeTemp)
    colors = ['#33bbff', '#F65314' ,'#FBBB00', '#7CBB00', '#9999ff', '#ff99ff', '#ffcc99']
    plt.pie(pieValues, labels=pieLabels, colors=colors, autopct=lambda p : '{:.1f}%\n${:,.0f}'.format(p,p * sum(pieValues)/100), pctdistance=0.80, explode=explode, startangle=90, counterclock=False)
    centre_circle = plt.Circle((0,0),0.60,fc='white')
    label = ax1.annotate(f'{title}\n${total}', fontsize = 12, xy=(0, -0.10), ha="center")
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    plt.savefig(f'./Budget/Temp/{category}.png',dpi=70, bbox_inches='tight')
    img = openpyxl.drawing.image.Image(f'./Budget/Temp/{category}.png')
    img.anchor = f'{location}'
    ws.add_image(img)
    wb.save('./Budget/Output/finalData.xlsx')
    plt.close()

donutChart('Housing', 'Housing', 'F12', yearExpensesHousing)
donutChart('Food', 'Food', 'K12', yearExpensesFood)
donutChart('Transportation', 'Transportation', 'P12', yearExpensesTransportation)
donutChart('Personal', 'Personal', 'A26', yearExpensesPersonal)
donutChart('Banking', 'Banking', 'F26', yearExpensesBanking)
donutChart('Entertainment', 'Entertainment', 'K26', yearExpensesEntertainment)
donutChart('Utilities', 'Utilities', 'P26', yearExpensesUtilities)

for rows in ws.iter_rows(min_col=1, max_col=30, min_row=1, max_row=39):
    for cell in rows:
        cell.fill = openpyxl.styles.PatternFill(start_color='00FFFFFF', end_color='00FFFFFF', fill_type = "solid")

ws = wb.create_sheet('Main',0)
ytdRemaining = yearIncome - yearExpenses
pieLabels = ('Income', 'Expenses')
pieValues = (yearIncome, yearExpenses)
fig, ax1 = plt.subplots()
explodeTemp = []
for i in pieLabels:
    explodeTemp.append(0.02)
explode = tuple(explodeTemp)
colors = ['#7CBB00', '#F65314']
plt.pie(pieValues, labels=pieLabels, colors=colors, autopct=lambda p : '{:.1f}%\n${:,.0f}'.format(p,p * sum(pieValues)/100), pctdistance=0.80, explode=explode, startangle=90, counterclock=False)
centre_circle = plt.Circle((0,0),0.60,fc='white')
label = ax1.annotate(f'YTD Balance', fontsize = 12, xy=(0, 0), ha="center")
fig = plt.gcf()
fig.gca().add_artist(centre_circle)
plt.savefig('./Budget/Temp/tempYTD.png', dpi = 70, bbox_inches='tight')
img = openpyxl.drawing.image.Image('./Budget/Temp/tempYTD.png')
img.anchor = 'A1'
ws.add_image(img)
wb.save('./Budget/Output/finalData.xlsx')
plt.close()

img = openpyxl.drawing.image.Image('./Budget/Temp/temp.png')
img.anchor = 'I1'
ws.add_image(img)
wb.save('./Budget/Output/finalData.xlsx')
plt.close()

mainText = openpyxl.styles.Font(bold=True, size=12)
positiveValue = openpyxl.styles.Font(bold=True, color='7CBB00', size=12)
negativeValue = openpyxl.styles.Font(bold=True, color='F65314', size=12)

for rows in ws.iter_rows(min_col=1, max_col=30, min_row=1, max_row=39):
    for cell in rows:
        cell.fill = openpyxl.styles.PatternFill(start_color='00FFFFFF', end_color='00FFFFFF', fill_type = "solid")

yearInvestements = abs(df[df['Category2'] == 'Investing']['CAD$'].sum())
ws['B16'] = 'YTD Income ='
ws['B16'].font = mainText
ws['D16'] = f'${yearIncome}'
ws['D16'].font = positiveValue
ws['B17'] = 'Available ='
ws['B17'].font = mainText
ws['D17'] = f'${ytdRemaining}'
ws['D17'].font = positiveValue
ws['B18'] = 'Invested ='
ws['B18'].font = mainText
ws['D18'] = f'${yearInvestements}'
ws['D18'].font = positiveValue

wb.save('./Budget/Output/finalData.xlsx')

finalDataLocation = os.getcwd() + '/Budget/Output/finalData.xlsx' #save final document with all data and charts

openFile = input('Would you like to open the Final Data file? (Yes / No): ')
if openFile == 'y':
    subprocess.Popen([finalDataLocation], shell=True)
else:
    pass