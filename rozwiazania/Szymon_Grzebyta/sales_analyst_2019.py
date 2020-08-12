import pandas as pd
import glob
import os
import re
import calendar
import numpy as np
import datetime
import matplotlib.pyplot as plt
import openpyxl

def task1(df):
    # min, max, mean for each product
    data_frame = df.__deepcopy__(df)
    data_frame = data_frame[['Product', 'Price Each']]
    df2 = (data_frame.set_index("Product")
           .select_dtypes(np.number)
           .stack()
           .groupby(level=0)
           .agg(['min', 'max', 'mean']))

    df2 = df2.sort_values(by='max')
    print(df2)
    df2.to_excel('Task1.xlsx')


def task2(df):
    # total income by each product
    data_frame = df.__deepcopy__(df)
    data_frame = data_frame[['Product', 'Quantity Ordered', 'Price Each']]
    data_frame['Total income'] = data_frame['Quantity Ordered'] * data_frame['Price Each']
    df2 = data_frame.groupby(['Product', 'Price Each']).sum()
    df2 = df2.sort_values(by='Total income')
    print(df2)
    df2.to_excel('Task2.xlsx')


def task3(df):

    data_frame = df.__deepcopy__(df)
    data_frame = data_frame[[ 'Month', 'Month Name', 'Product','Quantity Ordered', 'Price Each']]
    data_frame['Total income'] = data_frame['Quantity Ordered'] * data_frame['Price Each']
    data_frame = data_frame.drop(['Price Each'], axis = 1)
    df2 = data_frame.groupby(['Month', 'Month Name']).sum()

    df3 = df.__deepcopy__(df)
    df3 = df3[['Month', 'Month Name', 'Product', 'Quantity Ordered', 'Price Each']]
    df3['Total income'] = df3['Quantity Ordered'] * df3['Price Each']
    df3 = df3.groupby(['Product']).sum()
    top_product = df3.sort_values(['Total income'], ascending=False)
    top_product = top_product.iloc[0]
    top_product = str(top_product.name)

    data_frame = df.__deepcopy__(df)
    data_frame = data_frame[['Month', 'Month Name', 'Product', 'Quantity Ordered', 'Price Each']]
    data_frame['Total income MacBook Pro Laptop'] = data_frame['Quantity Ordered'] * data_frame['Price Each']
    data_frame = data_frame[data_frame['Product'] == top_product]
    data_frame = data_frame.drop(['Price Each'], axis=1)
    df4 = data_frame.groupby(['Month', 'Month Name']).sum()


    result = df2.join(df4, lsuffix=' 2019_combined', rsuffix=' MacBook Pro Laptop')
    print(result)
    result1 = result[['Total income MacBook Pro Laptop', 'Total income' ]]
    result_percentage = result1.pct_change()

    result_percentage.plot()
    plt.savefig('Task3_Line_plot_by_months_2019_vs_top_product_Percentage_Change.png')
    plt.show()

    result1.plot()
    plt.savefig('Task3_Line_plot_by_months_2019_vs_top_product.png')
    plt.show()
    result1.to_excel('Task3.xlsx')

def task4(df):

    df1 = df.__deepcopy__(df)
    df1 = df1[['Month', 'Month Name', 'Product', 'Quantity Ordered', 'Price Each', 'City']]
    df1['Total income'] = df1['Quantity Ordered'] * df1['Price Each']
    df1= df1.drop(['Price Each'], axis = 1)
    df1= df1.drop(['Quantity Ordered'], axis = 1)
    df1 = df1.groupby(['Month', 'Month Name', 'City'], as_index=False ).sum()
    df1 = df1.pivot(index = 'Month', columns= 'City', values = 'Total income')
    df1.plot()
    plt.savefig('Task4_Line_plot_by_months_diffrent_cities.png')
    plt.show()
    print(df1)

def task4_bar_chart(df):

    df1 = df.__deepcopy__(df)
    df1 = df1[['Month', 'Month Name', 'Product', 'Quantity Ordered', 'Price Each', 'City']]
    df1['Total income'] = df1['Quantity Ordered'] * df1['Price Each']
    df1= df1.drop(['Price Each'], axis = 1)
    df1= df1.drop(['Quantity Ordered'], axis = 1)
    df1 = df1.groupby(['City'], as_index=False ).sum()
    df1 = df1.sort_values(by='Total income')
    df1= df1.drop(['Month'], axis = 1)
    df1.reset_index(drop=True)
    df1.plot.bar(x = 'City', y= 'Total income')
    plt.tight_layout()
    plt.savefig('Task4_Bar_chart_by_diffrent_cities.png')
    plt.show()



if __name__ == "__main__":

    # finding path to 'dane'
    path = os.path.dirname(os.path.abspath(__file__))
    path = path[0:-28] + '/dane'
    li = []
    df = pd.DataFrame

    # creating list of all month data frames
    for filename in glob.glob(os.path.join(path, '*.csv')):
        df1 = pd.read_csv(filename)
        df1 = df1.dropna()

        # extracting month from .csv file name
        month_name = re.search('Sales_(.*)_2019', filename)
        month_name = month_name.group(1)
        # converting string to datetime format
        df1['Month'] = datetime.datetime.strptime(month_name, "%B")

        indexNames = df1[df1['Product'] == 'Product'].index
        df1.drop(indexNames, inplace=True)
        li.append(df1)

    # adding together all months data frames
    df = pd.concat(li, axis=0)

    # deleting unnecessary records
    indexNames = df[df['Product'] == 'Product'].index
    df.drop(indexNames, inplace=True)
    df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)

    # changing types of variables
    df['Order ID'] = (df['Order ID'].astype(str)).astype(int)
    df['Quantity Ordered'] = (df['Quantity Ordered'].astype(str)).astype(int)
    df['Price Each'] = (df['Price Each'].astype(str)).astype(float)
    df['Order Date'] = pd.to_datetime(df['Order Date'])
    df['Month'] = pd.DatetimeIndex(df['Month']).month
    df['Month Name'] = df['Month'].apply(lambda x: calendar.month_abbr[x])
    df['City'] = df["Purchase Address"].apply(lambda x: re.search(', (.*),', x).group(1))
    # sorting by months
    df = df.sort_values(by='Month')
    df = df.reset_index(drop=True)
    df.to_csv('2019_combained.csv', index=False)

    task1(df)
    task2(df)
    task3(df)
    task4(df)
    task4_bar_chart(df)


