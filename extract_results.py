# Tokyo Gas PoC - Scenario A3
# Code to extract results from csv output files from A1

import pandas as pd
import datetime


def read_values(fn_sales, fn_purchase, fn_vol):
    # empty dataframe for JKM values
    df_JKM = pd.DataFrame()

    # read and store JKM sales at monthly intervals
    df_sales1 = pd.read_csv(fn_sales)
    # print(df_sales1["Value"])
    df_JKM["Year"] = df_sales1["Year"]
    df_JKM["Month"] = df_sales1["Month"]
    df_JKM["Sales"] = df_sales1["Value"]
    df_JKM["Day"] = '1'

    # read and store JKM Purchases at monthly intervals
    df_purchases1 = pd.read_csv(fn_purchase)
    # print(df_purchases1["Value"])
    df_JKM["Purchases"] = df_purchases1["Value"]
    df_sales1["Date"] = pd.to_datetime(dict(year=df_sales1.Year, month=df_sales1.Month, day=df_sales1.Day)) # no warnings, format='%d/%m/%Y'
    df_JKM["Date"] = df_sales1["Date"].dt.strftime('%d/%m/%Y')
    # df_JKM["Date"] = pd.to_datetime(df_sales1[["Year", "Month", "Day"]]) # warnings

    # test formed dataframe for JKM
    # print(df_JKM)

    # read and store the Gas storage tank volumes at daily intervals
    df_endvol1 = pd.read_csv(fn_vol)

    # select "TNK-3M-1" values
    mask = df_endvol1["Name"].isin(["TNK-3M-1"])
    df_3M1 = df_endvol1[mask]
    df_3M1["Date"] = pd.to_datetime(dict(year=df_3M1.Year, month=df_3M1.Month, day=df_3M1.Day))
    df_3M1["Date"] = df_3M1["Date"].dt.strftime('%d/%m/%Y')
    # df_3M1["Date"] = pd.to_datetime(df_3M1[["Year", "Month", "Day"]])
    # print(df_3M1)


   #select "TNK-7M-1" values
    mask = df_endvol1["Name"].isin(["TNK-7M-1"])
    df_7M1 = df_endvol1[mask]
    df_7M1["Date"] = pd.to_datetime(dict(year=df_7M1.Year, month=df_7M1.Month, day=df_7M1.Day))
    df_7M1["Date"] = df_7M1["Date"].dt.strftime('%d/%m/%Y')
    # df_7M1["Date"] = pd.to_datetime(df_7M1[["Year", "Month", "Day"]])
    # print(df_7M1)

    return df_JKM, df_3M1, df_7M1

def main():
    location = './Model Base (Scenario A1) Solution/'
    df_JKM, df_3M1, df_7M1 = read_values(location+'Sales.csv', location+'Purchases.csv', location+'End Volume.csv')
    print(df_JKM)

    # set up





if __name__ == '__main__':
    main()