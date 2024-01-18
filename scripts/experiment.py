import pandas as pd

def read_excel_file(dir):
    df = pd.read_excel(dir)

    # Select only the necessary columns
    df = df[['Klant','Klant nummer', 'Store', 'Serienummer']]

    # Remove NaN values
    df = df.dropna()

    # Only keep the numeric values in de dataframe
    df = df.astype({'Klant nummer':'string'})
    df['Klant nummer'] = df['Klant nummer'].str.extract('(\d+)')

    # Remove duplicate klant nummers
    df = df.drop_duplicates(subset=['Serienummer'])

    # Reset the index
    df = df.reset_index(drop=True)

    # Make the DataFrame smaller by combining the serial numbers together in an array
    df_working_data = pd.DataFrame(columns=['Customer', 'Customer Number', 'Serial Numbers', 'Store'])
    for index, row in df.iterrows():
        cust_number = row['Klant nummer']
        serial_numbers = df['Serienummer'].loc[df['Klant nummer'] == cust_number].tolist()
        df_working_data.loc[len(df_working_data)] = [row['Klant'], cust_number, serial_numbers, row['Store']]

    # Remove duplicate customer numbers
    df_working_data = df_working_data.drop_duplicates(subset=['Customer Number'])

    return df_working_data

df = read_excel_file(r'C:\Users\Kassa\Documents\Scripts\ReturnRental\excel\TestReturnRental_Case1.xlsx')

df_succes = pd.DataFrame(columns=['Customer','Customer Number', 'Serial Number', 'Interaction', 'Store'])
df_failed = pd.DataFrame(columns=df_succes.columns.append(pd.Index(['Reason'])))

print(df)

for index, row in df.iterrows():
    serie = pd.concat([row, pd.Series(['Testing'], index=['Reason'])])
    df_failed.loc[len(df_failed)] = serie

print(df_failed)