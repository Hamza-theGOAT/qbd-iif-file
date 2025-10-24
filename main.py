import pandas as pd


def prepCheques(xlsx):
    df = pd.read_excel(xlsx)
    df['Customer'] = df['Customer'].fillna("")
    df['DateTime'] = df['DateTime'].dt.strftime('%m/%d/%Y')
    df['Memo'] = '[' + df['Transaction ID'] + '] - [' + df['Description'] + ']'
    dfAcc = pd.read_excel(xlsx, sheet_name='Mapping')
    df = pd.merge(df, dfAcc, on='Account', how='left')
    df = df[df['Trans'] == 'Cheques']

    print(f"Refined DataFrame:\n{df}")
    return df


def main():
    # Prepare excel entries for cheques template
    df = prepCheques('src/main.xlsx')


if __name__ == '__main__':
    main()
