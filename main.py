import pandas as pd


def prepCheques(xlsx):
    df = pd.read_excel(xlsx)
    df['Customer'] = df['Customer'].fillna("")
    df['DateTime'] = df['DateTime'].dt.strftime('%m/%d/%Y')
    df['Memo'] = '[' + df['Transaction ID'] + '] - [' + df['Description'] + ']'
    dfAcc = pd.read_excel(xlsx, sheet_name='Mapping')
    df = pd.merge(df, dfAcc, on='Account', how='left')
    df = df[df['Trans'] == 'Deposits']

    print(f"Refined DataFrame:\n{df}")
    return df


def iifCheques(df, iif='src/main.iif'):
    dfEntries = df.groupby('Transaction ID')

    with open(iif, 'w', encoding="utf-8") as f:
        # Write headers
        f.write('!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\n')
        f.write('!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\tNAME\n')
        f.write('!ENDTRNS\n')

        # Processing each group
        for trnID, dfEntry in dfEntries:
            dfCrd = dfEntry[dfEntry['Debit'].isna()]
            dfDr = dfEntry[dfEntry['Debit'].notna()]
            print(f"Debit Side:\n{dfDr}")
            print(f"Credit Side:\n{dfCrd}")

            # TRNS Entry (Main Line)
            for _, mL in dfCrd.iterrows():
                f.write(
                    f"TRNS\tCHECK\t{mL['DateTime']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{trnID}\t{mL['Memo']}\t\n")

            # SPL Entry (Sub/Item Line)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"SPL\tCHECK\t{mL['DateTime']}\t{mL['ACCNT']}\t{mL['Debit']}\t{trnID}\t{mL['Memo']}\t{mL['Customer']}\t\n")

            # Transaction Break
            f.write("ENDTRNS\n")


def iifDeposits(df, iif='src/main.iif'):
    dfEntries = df.groupby('Transaction ID')

    with open(iif, 'w', encoding="utf-8") as f:
        # Write headers
        f.write('!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\n')
        f.write('!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\tNAME\n')
        f.write('!ENDTRNS\n')

        # Processing each group
        for trnID, dfEntry in dfEntries:
            dfCrd = dfEntry[dfEntry['Debit'].isna()]
            dfDr = dfEntry[dfEntry['Debit'].notna()]
            print(f"Debit Side:\n{dfDr}")
            print(f"Credit Side:\n{dfCrd}")

            # TRNS Entry (Main Line)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"TRNS\tDEPOSIT\t{mL['DateTime']}\t{mL['ACCNT']}\t{mL['Debit']}\t{trnID}\t{mL['Memo']}\t\n")

            # SPL Entry (Sub/Item Line)
            for _, mL in dfCrd.iterrows():
                f.write(
                    f"SPL\tDEPOSIT\t{mL['DateTime']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{trnID}\t{mL['Memo']}\t{mL['Customer']}\t\n")

            # Transaction Break
            f.write("ENDTRNS\n")


def main():
    # Prepare excel entries for cheques template
    df = prepCheques('src/main.xlsx')

    # Create IIF Cheques file
    iifDeposits(df)


if __name__ == '__main__':
    main()
