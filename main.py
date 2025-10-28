import pandas as pd


def prepCheques(xlsx, ty):
    df = pd.read_excel(xlsx)
    df['Customer'] = df['Customer'].fillna("")
    df['DateTime'] = df['DateTime'].dt.strftime('%m/%d/%Y')
    # df['Transaction ID'] = df['Transaction ID'].astype(
    #     str).str.replace('.0', '')
    # df['Memo'] = '[' + df['Transaction ID'] + '] - [' + df['Description'] + ']'
    # dfAcc = pd.read_excel(xlsx, sheet_name='Mapping')
    # df = pd.merge(df, dfAcc, on='Account', how='left')
    df = df[df['Trans'] == ty]

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
            # print(f"Debit Side:\n{dfDr}")
            # print(f"Credit Side:\n{dfCrd}")

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


def iifInvoice(df, iif='src/main.iif'):
    dfEntries = df.groupby('Transaction ID')
    # Open IIF file for writing
    with open(iif, 'w', encoding='utf-8') as f:
        # Write headers
        f.write(
            '!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tMEMO\n')
        f.write(
            '!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tMEMO\n')
        f.write('!ENDTRNS\n')

        # Processing each group
        for trnID, dfEntry in dfEntries:
            dfCrd = dfEntry[dfEntry['Debit'].isna()]
            dfDr = dfEntry[dfEntry['Debit'].notna()]
            # print(f"Debit Side:\n{dfDr}")
            # print(f"Credit Side:\n{dfCrd}")

            # Write TRNS row (Payment details)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"TRNS\tINVOICE\t{mL['DateTime']}\t{mL['ACCNT']}\t{mL['Debit']}\t{
                        mL['Customer']}\t{trnID}\t{mL['Memo']}\n"
                )

            # Write SPL row (AR account for customer)
            for _, mL in dfCrd.iterrows():
                f.write(
                    f"SPL\tINVOICE\t{mL['DateTime']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{
                        mL['Customer']}\t{trnID}\t{mL['Memo']}\n"
                )
            # End the transaction
            f.write('ENDTRNS\n')


def iifCustPmt(df, iif='src/main.iif'):
    dfEntries = df.groupby('Transaction ID')
    # Open IIF file for writing
    with open(iif, 'w', encoding='utf-8') as f:
        # Write headers
        f.write(
            '!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tPAYMETH\tMEMO\n')
        f.write(
            '!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tPAYMETH\tMEMO\n')
        f.write('!ENDTRNS\n')

        # Processing each group
        for trnID, dfEntry in dfEntries:
            dfCrd = dfEntry[dfEntry['Debit'].isna()]
            dfDr = dfEntry[dfEntry['Debit'].notna()]
            # print(f"Debit Side:\n{dfDr}")
            # print(f"Credit Side:\n{dfCrd}")

            # Write TRNS row (Payment details)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"TRNS\tPAYMENT\t{mL['DateTime']}\t{mL['ACCNT']}\t{mL['Debit']}\t{
                        mL['Customer']}\t{trnID}\te-CHECK\t{mL['Memo']}\n"
                )

            # Write SPL row (AR account for customer)
            for _, mL in dfCrd.iterrows():
                f.write(
                    f"SPL\tPAYMENT\t{mL['DateTime']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{
                        mL['Customer']}\t{trnID}\te-CHECK\t{mL['Memo']}\n"
                )
            # End the transaction
            f.write('ENDTRNS\n')


def iifBill(df, iif='src/main.iif'):
    dfEntries = df.groupby('Transaction ID')
    # Open IIF file for writing
    with open(iif, 'w', encoding='utf-8') as f:
        # Write headers
        f.write(
            '!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tMEMO\n')
        f.write(
            '!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tNAME\tDOCNUM\tMEMO\n')
        f.write('!ENDTRNS\n')

        # Processing each group
        for trnID, dfEntry in dfEntries:
            dfCrd = dfEntry[dfEntry['Debit'].isna()]
            dfDr = dfEntry[dfEntry['Debit'].notna()]
            # print(f"Debit Side:\n{dfDr}")
            # print(f"Credit Side:\n{dfCrd}")

            # Write TRNS row (Payment details)
            for _, mL in dfCrd.iterrows():
                f.write(
                    f"TRNS\tBILL\t{mL['DateTime']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{
                        mL['Customer']}\t{trnID}\t{mL['Memo']}\n"
                )

            # Write SPL row (AR account for customer)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"SPL\tBILL\t{mL['DateTime']}\t{mL['ACCNT']}\t{mL['Debit']}\t{
                        mL['Customer']}\t{trnID}\t{mL['Memo']}\n"
                )
            # End the transaction
            f.write('ENDTRNS\n')

# [Filter GS for multi-Entries]


def iifCompChecks(df, iif):
    pass


def sortData(xlsx):
    df = pd.read_excel(xlsx)
    checks = df[df['Type'] == 'Checks']
    compoundChecks = df[df['Type'] == 'CompoundChecks']

    return {
        "checks": checks,
        "compoundChecks": compoundChecks
    }


def iifWriter(sortedData):
    for key, value in sortData.items():
        if key == 'checks':
            iifCheques(value, 'src/checks.iif')
        elif key == 'compoundChecks':
            iifCompChecks(value, 'src/compChecks.iif')


def main():
    # Prepare excel entries for cheques template
    dp = sortData('src/main.xlsx')

    # Create IIF Cheques file
    iifWriter(dp)


if __name__ == '__main__':
    main()
