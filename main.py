import pandas as pd
import numpy as np
from datetime import datetime as dt


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
    dfEntries = df.groupby('Trns')

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
                    f"TRNS\tCHECK\t{mL['DATE']}\t{mL['ACCNT']}\t-{mL['Credit']}\t{trnID}\t{mL['MEMO']}\t\n")

            # SPL Entry (Sub/Item Line)
            for _, mL in dfDr.iterrows():
                f.write(
                    f"SPL\tCHECK\t{mL['DATE']}\t{mL['ACCNT']}\t{mL['Debit']}\t{trnID}\t{mL['MEMO']}\t{mL['Customer']}\t\n")

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
    dfDr = df[df['Debit'].notna()]
    dfCrd = df[df['Debit'].isna()]
    dfExp = dfDr[df['ACCNT'] != '12100 · Accounts Receivable']
    dfAR = dfDr[df['ACCNT'] == '12100 · Accounts Receivable']
    print(f"Standard Compound Entry Check:\n{dfExp}")
    print(f"Compound AR Check\n{dfAR}")

    with open(iif, 'w', encoding='utf-8') as f:
        # Write headers
        f.write('!TRNS\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\n')
        f.write('!SPL\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tDOCNUM\tMEMO\tNAME\n')
        f.write('!ENDTRNS\n')
        for _, row in dfAR.iterrows():
            AmtCrd = -row['Debit']
            crd = dfCrd.loc[dfCrd['REF'] == row['REF']].iloc[0]
            f.write(
                f"TRNS\tCHECK\t{crd['DATE']}\t{crd['ACCNT']}\t{AmtCrd}\t{crd['REF']}\t{crd['MEMO']}\t\n")
            f.write(
                f"SPL\tCHECK\t{row['DATE']}\t{row['ACCNT']}\t{row['Debit']}\t{row['REF']}\t{row['MEMO']}\t{row['Customer']}\t\n")
            f.write("ENDTRNS\n")

        for _, group in dfExp.groupby('REF'):
            AmtCrd = -group['Debit'].sum()
            crd = dfCrd.loc[dfCrd['REF'] == group['REF'].iloc[0]].iloc[0]
            f.write(
                f"TRNS\tCHECK\t{crd['DATE']}\t{crd['ACCNT']}\t{AmtCrd}\t{crd['REF']}\t{crd['MEMO']}\t\n")
            for _, row in group.iterrows():
                f.write(
                    f"SPL\tCHECK\t{row['DATE']}\t{row['ACCNT']}\t{row['Debit']}\t{row['REF']}\t{row['MEMO']}\t{row['Customer']}\t\n")
            f.write("ENDTRNS\n")


def sortData(xlsx):
    df = pd.read_excel(xlsx)
    df['Customer'] = df['Customer'].fillna('')
    df = df[df['DATE'].notna()]
    df['DATE'] = pd.to_datetime(
        df['DATE'], format='%d/%m/%Y').dt.strftime('%m/%d/%Y')

    dfFlt = pd.DataFrame()

    for key, dfGrp in df.groupby('Trns'):
        dfGrpCr = dfGrp[dfGrp['Debit'].isna()]
        dfGrpDr = dfGrp[dfGrp['Debit'].notna()]

        # Default type
        dfGrp = dfGrp.copy()
        dfGrp['Type'] = 'check'

        if not dfGrp.empty:
            crdAcc = dfGrpCr['ACCNT'].iloc[0]
            drAccs = [acc for acc in dfGrpDr['ACCNT']]
            arLen = len(dfGrp[dfGrp['ACCNT'] == '12100 · Accounts Receivable'])
            apLen = len(dfGrp[dfGrp['ACCNT'] == '23000 · Accounts Payable'])

            print(f'[{dfGrp['Trns'].iloc[0]}]')
            print(f'Credit Account: {crdAcc}')
            print(f'Debit Accounts: {drAccs}')

            if crdAcc == '23000 · Accounts Payable':
                if '12100 · Accounts Receivable' in drAccs:
                    dfGrp['Type'] = 'compBill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
                else:
                    dfGrp['Type'] = 'bill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            elif arLen > 1:
                dfGrp['Type'] = 'compCheck'
                print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            else:
                dfGrp['Type'] = 'check'
                print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            print('_'*100)

        dfFlt = pd.concat([dfFlt, dfGrp])

    dfFlt.to_excel('src/FilteredData.xlsx', sheet_name='Data', index=False)

    # print(f"Main Excel:\n{df}")
    print(f"Filtered Excel:\n{dfFlt}")


def iifWriter(sortedData):
    for key, value in sortedData.items():
        if key == 'checks':
            iifCheques(value, 'src/checks.iif')
        elif key == 'compoundChecks':
            iifCompChecks(value, 'src/compChecks.iif')


def main():
    # Prepare excel entries for cheques template
    sortData('src/main.xlsx')


if __name__ == '__main__':
    main()
