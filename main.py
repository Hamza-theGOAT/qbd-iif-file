import pandas as pd
import os
from dotenv import load_dotenv

# Load variables
load_dotenv()


def iifCompChecks(df, iif):
    dfDr = df[df['Debit'].notna()]
    dfCrd = df[df['Debit'].isna()]
    dfExp = dfDr[df['ACCNT'] != os.getenv('AR')]
    dfAR = dfDr[df['ACCNT'] == os.getenv('AR')]
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
    bankAccs = os.getenv('BANKACCS')

    for key, dfGrp in df.groupby('Trns'):
        dfGrpCr = dfGrp[dfGrp['Debit'].isna()]
        dfGrpDr = dfGrp[dfGrp['Debit'].notna()]

        # Default type
        dfGrp = dfGrp.copy()
        dfGrp['Type'] = 'check'

        if not dfGrp.empty:
            crdAcc = dfGrpCr['ACCNT'].iloc[0]
            drAccs = [acc for acc in dfGrpDr['ACCNT']]
            arLen = len(dfGrp[dfGrp['ACCNT'] == os.getenv('AR')])
            apLen = len(dfGrp[dfGrp['ACCNT'] == os.getenv('AP')])

            print(f'[{dfGrp['Trns'].iloc[0]}]')
            print(f'Credit Account: {crdAcc}')
            print(f'Debit Accounts: {drAccs}')

            if crdAcc == os.getenv('AP'):
                if os.getenv('AR') in drAccs:
                    dfGrp['Type'] = 'compBill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
                else:
                    dfGrp['Type'] = 'bill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            elif any(acc in drAccs for acc in bankAccs):
                dfGrp['Type'] = 'deposit'
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


def main():
    # Prepare excel entries for cheques template
    sortData('src/main.xlsx')


if __name__ == '__main__':
    main()
