import pandas as pd
import os
from dotenv import load_dotenv

# Load variables
load_dotenv()


class IIFBuilder:
    def __init__(self, df: pd.DataFrame, outDir: str = 'src'):
        self.outDir = outDir
        self.df = df

    def iifChecks(self, filename: str = 'checks.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]

        with open(iif, 'w', encoding="utf-8") as f:
            # Write headers
            f.write('!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\n')
            f.write('!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')

            # TRNS Entry (Main Line)
            for _, r in dfCrd.iterrows():
                f.write(
                    f"TRNS\tCHECK\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t-{r['Credit']}\t{r['MEMO']}\t\n")

            # SPL Entry (Sub/Item Line)
            for _, r in dfDr.iterrows():
                f.write(
                    f"SPL\tCHECK\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")

            # Transaction Break
            f.write("ENDTRNS\n")

    def iifCompChecks(self, filename: str = 'compChecks.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]
        dfExp = dfDr.copy()
        dfExp = dfExp[dfExp['ACCNT'] != os.getenv('AR')]
        dfAR = dfDr.copy()
        dfAR = dfAR[dfAR['ACCNT'] == os.getenv('AR')]

        with open(iif, 'w', encoding='utf-8') as f:
            # Write headers
            f.write('!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\n')
            f.write('!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')
            for _, r in dfAR.iterrows():
                AmtCrd = -r['Debit']
                crd = dfCrd.loc[dfCrd['REF'] == r['REF']].iloc[0]
                f.write(
                    f"TRNS\tCHECK\t{crd['DATE']}\t{crd['REF']}\t{crd['ACCNT']}\t{AmtCrd}\t{crd['MEMO']}\t\n")
                f.write(
                    f"SPL\tCHECK\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write("ENDTRNS\n")

            for _, group in dfExp.groupby('REF'):
                AmtCrd = -group['Debit'].sum()
                crd = dfCrd.loc[dfCrd['REF'] == group['REF'].iloc[0]].iloc[0]
                f.write(
                    f"TRNS\tCHECK\t{crd['DATE']}\t{crd['REF']}\t{crd['ACCNT']}\t{AmtCrd}\t{crd['MEMO']}\t\n")
                for _, r in group.iterrows():
                    f.write(
                        f"SPL\tCHECK\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write("ENDTRNS\n")

    def iifDeposits(self, filename: str = 'deposits.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]

        with open(iif, 'w', encoding="utf-8") as f:
            # Write headers
            f.write('!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\n')
            f.write('!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')

            # TRNS Entry (Main Line)
            for _, r in dfCrd.iterrows():
                f.write(
                    f"TRNS\tDEPOSIT\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t\n")

            # SPL Entry (Sub/Item Line)
            for _, r in dfDr.iterrows():
                f.write(
                    f"SPL\tDEPOSIT\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t-{r['Credit']}\t{r['MEMO']}\t{r['Customer']}\t\n")

            # Transaction Break
            f.write("ENDTRNS\n")

    def iifBills(self, filename: str = 'bills.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]

        with open(iif, 'w', encoding='utf-8') as f:
            # Write headers
            f.write(
                '!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write(
                '!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')

            # Write TRNS row (Payment details)
            for _, r in dfCrd.iterrows():
                f.write(
                    f"TRNS\tBILL\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t-{r['Credit']}\t{r['MEMO']}\t{r['Customer']}\n"
                )

            # Write SPL row (AR account for customer)
            for _, r in dfDr.iterrows():
                f.write(
                    f"SPL\tBILL\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\n"
                )

            # End the transaction
            f.write('ENDTRNS\n')

    def iifCompBills(self, filename: str = 'compBills.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]
        dfExp = dfDr.copy()
        dfExp = dfExp[dfExp['ACCNT'] != os.getenv('AR')]
        dfAR = dfDr.copy()
        dfAR = dfAR[dfAR['ACCNT'] == os.getenv('AR')]

        with open(iif, 'w', encoding='utf-8') as f:
            # Write headers
            f.write('!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')

            # Invoices for the Receivable side
            for _, r in dfAR.iterrows():
                f.write(
                    f"TRNS\tINVOICE\t{r['DATE']}\t{r['REF']}\t{os.getenv('AR')}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write(
                    f"SPL\tINVOICE\t{r['DATE']}\t{r['REF']}\t{os.getenv('RA')}\t{-r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write("ENDTRNS\n")

            # Bills for the Payable side
            for _, r in dfAR.iterrows():
                f.write(
                    f"TRNS\tBILL\t{r['DATE']}\t{r['REF']}\t{os.getenv('AP')}\t{-r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write(
                    f"SPL\tBILL\t{r['DATE']}\t{r['REF']}\t{os.getenv('RA')}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write("ENDTRNS\n")

            for _, group in dfExp.groupby('REF'):
                AmtCrd = -group['Debit'].sum()
                crd = dfCrd.loc[dfCrd['REF'] == group['REF'].iloc[0]].iloc[0]
                f.write(
                    f"TRNS\tBILL\t{crd['DATE']}\t{crd['REF']}\t{crd['ACCNT']}\t{AmtCrd}\t{crd['MEMO']}\t\t{crd['Customer']}\t\n")
                for _, r in group.iterrows():
                    f.write(
                        f"SPL\tBILL\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\t\n")
                f.write("ENDTRNS\n")

    def iifJournal(self, filename: str = 'journals.iif'):
        # Output file path
        iif = os.path.join(self.outDir, filename)

        # Separate Dr and Cr portion
        dfDr = self.df[self.df['Debit'].notna()]
        dfCrd = self.df[self.df['Debit'].isna()]

        with open(iif, 'w', encoding='utf-8') as f:
            # Write headers
            f.write(
                '!TRNS\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write(
                '!SPL\tTRNSTYPE\tDATE\tDOCNUM\tACCNT\tAMOUNT\tMEMO\tNAME\n')
            f.write('!ENDTRNS\n')

            # Write TRNS row (Payment details)
            for _, r in dfCrd.iterrows():
                f.write(
                    f"TRNS\tJOURNAL\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t{r['Debit']}\t{r['MEMO']}\t{r['Customer']}\n"
                )

            # Write SPL row (AR account for customer)
            for _, r in dfDr.iterrows():
                f.write(
                    f"SPL\tJOURNAL\t{r['DATE']}\t{r['REF']}\t{r['ACCNT']}\t-{r['Credit']}\t{r['MEMO']}\t{r['Customer']}\n"
                )

            # End the transaction
            f.write('ENDTRNS\n')

    def writeIIF(self):
        for ty, dfTy in self.df.groupby('Type'):
            # print(f"Processing: {ty}")
            # print(f"DataFrame:\n{dfTy}")
            if ty == 'bill':
                self.iifBills()
            elif ty == 'compBill':
                self.iifCompBills()
            elif ty == 'deposit':
                self.iifDeposits()
            elif ty == 'journal':
                self.iifJournal()
            elif ty == 'check':
                self.iifChecks()
            elif ty == 'compCheck':
                self.iifCompChecks()
            else:
                KeyError('Invalid Type')


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
            arAcc = os.getenv('AR')
            apAcc = os.getenv('AP')
            arLen = len(dfGrp[dfGrp['ACCNT'] == arAcc])
            apLen = len(dfGrp[dfGrp['ACCNT'] == apAcc])
            dfGrp['WARNINGS'] = ""

            print(f'[{dfGrp['Trns'].iloc[0]}]')
            print(f'Credit Account: {crdAcc}')
            print(f'Debit Accounts: {drAccs}')

            if crdAcc == apAcc:
                if arAcc in drAccs:
                    dfGrp['Type'] = 'compBill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
                else:
                    dfGrp['Type'] = 'bill'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            elif crdAcc == arAcc:
                if arAcc in drAccs:
                    dfGrp['Type'] = 'compDeposit'
                    dfGrp['WARNINGS'] = '[ATTENTION] - [One off Event]'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
                else:
                    dfGrp['Type'] = 'deposit'
                    print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            elif crdAcc not in bankAccs:
                dfGrp['Type'] = 'journal'
                print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            elif arLen > 1:
                dfGrp['Type'] = 'compCheck'
                print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            else:
                dfGrp['Type'] = 'check'
                print(f'Setting Type to: {dfGrp['Type'].iloc[0]}')
            print('_'*100)

        dfFlt = pd.concat([dfFlt, dfGrp])

    dfFlt = dfFlt[os.getenv('DFFLTHDR').split(',')]
    dfFlt.to_excel('src/FilteredData.xlsx', sheet_name='Data', index=False)

    # print(f"Main Excel:\n{df}")
    print(f"Filtered Excel:\n{dfFlt}")

    return dfFlt


def main():
    # Prepare excel entries for cheques template
    df = sortData('src/main.xlsx')
    builder = IIFBuilder(df, 'src/iifExportFiles')
    builder.writeIIF()


if __name__ == '__main__':
    main()
