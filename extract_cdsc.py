import pandas as pd

def load_symbols():
    df = pd.read_excel('./data/CDSC.xlsx')
    symbol = [x['Symbol'] for _, x in df.iterrows() if (x['ISIN Type'] not in ['Promoter Share','Mutual Fund'])]
    return symbol
    

if __name__ == "__main__":
    load_symbols()
