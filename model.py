from sklearn import linear_model
import pandas as pd
from pandas import Series,DataFrame
import codecs

def create_df(csv): 
    with codecs.open(csv, 'r', 'utf-8', 'ignore') as f:
        df = pd.read_csv(f, delimiter = ',')
    df.columns = [
             'payment',
             'shipping_company',
             'use_points',
             'coupon_code',
             'use_coupon',
             'discount',
             'change',
             'terminal',
             'categoly',
             'repeat'
             ]
    return df

def l_regression(df):
    clf = linear_model.LinearRegression()
    repeat_except_quality =  df.drop('repeat')
    X = repeat_except_quality.as_matrix()
    Y = df['repeat'].as_matrix()

    clf.fit(X,Y)
    print(pd.DataFrame({"Name":repeat_except_quality.columns,
                    "Coefficients":clf.coef_}).sort_values(by='Coefficients') )

def main():
    df = create_df('datasets_tmp.csv')
    l_regression(df)

if __name__ == '__main__':
    main()
