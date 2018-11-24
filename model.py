from sklearn import linear_model
import pandas as pd
from pandas import Series,DataFrame
import numpy as np
import sys

def create_df(f): 
    '''
    データフレーム作る
    '''
    df = pd.read_excel(f)
    df.columns = [
            'user_number',
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
    # 必要ないカラムを削除
    df.drop('coupon_code', axis=1, inplace=True)
    df.drop('user_number', axis=1, inplace=True)
    return df

def get_dummies(df):
    '''
    カテゴリ変数をダミー変数で書き出す
    '''
    df = pd.get_dummies(df)
    return df

def model(df):
    '''
    ロジスティック回帰で解析させる
    '''
    clf = linear_model.LogisticRegression()
    repeat_except =  df.drop('repeat', axis = 1)
    X = repeat_except.values
    Y = df['repeat'].values
    clf.fit(X,Y)
    print(pd.DataFrame({"Name":repeat_except.columns,
                   "Coefficients":clf.coef_[0].astype('float')}).sort_values(by='Coefficients') )
    print('score : ',clf.score(X,Y))

def main():
    '''
    結果
        -1 <= coef <= 1
        逆相関 < 相関
    '''

    df = get_dummies(create_df(sys.argv[1]))
    model(df)

if __name__ == '__main__':
    main()
