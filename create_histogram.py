#-*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import codecs
import sys

def create_df(csv):
    with codecs.open(csv, "r", "utf-8", "ignore") as f:
        df = pd.read_csv(f, delimiter=",")
    global columns
    columns = [
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
    df.columns = columns
    return df

def grouping(df):
    grouped = df.groupby('repeat')
    return grouped

def frequency(df):
    return print(df['repeat'].value_counts())

def create_histgram(df):
    count = 1
    for column in columns:
        x = df[column].value_counts()
        y = df['repeat'].value_counts()
        plt.hist([x,y],label = ['repeat', column])
        #plt.title(column)
        plt.savefig('results/'+ column + '_' + str(count) + '.png')
        count += 1

def main():
    df = create_df('datasets_tmp.csv')
    create_histgram(df)

if __name__ == '__main__':
    main()
