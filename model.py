#!/Users/10380r/.pyenv/shims python3
from sklearn import linear_model
import pandas as pd
from pandas import Series,DataFrame
import numpy as np
import datetime

def create_df(f):
    '''
    DataFrameを作る
    出力結果を見ると、クーポンコードが多すぎたので一旦消している。(line 25)
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
    df.drop('coupon_code', axis=1, inplace=True)
   # df.drop('user_number', axis=1, inplace=True)

    return df


def get_dummies(df):
    '''
    カテゴリ変数をダミー変数に変換
    '''
    df = pd.get_dummies(df)
    return df


def model(df):
    '''
    実際にモデルに投げる
    重回帰よりロジスティック回帰の方が精度が出たので、そちらを使用した
    重回帰にする場合、LinearRegressionにする
    '''
    clf = linear_model.LogisticRegression()
    repeat_except =  df.drop('repeat', axis = 1)
    X = repeat_except.values
    Y = df['repeat'].values
    clf.fit(X,Y)
    # 出力
    #print(pd.DataFrame({"Name":repeat_except.columns,
    #               "Coefficients":clf.coef_[0].astype('float')}).sort_values(by='Coefficients') )
    #print('score : ',clf.score(X,Y))
    
    Y_pred = clf.predict(X)
    df['pred']  = Y_pred

    return df 

    
def main():
    '''
    file変数に実際のエクセルファイルを入れる
    エクセルの注意点
        ・クーポンコードが複数ある場合の区切りを , から スペース に入れ替える
        ・カラムを増やす時は要メンテナンス
        
    結果について
        Coefficients :
        ・-になればなるほど逆相関
        ・+になればなるほど相関
        Score :
        ・データを割った際の精度
    '''
    today = str(datetime.date.today()) 
    drop_columns = ['use_coupon','discount','change','repeat','payment_クレジットカード','payment_代金引換','payment_振込用紙（コンビニ払）','payment_楽天ペイ','payment_決済不要','shipping_company_ヤマト運輸','shipping_company_佐川急便','terminal_PC','terminal_SP','terminal_TB','categoly_お手軽材料','categoly_お茶・スナック','categoly_お菓子・パン作りの型','categoly_お菓子・パン作りの道具','categoly_かんたん手作りキット','categoly_イースト・天然酵母','categoly_キッチン道具・雑貨・衛生資材','categoly_ケーキ装飾','categoly_スーパーフード・健康補助食品','categoly_チョコレート・ココア','categoly_デコレーション・トッピング','categoly_ドライフルーツ・加工野菜・果物','categoly_ナッツ(アーモンド・くるみ等)','categoly_バター・乳製品・油脂・卵','categoly_ラッピング','categoly_世界の食材','categoly_冷凍スポンジ・クッキー・パン','categoly_卸商品','categoly_和菓子材料','categoly_和食材','categoly_塩とスパイス','categoly_小麦粉・ミックス粉・雑穀粉','categoly_書籍・その他','categoly_栗・芋・かぼちゃ・シード','categoly_特集','categoly_砂糖・はちみつ・ジャム','categoly_膨張剤・香料・色素・凝固剤・添加物','categoly_菓子・パン袋','categoly_酒・リキュール類','categoly_鮮度保持材・保冷材']

    file = '購入データ（最新）.xlsx'

    df = get_dummies(create_df(file))
    results = model(df)
    results.drop(drop_columns, axis=1,inplace=True)
    results.drop('use_points', axis=1,inplace=True)
    results.to_csv('results/%s-result.txt' %(today), sep=',', index=False, header=True)


if __name__ == '__main__':
    main()
