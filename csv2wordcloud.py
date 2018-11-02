import csv
import MeCab
from wordcloud import WordCloud

#csvファイルを指定-------------
csvFile = input('csv file name : ')
savename = input('Img savename : ')
#------------

csvFile = open(csvFile, "r", encoding="ms932", errors="", newline="" )
f = csv.reader(csvFile, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)
li  = [i[0] for i in f]

del li[0]
text = ''
for i in li:
    text += i

#mecabを使用して形態素解析
m = MeCab.Tagger()
keywords = m.parse(text)

#空のリストを作成
wordList = []

for row in keywords.split('\n'): #改行でスプリット
    word = row.split('\t')[0] #タブでスプリット
    if word == 'EOS': #EOSが出たら終了
        break
    else:
        pos = row.split('\t')[1].split(',')[0] #
        if pos == '名詞' or pos == '形容詞': #名詞と形容詞のみ抽出
            wordList.append(word) #output_wordsにappend

#対応するフォントのパスを以下に指定してください。
fpath = "/Library/Fonts/Ricty-Regular.ttf"
#------------------------------------------------

texts = ''
for i in wordList:
    texts += i +' '

wordcloud = WordCloud(background_color="white", font_path=fpath, width=900, height=500).generate(texts)

# 画像の生成
wordcloud.to_file(savename + '.png')
print('Done!')
