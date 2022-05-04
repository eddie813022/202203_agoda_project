import konlpy
import sqlite3
import nltk

con = sqlite3.connect('text.db')
cur = con.cursor()

# POS tag a sentence
sentence = u"저는 대만 사람이에요"
words = konlpy.tag.Okt().pos(sentence)

# Define a chunk grammar, or chunking rules, then chunk
grammar = """
NP: {<N.*>*<Suffix>?}   # Noun phrase
VP: {<V.*>*}            # Verb phrase
AP: {<A.*>*}            # Adjective phrase
"""
parser = nltk.RegexpParser(grammar)
chunks = parser.parse(words)

ap = []
for row in chunks:
    if "NP" in str(row):
        strs = str(row)
        strs = strs.strip("()")
        strs = strs.replace("NP ","")
        strs = strs.replace("/Noun","")
        strs = strs.replace(" ",",")
        if "," in strs:
            strs_ = strs.split(",")
            for i in range(len(strs_)):
                ap.append(strs_[0])
                strs_.pop(0)
        else:
            ap.append(strs)
        # cur.execute("INSERT INTO korea_np (id,word) VALUES (?,?)",strs)
        # con.commit()
    elif "AP" in strs:
        strs = str(row)
        # cur.execute("INSERT INTO korea_ap (id,word) VALUES (?,?)",strs)
        # con.commit()

con.close()
print(ap)

