# -*- coding: utf-8 -*-
"""
Created on Fri Dec 21 11:31:34 2018

@author: User
"""

import os
import datetime
from datetime import date
from datetime import datetime as dt
import win32com.client as win32com
from bs4 import BeautifulSoup
import re
import zipfile
import urllib
import glob
import sys
import datetime
import numpy as np
import pandas as pd
import sys 
from datetime import timedelta
import re
from nltk.corpus import stopwords

from sklearn.feature_extraction.text import CountVectorizer, TfidfTransformer
from sklearn.naive_bayes import MultinomialNB
from sklearn.svm import SVC, LinearSVC
from sklearn.metrics import classification_report, f1_score, accuracy_score, confusion_matrix
from sklearn.pipeline import Pipeline
from sklearn.grid_search import GridSearchCV
from sklearn.cross_validation import StratifiedKFold, cross_val_score, train_test_split 
from sklearn.tree import DecisionTreeClassifier 
from sklearn.learning_curve import learning_curve


win32com.dynamic.Dispatch('Outlook.Application')
outlook = win32com.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace('MAPI')
email = 'munkhnaran.gankhuyag84@spsmail.cuny.edu'

#create a function to find folder

def findFolder(folderName,searchIn):
    print(folderName)
    print(searchIn)
    result = None
    try:
        lowerAccount = searchIn.Folders
        firstItem = lowerAccount.GetFirst()
        if firstItem.Name.lower() == folderName.lower():
            return firstItem
        else:
            while True:
                NextItem = lowerAccount.GetNext()
                if NextItem is None:
                    print("Trouble accessing subfolder")
                    break
                else:
                    print(NextItem.Name)
                    if NextItem.Name.lower() == folderName.lower():
                        result = NextItem
                        break
                    else:
                        continue
    except:
        print("Looks like we had an issue accessing the searchIn object")
    return result


#Access the main folder
main_folder = findFolder(email, mapi)


inbox = findFolder('Inbox', main_folder)


junk = findFolder('Junk Email',main_folder)


messages_inbox = inbox.Items
messages_inbox.Sort('CreationTime', 1)

messages_inbox.GetFirst().Subject


messages_spam = junk.Items
messages_spam.Sort('CreationTime', 1)

messages_spam.GetFirst().Subject


x = messages_spam.GetFirst().Body

spam = []

spam.append(x)

for i in range(messages_spam.Count-1):
    spam.append(messages_spam.GetNext().Body)

spam = [item.replace('\r','').replace('\n','').replace('/','') for item in spam]

spam = pd.DataFrame(spam, columns=['email'])

spam['type'] = 'spam'
spam = spam[['type','email']]

x = messages_inbox.GetFirst().Body

inbox = []

inbox.append(x)

for i in range(messages_inbox.Count-1):
    inbox.append(messages_inbox.GetNext().Body)
    
    
import random
random.shuffle(inbox)


inbox2 = inbox[:15]


inbox2 = [item.replace('\r','').replace('\n','').replace('/','') for item in inbox2]


inbox2 = pd.DataFrame(inbox2, columns =['email'])


inbox2['type']= 'inbox'
inbox2 = inbox2[['type','email']]


all_email = pd.concat([inbox2,spam]).reset_index(drop=True)


all_email.groupby('type').describe()

all_email['length'] = all_email['email'].map(lambda text: len(text))
all_email.head()


all_email.hist(column='length', by='type', bins=50)

from textblob import TextBlob
def split_into_tokens(message):
  # convert bytes into proper unicode
    return TextBlob(message).words

all_email.email.head().apply(split_into_tokens)

def split_into_lemmas(message):
    words = TextBlob(message).words
    words = [x for x in words if x not in stopwords.words('english')]
    words = [x for x in words if x.isalpha()]
    # for each word, take its "base form" = lemma 
    return [word.lemma for word in words]

all_email.email.head().apply(split_into_lemmas)


bow_transformer = CountVectorizer(analyzer=split_into_lemmas).fit(all_email['email'])
len(bow_transformer.vocabulary_)


message4 = all_email['email'][9]
print(message4)


bow4 = bow_transformer.transform([message4])

bow4.shape

print(bow_transformer.get_feature_names()[72])


messages_bow = bow_transformer.transform(all_email['email'])
print('sparse matrix shape:', messages_bow.shape)
print('number of non-zeros:', messages_bow.nnz)
print('sparsity: %.2f%%' % (100.0 * messages_bow.nnz / (messages_bow.shape[0] * messages_bow.shape[1])))

tfidf_transformer = TfidfTransformer().fit(messages_bow)
tfidf4 = tfidf_transformer.transform(bow4)
print(tfidf4)

print(tfidf_transformer.idf_[bow_transformer.vocabulary_['u']])
print(tfidf_transformer.idf_[bow_transformer.vocabulary_['school']])


messages_tfidf = tfidf_transformer.transform(messages_bow)
print(messages_tfidf.shape)

spam_detector = MultinomialNB().fit(messages_tfidf, all_email['type'])

messages_tfidf = tfidf_transformer.transform(messages_bow)
print(messages_tfidf.shape)


print('predicted: ' , spam_detector.predict(tfidf4)[0])
print('expected: ', all_email.type[9])

all_predictions = spam_detector.predict(messages_tfidf)
print(all_predictions)



print('accuracy', accuracy_score(all_email['type'], all_predictions))
print('confusion matrix\n', confusion_matrix(all_email['type'], all_predictions))
print('(row=expected, col=predicted)')


import matplotlib.pyplot as plt

plt.matshow(confusion_matrix(all_email['type'], all_predictions), cmap=plt.cm.binary, interpolation='nearest')
plt.title('confusion matrix')
plt.colorbar()
plt.ylabel('expected label')
plt.xlabel('predicted label')


print(classification_report(all_email['type'], all_predictions))


msg_train, msg_test, label_train, label_test = \
    train_test_split(all_email['email'], all_email['type'], test_size=0.2)

print(len(msg_train), len(msg_test), len(msg_train) + len(msg_test))

pipeline = Pipeline([
    ('bow', CountVectorizer(analyzer=split_into_lemmas)),  # strings to token integer counts
    ('tfidf', TfidfTransformer()),  # integer counts to weighted TF-IDF scores
    ('classifier', MultinomialNB()),  # train on TF-IDF vectors w/ Naive Bayes classifier
])
