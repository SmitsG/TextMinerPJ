from Bio import Entrez
from Bio import Medline
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import operator
import re
import xlwt
from wordcloud import WordCloud

def main(searchWords,maxNumberAbstracts):
    countRelatedArticles(searchWords)
    idList = downloadPubMedIDs(searchWords, maxNumberAbstracts)
    records = parseMedlineRecords(idList)
    getRecordInformation(records)
    mostCommonWords = wordCount(dictionaryOfAllPubMedIds)
    createExcelFile(mostCommonWords, dictionaryOfAllPubMedIds)
    createWordCloudAndWriteToPng(mostCommonWords)

"""
This will query PubMed for articles having to do with the searchterm(s).
We first check how many of such articles there are.
As an extra feature, the amount of related articles can be returned to the user.
"""
def countRelatedArticles(searchWords):
    Entrez.email = "gerwinsmits@hotmail.com"     # Always tell NCBI who you are
    handle = Entrez.egquery(term=searchWords)
    record = Entrez.read(handle)
    for row in record["eGQueryResult"]:
        if row["DbName"]=="pubmed":
            print(row["Count"])

#Now we use the Bio.Entrez.efetch function to download the PubMed IDs of these articles:
def downloadPubMedIDs(searchWords, maxNumberAbstracts):
    handle = Entrez.esearch(db="pubmed", term=searchWords, retmax=maxNumberAbstracts) # retmax aanpassen voor meer resultaten
    record = Entrez.read(handle)
    idList = record["IdList"]
    return(idList)

"""
Now that we’ve got them, we want to get the corresponding Medline records and extract the information from them.
Here, we’ll download the Medline records in the Medline flat-file format, and use the Bio.Medline module to parse them:
"""
def parseMedlineRecords(idlist):
    handle = Entrez.efetch(db="pubmed", id=idlist, rettype="medline",
                           retmode="text")
    records = Medline.parse(handle)
    records = list(records)
    return(records)

"""
1. First, a dictionary is created to store all information from the records as values. The PubMedID are keys, so we can easily search for the relevant information by PubMedID.
2. Next, we use a loop to do something with the information for each article
3. Then we use a function to get information from the records, and to safe all information from the records in variables(lists and String for PubmedID). These variables will be returned. (safeRecordInformation)
4. The abstracts received from step 3. are used for textmining with NLTK. A  list of words without stopwords and interpunction will be returned. (nltkAbstractTextMining)
5. All information(in lists) for each PubMedID will be stored in a list(nested list). The nested list will be returned.
6. The function StoreInformationInDictionary will store the information.The nested list will be the value and the PubMedID the key.(StoreInformationInDictionary)
"""
#1.
dictionaryOfAllPubMedIds = {}
def getRecordInformation(records):
    #2.
    for record in records:
        #3.
        pubmedID, title, authors, source, abstracts = safeRecordInformation(record)
        #4.
        abstractWordsListWithoutStopWords = nltkAbstractTextMining(abstracts, record)
        #5.
        allInformationPubMedID = addAllInformationPubMedIDToLists(pubmedID,title,authors,source,abstracts, abstractWordsListWithoutStopWords)
        #6.
        storeInformationInDictionary(pubmedID, allInformationPubMedID, dictionaryOfAllPubMedIds)

#3.
def safeRecordInformation(record):
    pubmedID = ""
    abstracts = []
    title = []
    authors = []
    source = []
    pubmedID += record.get("PMID", "?")
    title.append(record.get("TI", "?"))
    authors.append(record.get("AU", "?"))
    source.append(record.get("SO", "?"))
    abstracts.append(record.get("AB", "?"))
    return pubmedID, title, authors, source, abstracts

#4.
def nltkAbstractTextMining(abstracts, record):
 #NLTK abstract search
    abstracts = record.get("AB", "?")
    abstractsWords = word_tokenize(abstracts)
    stopWords = set(stopwords.words("english"))
    stopWords.update(['.', ',', '"', "'", '?', '!', ':', ';', '(', ')', '[', ']', '{', '}'])
    abstractWordsListWithoutStopWords = []
    wordsNotNeeded = ["CONCLUSIONS", "CONCLUSION", "RESULTS", "BACKGROUND", "SIGNIFICANCE", "induced", "METHODS"]
    for w in abstractsWords:
        if w not in stopWords:
            if re.search("[A-Z]{2,}", w) is None:
                pass

            elif w in wordsNotNeeded:
                pass

            else:
                abstractWordsListWithoutStopWords.append(w)

    return abstractWordsListWithoutStopWords

#5.
def addAllInformationPubMedIDToLists(pubmedID, title, authors, source, abstracts, abstractWordsListWithoutStopWords):
    allInformationPubMedID = []
    allInformationPubMedID.append(title)
    allInformationPubMedID.append(authors)
    allInformationPubMedID.append(source)
    allInformationPubMedID.append(abstracts)
    allInformationPubMedID.append(abstractWordsListWithoutStopWords)
    return allInformationPubMedID

#6.
def storeInformationInDictionary(pubmedID, allInformationPubMedID, dictionaryOfAllPubMedIds):
    dictionaryOfAllPubMedIds[pubmedID] = allInformationPubMedID

"""
This function will count how many times a word exists.
After this, it will get the most occuring words.
"""
def wordCount(dictionaryOfAllPubMedIds):
    wordsDictionary = {}
    for key, value in dictionaryOfAllPubMedIds.items():
        for word in value[4]:
            if word not in wordsDictionary:
                wordsDictionary[word] = 1

            if word in wordsDictionary:
                wordsDictionary[word] += 1

    mostOccuringWords = dict(sorted(wordsDictionary.items(), key=operator.itemgetter(1), reverse=True)[:1000])

    return(mostOccuringWords)

"""
This function wil create an excel file that will contains necessary information for the user.
"""

def createExcelFile(mostOccuringWords, dictionaryOfAllPubMedIds):
    try:

        # The excel file should be refreshed each time the text mining runs.
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Most occuring words")
        ws2 = wb.add_sheet("PubMed information")

        index1 = 1
        ws.write(0,0, "Word")
        ws.write(0,1, "Value")
        for field in mostOccuringWords:
            ws.write(index1,0,field)
            index1 += 1

        index1 = 1
        for field, value in mostOccuringWords.items():
            ws.write(index1,1,value)
            index1 += 1

        index1 = 1
        ws2.write(0,0,"PubMedID")
        for key in dictionaryOfAllPubMedIds:
            ws2.write(index1,0,key)
            index1 += 1

        index1 = 1
        ws2.write(0,1,"Titel")
        for key, value in dictionaryOfAllPubMedIds.items():
            for titel in value[0]:
                ws2.write(index1,1,titel)
                index1 += 1

        index1 = 1
        ws2.write(0,2,"Source")
        for key, value in dictionaryOfAllPubMedIds.items():
            for source in value[2]:
                ws2.write(index1,2,source)
                index1 += 1

        index1 = 1
        ws2.write(0,3,"Abstracts")
        for key, value in dictionaryOfAllPubMedIds.items():
            for abstract in value[3]:
                ws2.write(index1,3,abstract)
                index1 += 1

        wb.save('C:/Users/Beheerder/Google Drive/Python Projects/School/Periode 8/Project/Excel.xls')

    except(PermissionError):
        print("File still open")

"""
This function will create a wordCloud that contains the related words.
The clouds give greater prominence to words that appear more frequently in the source text.
This function also writes the WordCloud to a png file.
"""
def createWordCloudAndWriteToPng(mostOccuringWords):
    wordcloudObject = WordCloud()
    wordcloudObject.height = 2000
    wordcloudObject.width = 2000
    wordcloudObject.max_words = 300
    wordcloudObject.generate_from_frequencies(mostOccuringWords)
    wordcloudObject.to_file("WordCloud.png")




