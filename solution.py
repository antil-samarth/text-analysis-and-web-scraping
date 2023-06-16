import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import nltk
from nltk.tokenize import word_tokenize
import re

df = pd.read_excel('input.xlsx')
check = []

#analyze the text data and update the excel sheet
def analyze(url_id, worksheet, text, url):
    
    stop_words = set()
    with open('StopWords/StopWords_Auditor.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_Currencies.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_DatesandNumbers.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_Generic.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_GenericLong.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_Geographic.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    with open('StopWords/StopWords_Names.txt', 'r') as file:
        stop_words.update(file.read().splitlines())

    text.strip()
    tokens = word_tokenize(text)

    filtered_tokens = [word for word in tokens if not word.lower() in stop_words]

    with open('MasterDictionary/positive-words.txt', 'r') as file:
        positive_words = set(file.read().splitlines())

    with open('MasterDictionary/negative-words.txt', 'r') as file:
        negative_words = set(file.read().splitlines())
    
    positive_score = sum([1 for word in filtered_tokens if word.lower() in positive_words])
    negative_score = sum([1 for word in filtered_tokens if word.lower() in negative_words])
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / ((len(filtered_tokens)) + 0.000001)
    
    sentences = nltk.sent_tokenize(text)
    num_sentences = len(sentences)
    num_words = len(tokens)
    num_complex_words = len([word for word in tokens if len(word) > 2 and word.isalpha() and word.lower() not in stop_words])
    num_syllables = sum([len([c for c in word if c.lower() in 'aeiou']) for word in tokens])
    num_personal_pronouns = len(re.findall(r'\b(I|we|my|ours|(?!US)us)\b', text, re.IGNORECASE))
    avg_sentence_length = num_words / num_sentences
    pct_complex_words = num_complex_words / num_words
    fog_index = 0.4 * (avg_sentence_length + pct_complex_words)
    avg_words_per_sentence = num_words / num_sentences
    avg_word_length = sum([len(word) for word in tokens]) / num_words
    num_syllables_per_word = num_syllables / num_words

    """ print(f'Positive Score: {positive_score}')
    print(f'Negative Score: {negative_score}')
    print(f'Polarity Score: {polarity_score}')  
    print(f'Subjectivity Score: {subjectivity_score}')
    print(f'Average Sentence Length: {avg_sentence_length}')
    print(f'Percentage of Complex Words: {pct_complex_words}')
    print(f'Fog Index: {fog_index}')
    print(f'Complex Word Count: {num_complex_words}')
    print(f'Word Count: {num_words}')
    print(f'Syllables Per Word: {num_syllables_per_word}')
    print(f'Personal Pronouns: {num_personal_pronouns}')
    print(f'Average Word Length: {avg_word_length}') """
    
    
    for row in worksheet.iter_rows(min_row=2):
        row[0].value = url_id
        row[1].value = url
        # Update the values in each column of the matched row using variables defined above
        print(row[0].value)
        row[2].value = positive_score
        row[3].value = negative_score
        row[4].value = polarity_score
        row[5].value = subjectivity_score
        row[6].value = avg_sentence_length
        row[7].value = pct_complex_words
        row[8].value = fog_index
        row[9].value = avg_words_per_sentence
        row[10].value = num_complex_words
        row[11].value = num_words
        row[12].value = num_syllables_per_word
        row[13].value = num_personal_pronouns
        row[14].value = avg_word_length


#extract function to extract text data from each url and pass it to the analyze function
def extract(analyze, df, check):
    workbook = openpyxl.load_workbook('Output Data Structure.xlsx')
    worksheet = workbook['Sheet1']
    
    for index, row in df.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        
        response = requests.get(url)
        if response.status_code == 404:
            print(f'404 for {url_id}')
            check.append(url_id)
            continue
    
        soup = BeautifulSoup(response.content, 'html.parser')
    
        if soup.find('h1', class_='entry-title') is None:
            article_title = soup.find('h1', class_='tdb-title-text').text.strip()
            article_text_divs = soup.find_all('div', class_='tdb-block-inner')
            article_text = ''
            for div in article_text_divs:
                if 'td-fix-index' in div['class']:
                    for pre in div.find_all('pre'):
                        pre.replace_with('')
                    article_text += '\n'.join([p.text.strip() for p in div.find_all('p')])
            if not article_text:
                print('Article text not found')
        else:
            article_title = soup.find('h1', class_='entry-title').text.strip()
            article_text_div = soup.find('div', class_='td-post-content tagdiv-type')
            for pre in article_text_div.find_all('pre'):
                pre.replace_with('')
            article_text = article_text_div.text.strip()
            if not article_text:
                print('Article text not found')
 
        if article_text == '':
            print(f'No text found for {url_id}')
            check.append(url_id)
            continue
    
        if article_title == '':
            print(f'No title found for {url_id}')
            check.append(url_id)
            continue

        # Call the analyze function to analyze the text and update the worksheet
        analyze(url_id, worksheet, article_title + ' ' + article_text, url)
    
    workbook.save('Output Data Structure.xlsx')

extract(analyze, df, check)
        
#check to see if any urls were not analyzed
print('To check:')
print(check)


