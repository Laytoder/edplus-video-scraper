import pandas as pd
from youtubesearchpython import VideosSearch
from colorama import Fore, Style

doc = pd.read_excel('./Topics.xlsx', usecols=['Module A Exam Topics', '12 questions'])
topics = doc['Module A Exam Topics'].tolist()
questions = doc['12 questions'].tolist()

ignore = ['Module B Exam Topics', 'Module C Exam Topics', 'Final Exam Topics']

def hyperlink(url):
    return '=HYPERLINK("%s", "%s")' % (url, url)

links = []

idx = 0
for topic in topics:
    if topic in ignore or isinstance(topic, float):
        print('ignored')
        if isinstance(topic, float):
            links.append('')
        else:
            links.append(questions[idx])
    else:
        query = topic + ' ASU'
        videosSearch = VideosSearch(query, limit = 5)
        added = False
        for video in videosSearch.result()['result']:
            print()
            print(Fore.GREEN + '[FOUND VIDEO]')
            print(Style.RESET_ALL, end='')
            print('Original Title: ', end = '')
            print(Fore.BLUE + topic)
            print(Style.RESET_ALL, end='')
            print('Title: ', end = '')
            print(Fore.BLUE + video['title'])
            print(Style.RESET_ALL, end='')
            print(video['link'])
            if video['title'].lower() == ('Topic: ' + topic).lower() or video['title'].lower() == topic.lower():
                print(Fore.GREEN + 'AUTOMATICALLY APPROVED')
                print(Style.RESET_ALL, end='')
                links.append(hyperlink(video['link']))
                added = True
                break
            approve = (input('Approve Video(Y/N): ') or 'Y') == 'Y'
            if approve:
                links.append(hyperlink(video['link']))
                added = True
                break
        if not added:
            links.append('')
        idx += 1

doc['12 questions'] = pd.Series(links)

writer = pd.ExcelWriter('TOPICS_GENERATED.xlsx', engine='xlsxwriter')
doc.to_excel(writer, sheet_name='Sheet1', index=False)
writer.close()