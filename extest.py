# %%
from urllib.request import urlopen
from bs4 import BeautifulSoup

html = urlopen('http://www.pythonscraping.com/pages/page3.html')
bs = BeautifulSoup(html, 'html.parser')

string = ''
for sibling in bs.find('table', {'id':'giftList'}).tr.next_siblings:
    #print([s for s in sibling.stripped_strings])
    string = string+str(sibling)

# %%
soup = BeautifulSoup(string, 'html.parser')
content = soup.find_all('tr', {'class':'gift'})

# %%
def get_prize():
    prize = []
    counter = 0
    while len(content) > counter:
        dataColumn = content[counter].find_all('td')[2].string.replace('\n', '')
        prize.append(dataColumn)
        
        counter += 1
    return prize
# %%
import pandas as pd

RoadData = pd.read_excel('路段.xlsx')


# %% 
def get_RoadList(District):
    RoadList = RoadData[District].dropna().tolist()
    return RoadList

# %%
from tqdm import tqdm
for i in tqdm(get_RoadList('松山區')):
    print(i)

# %%
