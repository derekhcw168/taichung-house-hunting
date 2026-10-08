import email
from email import policy
from bs4 import BeautifulSoup
import os, sys, re, json

sys.stdout.reconfigure(encoding='utf-8')

# Test parsing the 2 Sinyi HTML files
sinyi_htmls = [
    '看完列入考慮的物件/大墩人文_［降價獨家－專任捷運大三房］南屯路二段 - 信義房屋.html',
    '看完列入考慮的物件/工學天下_［工學天下，次頂樓溫馨三房］工學北路 - 信義房屋.html'
]

for sf in sinyi_htmls:
    print(f"\nTesting Sinyi HTML: {os.path.basename(sf)}")
    with open(sf, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()
    soup = BeautifulSoup(html, 'html.parser')
    
    # Check title
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    print("  Title:", title)
    
    # Check text snippets for price, community, layout
    text = soup.get_text(separator=' ')
    price_match = re.search(r'(\d+)\s*萬', text)
    print("  Price match:", price_match.group(0) if price_match else "None")
    
    # Look for meta tags
    meta_desc = soup.find('meta', attrs={'name': 'description'})
    if meta_desc:
        print("  Meta description:", meta_desc.get('content', '')[:120])

# Test parsing YCUT (有巢氏)
ycut_sample = '待看物件/比佛利_比佛利山莊3房車位 _ 次頂樓 _ 雙衛浴開窗.mhtml'
if os.path.exists(ycut_sample):
    print(f"\nTesting YCUT MHTML: {os.path.basename(ycut_sample)}")
    with open(ycut_sample, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    for p in msg.walk():
        if p.get_content_type() == 'text/html':
            html = p.get_payload(decode=True).decode('utf-8', errors='ignore')
            soup = BeautifulSoup(html, 'html.parser')
            print("  Title:", soup.title.string if soup.title else "")
            meta_desc = soup.find('meta', attrs={'name': 'description'})
            if meta_desc:
                print("  Meta description:", meta_desc.get('content', '')[:150])
            break
