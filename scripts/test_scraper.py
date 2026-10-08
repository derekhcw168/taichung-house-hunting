import urllib.request, ssl, re, json
from bs4 import BeautifulSoup
import sys

sys.stdout.reconfigure(encoding='utf-8')
ctx = ssl._create_unverified_context()

def scrape_url(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8'
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
        html = resp.read().decode('utf-8', errors='ignore')

    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator=' ')
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    
    meta_desc = soup.find('meta', attrs={'name': 'description'})
    desc = meta_desc.get('content', '') if meta_desc else ''

    data = {
        'url': url,
        'title': title,
        'meta_description': desc,
        'community': '',
        'district': '',
        'price': None,
        'unit_price': None,
        'rooms': 3,
        'baths': 2,
        'layout': '3房2廳2衛',
        'floor': '',
        'age': None,
        'indoor': None,
        'total_area': None,
        'parking': '車位/未載明',
        'images': []
    }

    # Extract district
    for dist in ['南屯區', '南區', '西區', '北區', '西屯區', '北屯區', '東區', '中區', '大里區', '太平區', '潭子區']:
        if dist in desc or dist in text or dist in title:
            data['district'] = dist
            break

    # Extract price
    pm = re.search(r'(\d{3,4})\s*萬', desc + ' ' + title)
    if pm:
        data['price'] = float(pm.group(1))

    # Extract layout
    lay_m = re.search(r'(\d+)房(\d+)廳(\d+)衛', desc + ' ' + text)
    if lay_m:
        data['rooms'] = int(lay_m.group(1))
        data['baths'] = int(lay_m.group(3))
        data['layout'] = f"{data['rooms']}房{lay_m.group(2)}廳{data['baths']}衛"

    # Extract indoor area
    in_m = re.search(r'(?:主[＋+]陽|室內|主建物)\s*約?\s*([0-9.]+)\s*坪', desc + ' ' + text)
    if in_m:
        data['indoor'] = float(in_m.group(1))

    # Platform specific community name & images
    if 'sinyi.com.tw' in url:
        data['platform'] = '信義房屋'
        # Community
        comm_m = re.search(r'社區[：:\s]*([^\s,，。]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        elif '［' in title and '］' in title:
            data['community'] = title.split('［')[1].split('］')[0].replace('降價獨家－', '').replace('專任', '')
        else:
            data['community'] = title.split(' - ')[0]
        
        # Sinyi images
        case_m = re.search(r'/(?:buy/house/|o/)([A-Za-z0-9]+)', url)
        if case_m:
            case_id = case_m.group(1)
            for c in "ABCDEFGHIJKLMNOP":
                img_url = f"https://res.sinyi.com.tw/buy/{case_id}/bigimg/{c}.JPG"
                data['images'].append(img_url)

    elif 'yungyi' in url or 'yungching' in url or 'ycut' in url:
        data['platform'] = '永義/永慶/有巢氏'
        comm_m = re.search(r'社區名稱[：:\s]*([^\s,，。]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        else:
            data['community'] = title.split('｜')[0].split('_')[0].strip()
        # Find images
        imgs = re.findall(r'https?://(?:cloudfps|yccdn)\.[^\s"\'<>]+\.(?:jpg|png|jpeg)', html)
        data['images'] = list(set(imgs))[:15]

    elif '591.com.tw' in url:
        data['platform'] = '591'
        comm_m = re.search(r'社區[：:\s]*([^\s,，。<]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        else:
            data['community'] = title.split(' - ')[0].split('【')[0].split('｜')[0].strip()
        imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+!1000x\.[a-z]+', html)
        if not imgs:
            imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+\.(?:jpg|png|webp)', html)
        data['images'] = list(set(imgs))[:20]

    return data

test_sinyi = scrape_url('https://www.sinyi.com.tw/buy/house/1975QD')
print('Scraped Sinyi:', json.dumps({k: v for k, v in test_sinyi.items() if k != 'images'}, ensure_ascii=False, indent=2))
print('Sinyi candidate images:', len(test_sinyi['images']))
