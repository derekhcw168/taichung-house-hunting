import email
from email import policy
from bs4 import BeautifulSoup
import os, sys, re, json

sys.stdout.reconfigure(encoding='utf-8')

def inspect_mhtml(file_path):
    print(f"\n====================\nInspecting: {os.path.basename(file_path)}")
    with open(file_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)
    
    html_text = ""
    images = []
    for p in msg.walk():
        ct = p.get_content_type()
        if ct == 'text/html' and not html_text:
            data = p.get_payload(decode=True)
            try:
                html_text = data.decode('utf-8')
            except Exception:
                html_text = data.decode('big5', errors='ignore')
        elif ct.startswith('image/'):
            data = p.get_payload(decode=True) or b''
            cl = p.get('Content-Location', '')
            images.append({
                'content_type': ct,
                'size': len(data),
                'url': cl
            })
    
    print(f"HTML text length: {len(html_text)}")
    print(f"Total images found: {len(images)}")
    
    # Filter house images
    house_images = []
    for img in images:
        u = img['url'].lower()
        sz = img['size']
        if '591.com.tw/house/' in u or 'needle.591' not in u and sz > 15000 and any(ext in u for ext in ['.jpg', '.jpeg', '.webp', '.png', 'image']):
            house_images.append(img)
    
    print(f"Candidate house images: {len(house_images)}")
    if house_images:
        print("Sample house image:", house_images[0]['url'][:80], f"({house_images[0]['size']} bytes)")

    # Parse HTML with BeautifulSoup
    soup = BeautifulSoup(html_text, 'html.parser')
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    print(f"Page Title: {title}")

# Test 4 platforms
test_files = [
    '待看物件/文心寶典_南屯電梯3房2衛走路捷運南屯站三分鐘 - 591售屋網.mhtml',
    '待看物件/朋莊大樓_美村商圈綠園道三房 - 信義房屋.mhtml',
    '看完列入考慮的物件/大墩人文_［降價獨家－專任捷運大三房］南屯路二段 - 信義房屋.html',
    '待看物件/比佛利_比佛利山莊3房車位 _ 次頂樓 _ 雙衛浴開窗.mhtml',
    '看完列入考慮的物件/廣三佛羅里達_佛羅里達｜高樓層空間和採光都有｜舒適大3房 _ 台中市南區忠明南路廣三佛羅里達房屋出售 (YE0157095) _永義房屋-新時代房仲.mhtml'
]

for tf in test_files:
    if tf.endswith('.html'):
        print(f"\n====================\nInspecting HTML: {os.path.basename(tf)}")
        with open(tf, 'r', encoding='utf-8', errors='ignore') as f:
            c = f.read()
        soup = BeautifulSoup(c, 'html.parser')
        print(f"Page Title: {soup.title.string if soup.title else ''}")
    else:
        inspect_mhtml(tf)
