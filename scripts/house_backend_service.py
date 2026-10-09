import http.server
import json
import urllib.request
import urllib.parse
import ssl
import re
import os
import sys
import psycopg
import hashlib
import subprocess
import threading
import datetime
import time
from bs4 import BeautifulSoup
if sys.stdout is None:
    log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "backend.log")
    sys.stdout = open(log_path, "a", encoding="utf-8", buffering=1)
    sys.stderr = sys.stdout
else:
    sys.stdout.reconfigure(encoding='utf-8')

PORT = 8899
DB_CONFIG = "host=localhost port=5432 user=postgres password=postgres dbname=taichung_house"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGE_DIR = os.path.join(BASE_DIR, "物件圖片")
os.makedirs(IMAGE_DIR, exist_ok=True)

ssl_ctx = ssl._create_unverified_context()
CURRENT_TUNNEL_URL = ""

def trigger_async_github_sync(commit_msg):
    def _worker():
        try:
            time.sleep(1) # debounce
            scripts_dir = os.path.join(BASE_DIR, "scripts")
            sys.path.insert(0, scripts_dir)
            import export_all_properties_to_html
            export_all_properties_to_html.export_db_to_html()
            import sync_to_github
            sync_to_github.main(commit_msg)
        except Exception as e:
            print(f"[AsyncSync] GitHub sync failed: {e}")
    threading.Thread(target=_worker, daemon=True).start()

def update_api_config(url):
    global CURRENT_TUNNEL_URL
    CURRENT_TUNNEL_URL = url
    config = {
        "api_url": url,
        "updated_at": datetime.datetime.now(datetime.timezone.utc).isoformat()
    }
    cfg_file = os.path.join(BASE_DIR, "api_config.json")
    with open(cfg_file, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)

    repo_dir = os.path.join(os.path.expanduser('~'), 'Documents', 'GitHub', 'taichung-house-hunting')
    repo_cfg = os.path.join(repo_dir, "api_config.json")
    if os.path.exists(repo_dir):
        with open(repo_cfg, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)

    trigger_async_github_sync(f"Update Cloud API tunnel URL: {url}")

def start_tunnel_thread():
    def _tunnel_loop():
        cloudflared_bin = os.path.join(BASE_DIR, "cloudflared.exe")
        while True:
            try:
                if os.path.exists(cloudflared_bin):
                    cmd = [cloudflared_bin, 'tunnel', '--url', f'http://127.0.0.1:{PORT}']
                    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
                    for line in iter(proc.stdout.readline, ''):
                        m = re.search(r'https://[a-zA-Z0-9\.\-_]+\.trycloudflare\.com', line)
                        if m:
                            t_url = m.group(0)
                            print(f"==================================================")
                            print(f" 🌐 Cloudflare 高速穿透服務已就緒: {t_url}")
                            print(f"==================================================")
                            update_api_config(t_url)
                    proc.wait()
                else:
                    cmd = ['ssh', '-o', 'StrictHostKeyChecking=no', '-o', 'ServerAliveInterval=15', '-o', 'ServerAliveCountMax=3', '-R', f'80:127.0.0.1:{PORT}', 'nokey@localhost.run']
                    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
                    for line in iter(proc.stdout.readline, ''):
                        m = re.search(r'https://[a-zA-Z0-9\.\-_]+\.lhr\.life', line)
                        if m:
                            t_url = m.group(0)
                            print(f"==================================================")
                            print(f" 🌐 雲端穿透服務已就緒: {t_url}")
                            print(f"==================================================")
                            update_api_config(t_url)
                    proc.wait()
            except Exception as e:
                print(f"[TunnelError]: {e}")
            time.sleep(3)
    threading.Thread(target=_tunnel_loop, daemon=True).start()

def get_db():
    return psycopg.connect(DB_CONFIG, autocommit=True)

def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def normalize_property_url(raw_url):
    url = raw_url.strip()
    if '591.com.tw' in url:
        m = re.search(r'(?:/|id=|v2/sale/|detail/|detail/\d+/|house/)(\d{7,9})', url)
        if m:
            return f"https://sale.591.com.tw/home/house/detail/2/{m.group(1)}.html"
    return url

def scrape_property_url(raw_url):
    url = normalize_property_url(raw_url)
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8'
    }
    req = urllib.request.Request(url, headers=headers)
    opener = urllib.request.build_opener(
        urllib.request.HTTPSHandler(context=ssl_ctx),
        urllib.request.HTTPCookieProcessor()
    )
    with opener.open(req, timeout=15) as resp:
        html = resp.read().decode('utf-8', errors='ignore')

    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator=' ')
    title = soup.title.string.strip() if soup.title and soup.title.string else ""
    
    meta_desc = soup.find('meta', attrs={'name': 'description'})
    desc = meta_desc.get('content', '') if meta_desc else ''

    data = {
        'url': url,
        'original_url': raw_url,
        'title': title,
        'meta_description': desc,
        'community': '',
        'district': '南屯區',
        'address': '',
        'price': 0,
        'unit_price': 0,
        'rooms': 3,
        'living_rooms': 2,
        'baths': 2,
        'balconies': 1,
        'layout': '3房2廳2衛',
        'floor': '',
        'current_floor': 5,
        'total_floors': 10,
        'age': 20,
        'indoor': 25.0,
        'total_area': 35.0,
        'parking': '車位/未載明',
        'parking_desc': '',
        'management_fee': '未載明',
        'orientation': '坐北朝南',
        'platform': '591' if '591' in url else '其他房仲',
        'images': []
    }

    # 1. Parse JSON-LD Schema (Common on 591 and modern real estate sites)
    for s in soup.find_all('script', attrs={'type': 'application/ld+json'}):
        try:
            ld = json.loads(s.get_text())
            graph = ld.get('@graph', [ld]) if isinstance(ld, dict) else [ld]
            for item in graph:
                if isinstance(item, dict) and ('Apartment' in item.get('@type', []) or 'Product' in item.get('@type', []) or item.get('offers')):
                    if item.get('name'): data['title'] = item['name']
                    if item.get('offers', {}).get('price'): data['price'] = float(item['offers']['price']) / 10000.0
                    if item.get('floorSize', {}).get('value'): data['total_area'] = float(item['floorSize']['value'])
                    if item.get('address', {}).get('addressLocality'): data['district'] = item['address']['addressLocality']
                    if item.get('address', {}).get('streetAddress'): data['address'] = item['address']['streetAddress']
                    if item.get('numberOfRooms'): data['rooms'] = int(item['numberOfRooms'])
                    if item.get('image'): data['images'] = item['image'] if isinstance(item['image'], list) else [item['image']]
                    break
        except Exception:
            pass

    # 2. Extract District if not set
    for dist in ['南屯區', '南區', '西區', '北區', '西屯區', '北屯區', '東區', '中區', '大里區', '太平區', '潭子區']:
        if dist in desc or dist in text or dist in title:
            data['district'] = dist
            break

    # 3. Extract Price if not found in JSON-LD
    if not data['price']:
        pm = re.search(r'(\d{3,4})\s*萬', desc + ' ' + title)
        if pm: data['price'] = float(pm.group(1))

    # 4. Extract Layout
    lay_m = re.search(r'(\d+)房(\d+)廳(\d+)衛(?:\s*(\d+)陽)?', desc + ' ' + text)
    if lay_m:
        data['rooms'] = int(lay_m.group(1))
        data['living_rooms'] = int(lay_m.group(2))
        data['baths'] = int(lay_m.group(3))
        if lay_m.group(4):
            data['balconies'] = int(lay_m.group(4))
        data['layout'] = f"{data['rooms']}房{data['living_rooms']}廳{data['baths']}衛"
    elif data['rooms']:
        data['layout'] = f"{data['rooms']}房2廳2衛"

    # 5. Extract Total Area & Indoor Area
    if not data['total_area']:
        tot_m = re.search(r'(?:權狀|總建|建坪|面積)\s*約?\s*([0-9.]+)\s*坪', desc + ' ' + text)
        if tot_m: data['total_area'] = float(tot_m.group(1))

    in_m = re.search(r'(?:主[＋+]陽|室內|主建物)\s*約?\s*([0-9.]+)\s*坪', desc + ' ' + text)
    if in_m:
        data['indoor'] = float(in_m.group(1))
    elif data['total_area']:
        data['indoor'] = round(data['total_area'] * 0.62, 1)

    if data['price'] and data['indoor'] and not data['unit_price']:
        data['unit_price'] = round(data['price'] / data['indoor'], 1)

    # 6. Extract Floor
    fl_m = re.search(r'(\d+)\s*/\s*(\d+)\s*F', desc + ' ' + text, re.I)
    if fl_m:
        data['current_floor'] = int(fl_m.group(1))
        data['total_floors'] = int(fl_m.group(2))
        data['floor'] = f"{data['current_floor']}/{data['total_floors']}F"
    else:
        fl_s = re.search(r'(高樓層|中樓層|低樓層|\d+樓)', desc + ' ' + text)
        data['floor'] = fl_s.group(1) if fl_s else '高樓層'

    # 7. Extract Age
    age_m = re.search(r'屋齡\s*約?\s*([0-9.]+)\s*年', desc + ' ' + text)
    if age_m:
        data['age'] = float(age_m.group(1))

    # 8. Parking
    if any(k in (data['title'] + ' ' + desc + ' ' + text) for k in ['平車', '平面車位', '坡道平面', '平面']):
        data['parking'] = '坡道平面式'
        data['parking_desc'] = '平面車位'
    elif any(k in (data['title'] + ' ' + desc + ' ' + text) for k in ['機械', '機械車位', '機械式']):
        data['parking'] = '坡道機械式'
        data['parking_desc'] = '機械車位'

    # 9. Platform Specific Community & Image Handling
    if '591.com.tw' in url:
        data['platform'] = '591'
        # 591 descriptions almost always have "位於[社區名稱]"
        comm_m = re.search(r'位於([^，,。!！\s]+)', desc)
        if comm_m and comm_m.group(1).strip() not in ['台中市', '南屯區', '西區', '北區', '住宅']:
            data['community'] = comm_m.group(1).strip()
        else:
            comm_m = re.search(r'社區[：:\s]*([^\s,，。<]+)', text)
            if comm_m and comm_m.group(1).strip() not in ['找房', '行情', '推薦']:
                data['community'] = comm_m.group(1).strip()

        if not data['community'] and '|' in data['title']:
            parts = [p.strip() for p in data['title'].split('|') if p.strip()]
            if len(parts) >= 2: data['community'] = parts[1]

        if not data['images']:
            imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+!1000x\.[a-z]+', html)
            if not imgs:
                imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+\.(?:jpg|png|webp)', html)
            data['images'] = list(set(imgs))[:20]

    elif 'sinyi.com.tw' in url or 'sinyi.in' in url:
        data['platform'] = '信義房屋'
        clean_t = title.split(' - ')[0]
        if '［' in clean_t and '］' in clean_t:
            data['community'] = clean_t.split('［')[1].split('］')[0].replace('降價獨家－', '').replace('專任', '').strip()
        else:
            data['community'] = clean_t
        case_m = re.search(r'/(?:buy/house/|o/)([A-Za-z0-9]+)', url)
        if case_m:
            case_id = case_m.group(1)
            for c in "ABCDEFGHIJKLMNOP":
                data['images'].append(f"https://res.sinyi.com.tw/buy/{case_id}/bigimg/{c}.JPG")

    elif 'yungyi' in url or 'yungching' in url or 'ycut' in url or 'u-trust' in url:
        data['platform'] = '永義/永慶/有巢氏'
        comm_m = re.search(r'社區名稱[：:\s]*([^\s,，。]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        else: data['community'] = title.split('｜')[0].split('_')[0].split('|')[0].strip()
        imgs = re.findall(r'https?://(?:cloudfps|yccdn)\.[^\s"\'<>]+\.(?:jpg|png|jpeg)', html)
        data['images'] = list(set(imgs))[:15]

    if not data['community']:
        data['community'] = title.split(' ')[0] if title else "精選待看物件"

    return data

class HouseRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Enable CORS and Chrome Private Network Access (PNA)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, DELETE')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Access-Control-Request-Private-Network, *')
        self.send_header('Access-Control-Allow-Private-Network', 'true')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == '/api/status':
            self.handle_api_status()
        elif path == '/api/properties':
            self.handle_api_get_properties()
        elif path == '/api/move':
            # Support both GET and POST for compatibility with existing frontend
            prop_id = query.get('id', [None])[0]
            to_cat = query.get('to', [None])[0]
            self.handle_api_move(prop_id, to_cat)
        elif path == '/api/delete':
            prop_id = query.get('id', [None])[0]
            self.handle_api_delete(prop_id)
        elif path == '/api/config' or path == '/api_config.json':
            cfg_file = os.path.join(BASE_DIR, "api_config.json")
            if os.path.exists(cfg_file):
                with open(cfg_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Connection', 'close')
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_json({"api_url": CURRENT_TUNNEL_URL})
        elif path == '/' or path == '/index.html':
            # Serve the main HTML dashboard
            html_file = os.path.join(BASE_DIR, "591看屋物件地圖與比較分析.html")
            if os.path.exists(html_file):
                with open(html_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Connection', 'close')
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_error(404, "HTML file not found")
        elif path == '/favicon.ico':
            ico_file = os.path.join(BASE_DIR, "favicon.ico")
            if os.path.exists(ico_file):
                with open(ico_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'image/x-icon')
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Connection', 'close')
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_error(404, "Favicon not found")
        elif path.startswith('/api/properties/') and path.endswith('/images'):
            parts = path.strip('/').split('/')
            if len(parts) == 4 and parts[2].isdigit():
                self.handle_api_property_images(int(parts[2]))
            else:
                self.send_error(400, "Invalid property id")
        elif path.startswith('/api/images/') or path.startswith('/物件圖片/'):
            # Serve image directly from PostgreSQL (BYTEA)
            self.handle_api_serve_image(path)
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length) if content_length > 0 else b'{}'
        try:
            req_data = json.loads(body.decode('utf-8'))
        except Exception:
            req_data = {}

        prop_id = req_data.get('id') or query.get('id', [None])[0]
        to_cat = req_data.get('to') or query.get('to', [None])[0]
        url_arg = req_data.get('url') or query.get('url', [None])[0]
        perm_param = query.get('permanent', ['0'])[0].lower()
        permanent = bool(req_data.get('permanent')) or (perm_param in ['1', 'true', 'yes'])

        if path == '/api/move':
            self.handle_api_move(prop_id, to_cat)
        elif path == '/api/delete':
            self.handle_api_delete(prop_id, permanent=permanent)
        elif path == '/api/crawl' or path == '/api/import-url':
            self.handle_api_crawl(url_arg)
        elif path == '/api/quick-add':
            self.handle_api_quick_add(req_data)
        else:
            self.send_error(404, "Endpoint not found")

    def send_json(self, data, status=200):
        body_bytes = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body_bytes)))
        self.send_header('Connection', 'close')
        self.end_headers()
        self.wfile.write(body_bytes)

    def handle_api_status(self):
        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT COUNT(*) FROM properties;")
                    count = cur.fetchone()[0]
                    cur.execute("SELECT decision_status, COUNT(*) FROM properties GROUP BY decision_status;")
                    cat_counts = dict(cur.fetchall())
            self.send_json({
                "online": True,
                "database": "taichung_house",
                "total_properties": count,
                "status": cat_counts
            })
        except Exception as e:
            self.send_json({"online": False, "error": str(e)}, status=500)

    def handle_api_get_properties(self):
        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT 
                            p.id, p.code, p.community_name, p.title, p.price_total, p.unit_price,
                            p.rooms, p.living_rooms, p.bathrooms, p.layout_raw, p.floor_info,
                            p.current_floor, p.total_floors, p.age, p.total_area, p.indoor_area,
                            p.attached_area, p.indoor_total, p.public_area, p.public_ratio,
                            p.public_ratio_has_parking, p.public_ratio_desc, p.parking_type,
                            p.parking_desc, p.orientation, p.management_fee, p.decision_status,
                            p.original_source_platform, p.original_url, p.district,
                            COUNT(pi.id) as img_count,
                            MIN(pi.file_relpath) as sample_img
                        FROM properties p
                        LEFT JOIN property_images pi ON p.id = pi.property_id
                        GROUP BY p.id
                        ORDER BY 
                            CASE p.decision_status
                                WHEN 'consider' THEN 1
                                WHEN 'pending' THEN 2
                                WHEN 'rejected' THEN 3
                                ELSE 4
                            END,
                            p.id ASC;
                    """)
                    rows = cur.fetchall()
                    cols = [desc[0] for desc in cur.description]
                    
                    properties = []
                    for r in rows:
                        item = dict(zip(cols, r))
                        # format floats and decimals for JSON
                        for k, v in item.items():
                            if hasattr(v, '__float__'):
                                item[k] = float(v)
                        
                        # Generate map coordinates (approximate if not precise)
                        dist = item.get('district') or '南屯區'
                        coords = {
                            '南屯區': (24.138, 120.645),
                            '南區': (24.118, 120.662),
                            '西區': (24.145, 120.660),
                            '西屯區': (24.165, 120.645),
                            '北區': (24.160, 120.680),
                            '北屯區': (24.175, 120.690),
                            '東區': (24.135, 120.695),
                            '中區': (24.140, 120.680)
                        }
                        base_lat, base_lng = coords.get(dist, (24.138, 120.650))
                        # add reproducible deterministic jitter based on ID
                        jitter_lat = ((item['id'] * 17) % 50 - 25) * 0.0006
                        jitter_lng = ((item['id'] * 23) % 50 - 25) * 0.0006
                        item['lat'] = round(base_lat + jitter_lat, 5)
                        item['lng'] = round(base_lng + jitter_lng, 5)

                        properties.append(item)

            self.send_json({"success": True, "properties": properties})
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_api_move(self, prop_id, target_category):
        if not prop_id or not target_category:
            self.send_json({"success": False, "error": "Missing id or target category"}, status=400)
            return

        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    # check if prop_id is numeric or code
                    if str(prop_id).isdigit():
                        cur.execute("UPDATE properties SET decision_status = %s, updated_at = NOW() WHERE id = %s RETURNING id, community_name, decision_status;", (target_category, int(prop_id)))
                    else:
                        cur.execute("UPDATE properties SET decision_status = %s, updated_at = NOW() WHERE code = %s RETURNING id, community_name, decision_status;", (target_category, str(prop_id)))
                    
                    row = cur.fetchone()
                    if row:
                        self.send_json({"success": True, "id": row[0], "name": row[1], "status": row[2]})
                    else:
                        self.send_json({"success": False, "error": "Property not found"}, status=404)
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_api_delete(self, prop_id, permanent=False):
        if not prop_id:
            self.send_json({"success": False, "error": "Missing id to delete"}, status=400)
            return

        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    if str(prop_id).isdigit():
                        cur.execute("SELECT code, community_name FROM properties WHERE id = %s;", (int(prop_id),))
                    else:
                        cur.execute("SELECT code, community_name FROM properties WHERE code = %s;", (str(prop_id),))
                    
                    row = cur.fetchone()
                    if not row:
                        self.send_json({"success": False, "error": "Property not found"}, status=404)
                        return
                    
                    code, comm_name = row[0], row[1]
                    
                    if permanent:
                        # Permanent Hard Delete from DB
                        if str(prop_id).isdigit():
                            cur.execute("DELETE FROM properties WHERE id = %s;", (int(prop_id),))
                        else:
                            cur.execute("DELETE FROM properties WHERE code = %s;", (str(prop_id),))
                        msg = f"成功從資料庫徹底清除物件「{comm_name}」({code})！"
                    else:
                        # Soft Delete: move to category 'deleted'
                        if str(prop_id).isdigit():
                            cur.execute("UPDATE properties SET decision_status = 'deleted', updated_at = NOW() WHERE id = %s;", (int(prop_id),))
                        else:
                            cur.execute("UPDATE properties SET decision_status = 'deleted', updated_at = NOW() WHERE code = %s;", (str(prop_id),))
                        msg = f"已將物件「{comm_name}」({code})移至「刪除物件」分類！"

                    self.send_json({
                        "success": True,
                        "deleted_id": prop_id,
                        "permanent": permanent,
                        "code": code,
                        "community_name": comm_name,
                        "message": msg
                    })
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_api_property_images(self, prop_id):
        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    cur.execute("""
                        SELECT id, file_name, file_relpath, image_category, mime_type, file_size 
                        FROM property_images 
                        WHERE property_id = %s 
                        ORDER BY sort_order ASC, id ASC;
                    """, (int(prop_id),))
                    rows = cur.fetchall()
                    imgs = [{
                        "id": r[0],
                        "file_name": r[1],
                        "file_relpath": r[2],
                        "category": r[3],
                        "mime": r[4],
                        "size": r[5],
                        "url": f"/api/images/{r[0]}"
                    } for r in rows]
            self.send_json({"success": True, "property_id": prop_id, "count": len(imgs), "images": imgs})
        except Exception as e:
            self.send_json({"success": False, "error": str(e)}, status=500)

    def handle_api_serve_image(self, path):
        try:
            with get_db() as conn:
                with conn.cursor() as cur:
                    row = None
                    if path.startswith('/api/images/'):
                        parts = path[len('/api/images/'):].strip('/').split('/')
                        if len(parts) == 1 and parts[0].isdigit():
                            cur.execute("SELECT image_data, mime_type FROM property_images WHERE id = %s;", (int(parts[0]),))
                            row = cur.fetchone()
                        elif len(parts) >= 2 and parts[0].isdigit():
                            prop_id = int(parts[0])
                            file_name = urllib.parse.unquote(parts[1])
                            cur.execute("SELECT image_data, mime_type FROM property_images WHERE property_id = %s AND file_name = %s LIMIT 1;", (prop_id, file_name))
                            row = cur.fetchone()
                    elif path.startswith('/物件圖片/'):
                        rel_sub = '物件圖片/' + urllib.parse.unquote(path[len('/物件圖片/'):])
                        cur.execute("SELECT image_data, mime_type FROM property_images WHERE file_relpath = %s LIMIT 1;", (rel_sub,))
                        row = cur.fetchone()
                        if not row:
                            file_name = os.path.basename(rel_sub)
                            cur.execute("SELECT image_data, mime_type FROM property_images WHERE file_name = %s LIMIT 1;", (file_name,))
                            row = cur.fetchone()

                    if row and row[0]:
                        img_bytes = bytes(row[0])
                        mime = row[1] or 'image/jpeg'
                        self.send_response(200)
                        self.send_header('Content-Type', mime)
                        self.send_header('Content-Length', str(len(img_bytes)))
                        self.send_header('Cache-Control', 'public, max-age=604800')
                        self.send_header('Connection', 'close')
                        self.end_headers()
                        self.wfile.write(img_bytes)
                        return
                    else:
                        self.send_error(404, "Image not found in PostgreSQL")
        except Exception as e:
            self.send_error(500, f"Database image fetch error: {e}")

    def handle_api_crawl(self, url):
        if not url:
            self.send_json({"success": False, "error": "網址不能為空"}, status=400)
            return

        try:
            print(f"Scraping URL: {url}")
            data = scrape_property_url(url)
            
            with get_db() as conn:
                with conn.cursor() as cur:
                    # Check if property with this URL already exists
                    cur.execute("""
                        SELECT id, community_name FROM properties 
                        WHERE original_url = %s OR original_url = %s
                        LIMIT 1;
                    """, (url, data['url']))
                    existing = cur.fetchone()
                    if existing:
                        self.send_json({
                            "success": True,
                            "property_id": existing[0],
                            "community": existing[1],
                            "message": f"「{existing[1]}」已經在待看清單資料庫中！"
                        })
                        return

                    # Get max code number
                    cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM properties;")
                    next_id = cur.fetchone()[0]
                    prop_code = f"PROP_{next_id:02d}"

                    comm_name = data['community'] or data['title'].split(' ')[0] or "精選待看物件"
                    # Ensure community exists
                    cur.execute("""
                        INSERT INTO communities (name, district)
                        VALUES (%s, %s)
                        ON CONFLICT (name) DO UPDATE SET district = COALESCE(communities.district, EXCLUDED.district)
                        RETURNING id;
                    """, (comm_name, data['district']))
                    comm_id = cur.fetchone()[0]

                    # Insert property into DB
                    cur.execute("""
                        INSERT INTO properties (
                            code, community_name, community_id, city, district, address, title,
                            price_total, unit_price, rooms, living_rooms, bathrooms, balconies,
                            layout_raw, floor_info, current_floor, total_floors, age,
                            total_area, indoor_total, parking_type, parking_desc, management_fee,
                            orientation, decision_status, original_source_platform, original_url
                        ) VALUES (
                            %s, %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s,
                            %s, 'pending', %s, %s
                        ) RETURNING id;
                    """, (
                        prop_code, comm_name, comm_id, '臺中市', data['district'], data['address'], data['title'],
                        data['price'], data['unit_price'], data['rooms'], data['living_rooms'], data['baths'], data['balconies'],
                        data['layout'], data['floor'], data['current_floor'], data['total_floors'], data['age'],
                        data['total_area'], data['indoor'], data['parking'], data['parking_desc'], data['management_fee'],
                        data['orientation'], data['platform'], url
                    ))
                    prop_id = cur.fetchone()[0]

                    # Insert source
                    meta_json = json.dumps({
                        'auto_scraped': True,
                        'source_url': url,
                        'meta_description': data['meta_description']
                    }, ensure_ascii=False)
                    cur.execute("""
                        INSERT INTO property_sources (
                            property_id, platform, original_title, original_url, raw_metadata, description_text
                        ) VALUES (%s, %s, %s, %s, %s, %s);
                    """, (prop_id, data['platform'], data['title'], url, meta_json, data['meta_description']))

                    # Store cover image URL reference without downloading heavy binary files
                    if data.get('images'):
                        first_img = data['images'][0]
                        cur.execute("""
                            INSERT INTO property_images (
                                property_id, image_category, file_name, file_relpath, file_size, original_url, sort_order
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s);
                        """, (prop_id, 'photo', 'cover.jpg', first_img, 0, first_img, 1))

            # Automatically trigger async sync to HTML & GitHub
            trigger_async_github_sync(f"Auto-add crawled property {comm_name} (#{prop_id})")

            self.send_json({
                "success": True,
                "property_id": prop_id,
                "code": prop_code,
                "community": comm_name,
                "title": data['title'],
                "price": data['price'],
                "district": data['district'],
                "message": f"成功自動解析「{comm_name}」並寫入 PostgreSQL 資料庫！已即時加入待看清單。"
            })
        except Exception as e:
            self.send_json({"success": False, "error": f"擷取失敗: {str(e)}"}, status=500)

    def handle_api_quick_add(self, req_data):
        try:
            name = (req_data.get('name') or '').strip()
            if not name:
                self.send_json({"success": False, "error": "請填寫社區或建案名稱"}, status=400)
                return

            district = req_data.get('district') or '南屯區'
            price = float(req_data.get('price') or 0)
            layout = req_data.get('layout') or '3房2廳2衛'
            indoor = float(req_data.get('indoor') or 25.0)
            total_area = float(req_data.get('total_area') or (round(indoor * 1.38, 1)))
            unit_price = round(price / indoor, 1) if (indoor > 0 and price > 0) else 0.0
            floor = req_data.get('floor') or '高樓層'
            age = float(req_data.get('age') or 25.0)
            parking = req_data.get('parking') or '平面車位'
            url = req_data.get('url') or ''
            note = req_data.get('note') or '家人雲端快速收錄'
            category = req_data.get('category') or 'pending'

            with get_db() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT id FROM communities WHERE name = %s LIMIT 1;", (name,))
                    crow = cur.fetchone()
                    if crow:
                        comm_id = crow[0]
                    else:
                        cur.execute("""
                            INSERT INTO communities (name, district, city)
                            VALUES (%s, %s, %s) RETURNING id;
                        """, (name, district, '臺中市'))
                        comm_id = cur.fetchone()[0]

                    cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM properties;")
                    next_id = cur.fetchone()[0]
                    prop_code = f"PROP_{next_id:02d}"

                    cur.execute("""
                        INSERT INTO properties (
                            code, community_name, community_id, city, district, address, title,
                            price_total, unit_price, rooms, living_rooms, bathrooms, balconies,
                            layout_raw, floor_info, current_floor, total_floors, age,
                            total_area, indoor_total, parking_type, parking_desc, management_fee,
                            orientation, decision_status, original_source_platform, original_url
                        ) VALUES (
                            %s, %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s,
                            %s, %s, %s, %s, %s,
                            %s, %s, %s, %s
                        ) RETURNING id;
                    """, (
                        prop_code, name, comm_id, '臺中市', district, '', note,
                        price, unit_price, 3, 2, 2, 1,
                        layout, floor, 5, 14, age,
                        total_area, indoor, parking, parking, '未載明',
                        '坐北朝南', category, '家人雲端收錄', url
                    ))
                    prop_id = cur.fetchone()[0]

                    if url:
                        meta_json = json.dumps({'user_added': True, 'url': url}, ensure_ascii=False)
                        cur.execute("""
                            INSERT INTO property_sources (
                                property_id, platform, original_title, original_url, raw_metadata, description_text
                            ) VALUES (%s, %s, %s, %s, %s, %s);
                        """, (prop_id, '家人雲端收錄', name, url, meta_json, note))

            # Automatically trigger async sync to HTML & GitHub
            trigger_async_github_sync(f"Auto-add cloud quick property: {name} (#{prop_id})")

            self.send_json({
                "success": True,
                "property_id": prop_id,
                "code": prop_code,
                "community": name,
                "price": price,
                "district": district,
                "message": f"成功收錄「{name}」至 PostgreSQL 資料庫！已自動加入待看清單。"
            })
        except Exception as e:
            self.send_json({"success": False, "error": f"收錄失敗: {str(e)}"}, status=500)

def run_server():
    server_address = ('0.0.0.0', PORT)
    httpd = http.server.ThreadingHTTPServer(server_address, HouseRequestHandler)
    print(f"==================================================")
    print(f" 台中看屋決策平台 - PostgreSQL 雲端即時服務已啟動")
    print(f" 辦公室本機網址: http://localhost:{PORT}/")
    print(f" 資料庫: taichung_house (PostgreSQL 16)")
    print(f" 正在啟動雲端安全穿透通道 (供家人遠端同步)...")
    print(f"==================================================")
    start_tunnel_thread()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("Server stopped.")

if __name__ == '__main__':
    run_server()
