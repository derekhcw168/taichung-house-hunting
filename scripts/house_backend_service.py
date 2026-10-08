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

def get_db():
    return psycopg.connect(DB_CONFIG, autocommit=True)

def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def scrape_property_url(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8'
    }
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, context=ssl_ctx, timeout=12) as resp:
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
        'age': 30,
        'indoor': 25.0,
        'total_area': 35.0,
        'parking': '車位/未載明',
        'parking_desc': '',
        'management_fee': '未載明',
        'orientation': '坐北朝南',
        'platform': '其他房仲',
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

    # Extract unit price
    up_m = re.search(r'(\d{2,3}(?:\.\d+)?)\s*萬/坪', desc + ' ' + text)
    if up_m:
        data['unit_price'] = float(up_m.group(1))

    # Extract layout
    lay_m = re.search(r'(\d+)房(\d+)廳(\d+)衛(?:\s*(\d+)陽)?', desc + ' ' + text)
    if lay_m:
        data['rooms'] = int(lay_m.group(1))
        data['living_rooms'] = int(lay_m.group(2))
        data['baths'] = int(lay_m.group(3))
        if lay_m.group(4):
            data['balconies'] = int(lay_m.group(4))
        data['layout'] = f"{data['rooms']}房{data['living_rooms']}廳{data['baths']}衛"

    # Extract indoor area
    in_m = re.search(r'(?:主[＋+]陽|室內|主建物)\s*約?\s*([0-9.]+)\s*坪', desc + ' ' + text)
    if in_m:
        data['indoor'] = float(in_m.group(1))

    # Extract total area
    tot_m = re.search(r'(?:權狀|總建|建坪)\s*約?\s*([0-9.]+)\s*坪', desc + ' ' + text)
    if tot_m:
        data['total_area'] = float(tot_m.group(1))
    elif data['indoor']:
        data['total_area'] = round(data['indoor'] * 1.35, 2)

    # Extract floor
    fl_m = re.search(r'(\d+)\s*/\s*(\d+)\s*F', desc + ' ' + text, re.I)
    if fl_m:
        data['current_floor'] = int(fl_m.group(1))
        data['total_floors'] = int(fl_m.group(2))
        data['floor'] = f"{data['current_floor']}/{data['total_floors']}F"

    # Extract age
    age_m = re.search(r'屋齡\s*約?\s*([0-9.]+)\s*年', desc + ' ' + text)
    if age_m:
        data['age'] = float(age_m.group(1))

    # Parking
    if '平車' in title or '平面車位' in desc or '平面' in text:
        data['parking'] = '坡道平面式'
        data['parking_desc'] = '平面車位'
    elif '機械' in title or '機械' in desc:
        data['parking'] = '坡道機械式'
        data['parking_desc'] = '機械車位'

    # Platform specific community name & images
    if 'sinyi.com.tw' in url:
        data['platform'] = '信義房屋'
        # Clean title
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

    elif 'yungyi' in url or 'yungching' in url or 'ycut' in url:
        data['platform'] = '永義/永慶/有巢氏'
        comm_m = re.search(r'社區名稱[：:\s]*([^\s,，。]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        else: data['community'] = title.split('｜')[0].split('_')[0].split('|')[0].strip()
        
        imgs = re.findall(r'https?://(?:cloudfps|yccdn)\.[^\s"\'<>]+\.(?:jpg|png|jpeg)', html)
        data['images'] = list(set(imgs))[:15]

    elif '591.com.tw' in url:
        data['platform'] = '591'
        comm_m = re.search(r'社區[：:\s]*([^\s,，。<]+)', text)
        if comm_m: data['community'] = comm_m.group(1)
        else: data['community'] = title.split(' - ')[0].split('【')[0].split('｜')[0].strip()
        
        imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+!1000x\.[a-z]+', html)
        if not imgs:
            imgs = re.findall(r'https?://img\d*\.591\.com\.tw/house/[^\s"\'<>]+\.(?:jpg|png|webp)', html)
        data['images'] = list(set(imgs))[:20]

    if not data['community']:
        data['community'] = title.split(' ')[0] if title else "精選物件"

    if data['price'] and data['total_area'] and not data['unit_price']:
        data['unit_price'] = round(data['price'] / data['total_area'], 1)

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

        if path == '/api/move':
            self.handle_api_move(prop_id, to_cat)
        elif path == '/api/delete':
            self.handle_api_delete(prop_id)
        elif path == '/api/crawl' or path == '/api/import-url':
            self.handle_api_crawl(url_arg)
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

    def handle_api_delete(self, prop_id):
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
                    
                    # Delete from DB (cascades to sources and images)
                    if str(prop_id).isdigit():
                        cur.execute("DELETE FROM properties WHERE id = %s;", (int(prop_id),))
                    else:
                        cur.execute("DELETE FROM properties WHERE code = %s;", (str(prop_id),))

                    self.send_json({
                        "success": True,
                        "deleted_id": prop_id,
                        "code": code,
                        "community_name": comm_name,
                        "message": f"成功從資料庫刪除物件「{comm_name}」({code})！"
                    })
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
                    # Get max code number
                    cur.execute("SELECT COALESCE(MAX(id), 0) + 1 FROM properties;")
                    next_id = cur.fetchone()[0]
                    prop_code = f"PROP_{next_id:02d}"

                    comm_name = data['community']
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

                    # Download images directly into PostgreSQL (BYTEA) - no disk storage
                    saved_imgs = 0
                    for idx, img_url in enumerate(data['images'], 1):
                        try:
                            req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                            with urllib.request.urlopen(req, context=ssl_ctx, timeout=4) as resp:
                                img_data = resp.read()
                                if len(img_data) > 10000:
                                    ext = '.jpg'
                                    mime = 'image/jpeg'
                                    if 'png' in img_url:
                                        ext = '.png'
                                        mime = 'image/png'
                                    elif 'webp' in img_url:
                                        ext = '.webp'
                                        mime = 'image/webp'
                                    fname = f"photo_{saved_imgs+1:02d}{ext}"
                                    rel_path = f"物件圖片/{prop_code}_{clean_filename(comm_name)}/{fname}"
                                    
                                    cur.execute("""
                                        INSERT INTO property_images (
                                            property_id, image_category, file_name, file_relpath, file_size, original_url, sort_order, image_data, mime_type
                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s);
                                    """, (prop_id, 'photo', fname, rel_path, len(img_data), img_url, saved_imgs+1, img_data, mime))
                                    saved_imgs += 1
                        except Exception:
                            continue

            self.send_json({
                "success": True,
                "property_id": prop_id,
                "code": prop_code,
                "community": comm_name,
                "title": data['title'],
                "price": data['price'],
                "district": data['district'],
                "images_downloaded": saved_imgs,
                "message": f"成功自動抓取並建立「{comm_name}」！已存入 PostgreSQL 待看清單，共下載 {saved_imgs} 張照片。"
            })
        except Exception as e:
            self.send_json({"success": False, "error": f"擷取失敗: {str(e)}"}, status=500)

def run_server():
    server_address = ('127.0.0.1', PORT)
    httpd = http.server.ThreadingHTTPServer(server_address, HouseRequestHandler)
    print(f"==================================================")
    print(f" 台中看屋決策平台 - PostgreSQL 即時服務已啟動")
    print(f" 網址: http://localhost:{PORT}/")
    print(f" 資料庫: taichung_house (PostgreSQL 16)")
    print(f"==================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("Server stopped.")

if __name__ == '__main__':
    run_server()
