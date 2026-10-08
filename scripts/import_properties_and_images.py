import os, sys, re, csv, json, email, ssl, urllib.request, hashlib
from email import policy
from bs4 import BeautifulSoup
import psycopg

sys.stdout.reconfigure(encoding='utf-8')

# SSL context for downloading external HTML images if needed
ssl_ctx = ssl._create_unverified_context()

def clean_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', name).strip()

def parse_mhtml_file(file_path):
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
                'data': data,
                'size': len(data),
                'url': cl
            })
    return html_text, images

def extract_metadata_from_html(html_text, platform, file_name):
    soup = BeautifulSoup(html_text, 'html.parser')
    text = soup.get_text(separator=' ')
    
    meta = {}
    title = soup.title.string.strip() if soup.title and soup.title.string else file_name
    meta['page_title'] = title
    
    # Check meta tags
    meta_desc = soup.find('meta', attrs={'name': 'description'})
    meta['meta_description'] = meta_desc.get('content', '') if meta_desc else ""
    
    # Extract phone / broker
    phone_match = re.search(r'09\d{2}[-—\s]?\d{3}[-—\s]?\d{3}', text)
    meta['agent_phone'] = phone_match.group(0) if phone_match else None
    
    meta['features'] = []
    # Extract tags or feature keywords
    for tag in soup.find_all(['span', 'div', 'p'], class_=re.compile(r'tag|feature|label|badge|info', re.I)):
        t = tag.get_text().strip()
        if 2 <= len(t) <= 15 and t not in meta['features'] and '\n' not in t:
            meta['features'].append(t)
            if len(meta['features']) >= 12:
                break

    return soup, text, meta

def run_import():
    img_root_dir = os.path.abspath('物件圖片')
    os.makedirs(img_root_dir, exist_ok=True)
    print(f"Image directory initialized at: {img_root_dir}")

    # Load 591 CSV
    csv_map = {}
    if os.path.exists('591看屋物件綜合比較表.csv'):
        with open('591看屋物件綜合比較表.csv', 'r', encoding='utf-8-sig') as f:
            for r in csv.DictReader(f):
                csv_map[r['原始檔案名稱']] = r
    print(f"Loaded {len(csv_map)} items from 591 CSV.")

    # Load house_status.json
    status_map = {}
    if os.path.exists('house_status.json'):
        with open('house_status.json', 'r', encoding='utf-8') as f:
            status_map = json.load(f)

    # Collect all files across 3 folders
    folders = [
        ('待看物件', 'pending'),
        ('看完列入考慮的物件', 'consider'),
        ('看完之後不考慮的物件', 'rejected')
    ]
    
    all_files = []
    for fldr, fldr_status in folders:
        if not os.path.exists(fldr):
            continue
        for fn in sorted(os.listdir(fldr)):
            if fn.endswith('.mhtml') or fn.endswith('.html'):
                all_files.append({
                    'folder': fldr,
                    'file_name': fn,
                    'file_path': os.path.join(fldr, fn),
                    'default_status': fldr_status
                })

    print(f"Found {len(all_files)} total property files in folders.")

    # Connect to PostgreSQL
    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house", autocommit=True)
    with conn.cursor() as cur:
        # Clear existing property records for a clean run
        cur.execute("TRUNCATE TABLE property_images RESTART IDENTITY CASCADE;")
        cur.execute("TRUNCATE TABLE property_sources RESTART IDENTITY CASCADE;")
        cur.execute("TRUNCATE TABLE properties RESTART IDENTITY CASCADE;")

        prop_code_idx = 1
        total_images_saved = 0

        for item in all_files:
            fn = item['file_name']
            fp = item['file_path']
            fldr = item['folder']
            status = item['default_status']

            # Determine platform
            if '591' in fn:
                platform = '591'
            elif '信義' in fn:
                platform = '信義房屋'
            elif '永義' in fn or '新時代' in fn:
                platform = '永義房屋'
            elif '比佛利' in fn or 'ycut' in fn or '傑出名門_南屯' in fn or '大安黃金城_南區' in fn or '文心大國民_【市民' in fn or '靜園A區_樹義' in fn:
                platform = '有巢氏房屋'
            else:
                platform = '591'

            # Determine property code
            csv_row = csv_map.get(fn)
            if csv_row:
                csv_num = int(csv_row['編號'])
                prop_code = f"PROP_{csv_num:02d}"
                # override status if in status_map
                if str(csv_num) in status_map:
                    status = status_map[str(csv_num)]
            else:
                # 41 onwards
                csv_num = 40 + prop_code_idx
                prop_code = f"PROP_{csv_num:02d}"
                prop_code_idx += 1

            # Determine community name
            comm_name = fn.split('_')[0].strip()
            if '（' in comm_name:
                comm_name = comm_name.split('（')[0].strip()
            if '(' in comm_name:
                comm_name = comm_name.split('(')[0].strip()

            # Parse MHTML or HTML
            raw_html = ""
            raw_images = []
            if fn.endswith('.mhtml'):
                raw_html, raw_images = parse_mhtml_file(fp)
            else:
                with open(fp, 'r', encoding='utf-8', errors='ignore') as f:
                    raw_html = f.read()

            soup, full_text, meta = extract_metadata_from_html(raw_html, platform, fn)

            # Extract or map fields
            title = ""
            district = ""
            address = ""
            price_total = None
            unit_price = None
            rooms = None
            living_rooms = None
            bathrooms = None
            balconies = None
            layout_raw = ""
            floor_info = ""
            current_floor = None
            total_floors = None
            age = None
            total_area = None
            indoor_area = None
            attached_area = None
            indoor_total = None
            public_area = None
            land_area = None
            public_ratio = None
            public_ratio_has_parking = False
            public_ratio_desc = ""
            parking_type = ""
            parking_desc = ""
            orientation = ""
            management_fee = ""
            original_url = ""
            description_text = ""
            case_id = ""

            if csv_row:
                comm_name = csv_row['社區名稱']
                title = csv_row['房屋物件名稱']
                district_raw = csv_row['行政區域/商圈']
                district = district_raw.split()[0] if district_raw else ""
                
                try: price_total = float(csv_row['總價(萬)']) if csv_row['總價(萬)'] else None
                except: pass
                try: unit_price = float(csv_row['單價(萬/坪)']) if csv_row['單價(萬/坪)'] else None
                except: pass
                
                layout_raw = csv_row['格局']
                if '房' in layout_raw:
                    rm = re.search(r'(\d+)\s*房', layout_raw)
                    if rm: rooms = int(rm.group(1))
                if '廳' in layout_raw:
                    lm = re.search(r'(\d+)\s*廳', layout_raw)
                    if lm: living_rooms = int(lm.group(1))
                if '衛' in layout_raw:
                    bm = re.search(r'(\d+)\s*衛', layout_raw)
                    if bm: bathrooms = int(bm.group(1))
                if '陽' in layout_raw:
                    am = re.search(r'(\d+)\s*陽', layout_raw)
                    if am: balconies = int(am.group(1))

                floor_info = csv_row['樓層']
                if '/' in floor_info:
                    fl_m = re.search(r'(\d+)/(\d+)F', floor_info)
                    if fl_m:
                        current_floor = int(fl_m.group(1))
                        total_floors = int(fl_m.group(2))

                try: age = float(re.sub(r'[^\d.]', '', csv_row['屋齡'])) if csv_row['屋齡'] else None
                except: pass
                try: total_area = float(re.sub(r'[^\d.]', '', csv_row['權狀總坪數'])) if csv_row['權狀總坪數'] else None
                except: pass
                try: indoor_total = float(re.sub(r'[^\d.]', '', csv_row['室內實坪(主+附)'])) if csv_row['室內實坪(主+附)'] else None
                except: pass
                try: indoor_area = float(re.sub(r'[^\d.]', '', csv_row['主建物坪數'])) if csv_row['主建物坪數'] else None
                except: pass
                try: attached_area = float(re.sub(r'[^\d.]', '', csv_row['附屬建物坪數'])) if csv_row['附屬建物坪數'] else None
                except: pass
                try: public_area = float(re.sub(r'[^\d.]', '', csv_row['共用部分(公設)坪數'])) if csv_row['共用部分(公設)坪數'] else None
                except: pass
                try: land_area = float(re.sub(r'[^\d.]', '', csv_row['土地坪數'])) if csv_row['土地坪數'] else None
                except: pass
                try: public_ratio = float(re.sub(r'[^\d.]', '', csv_row['公設比(%)'])) if csv_row['公設比(%)'] else None
                except: pass

                public_ratio_desc = csv_row['公設比是否含車位']
                public_ratio_has_parking = '是' in public_ratio_desc
                parking_type = csv_row['車位型態']
                parking_desc = csv_row['車位詳細說明']
                orientation = csv_row['座向']
                management_fee = csv_row['管理費']
                original_url = csv_row['物件連結']
            else:
                # Parse directly from filename and HTML
                title = fn.replace('.mhtml', '').replace('.html', '')
                if '_' in title:
                    title = title.split('_', 1)[1]
                
                # Check for district
                for dist_name in ['南屯區', '南區', '西區', '北區', '西屯區', '北屯區', '東區', '中區']:
                    if dist_name in fn or dist_name in meta['meta_description']:
                        district = dist_name
                        break

                # Extract price
                p_match = re.search(r'(\d{3,4})\s*萬', meta['meta_description'] + ' ' + title)
                if p_match:
                    try: price_total = float(p_match.group(1))
                    except: pass

                # Extract layout
                lay_m = re.search(r'(\d+)房(\d+)廳(\d+)衛(?:\s*(\d+)陽)?', meta['meta_description'] + ' ' + title)
                if lay_m:
                    rooms = int(lay_m.group(1))
                    living_rooms = int(lay_m.group(2))
                    bathrooms = int(lay_m.group(3))
                    if lay_m.group(4): balconies = int(lay_m.group(4))
                    layout_raw = f"{rooms}房{living_rooms}廳{bathrooms}衛"
                elif '兩房' in title or '2房' in title:
                    rooms = 2
                elif '三房' in title or '3房' in title:
                    rooms = 3

                # Extract area
                ar_m = re.search(r'(?:坪數|室內|主[＋+]陽約?)\s*([0-9.]+)\s*坪', meta['meta_description'] + ' ' + full_text)
                if ar_m:
                    try: indoor_total = float(ar_m.group(1))
                    except: pass

                # Extract parking
                if '平車' in title or '平面車位' in title:
                    parking_type = '坡道平面式'
                elif '車位' in title or '休旅車' in title:
                    parking_type = '機械式/車位'
                else:
                    parking_type = '無車位/未載明'

            # Extract case_id & URL from HTML
            url_match = re.search(r'https?://[^\s"\'<>]+(?:591\.com\.tw|sinyi\.com\.tw|yungching\.com\.tw|ycut\.com\.tw)[^\s"\'<>]*', raw_html)
            if url_match and not original_url:
                original_url = url_match.group(0)

            # Insert or get Community ID
            cur.execute("""
                INSERT INTO communities (name, district)
                VALUES (%s, %s)
                ON CONFLICT (name) DO UPDATE SET district = COALESCE(communities.district, EXCLUDED.district)
                RETURNING id;
            """, (comm_name, district))
            community_id = cur.fetchone()[0]

            # Insert Property
            cur.execute("""
                INSERT INTO properties (
                    code, community_name, community_id, city, district, address, title,
                    price_total, unit_price, rooms, living_rooms, bathrooms, balconies,
                    layout_raw, floor_info, current_floor, total_floors, age,
                    total_area, indoor_area, attached_area, indoor_total, public_area, land_area,
                    public_ratio, public_ratio_has_parking, public_ratio_desc,
                    parking_type, parking_desc, orientation, management_fee, decision_status,
                    original_source_platform, original_file_name, original_url
                ) VALUES (
                    %s, %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s, %s, %s, %s, %s,
                    %s, %s, %s,
                    %s, %s, %s, %s, %s,
                    %s, %s, %s
                ) RETURNING id;
            """, (
                prop_code, comm_name, community_id, '臺中市', district, address, title,
                price_total, unit_price, rooms, living_rooms, bathrooms, balconies,
                layout_raw, floor_info, current_floor, total_floors, age,
                total_area, indoor_area, attached_area, indoor_total, public_area, land_area,
                public_ratio, public_ratio_has_parking, public_ratio_desc,
                parking_type, parking_desc, orientation, management_fee, status,
                platform, fn, original_url
            ))
            property_id = cur.fetchone()[0]

            # Insert Property Source (Rich JSONB)
            description_text = meta.get('meta_description', '')
            raw_meta_json = json.dumps({
                'source_folder': fldr,
                'extracted_features': meta.get('features', []),
                'agent_phone': meta.get('agent_phone'),
                'page_title': meta.get('page_title'),
                'csv_row': csv_row
            }, ensure_ascii=False)

            cur.execute("""
                INSERT INTO property_sources (
                    property_id, platform, original_case_id, original_title, original_url,
                    local_file_relpath, agent_phone, raw_metadata, description_text
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s);
            """, (
                property_id, platform, case_id, title, original_url,
                os.path.join(fldr, fn), meta.get('agent_phone'), raw_meta_json, description_text
            ))

            # Process Images
            prop_img_dir = os.path.join(img_root_dir, f"{prop_code}_{clean_filename(comm_name)}")
            os.makedirs(prop_img_dir, exist_ok=True)

            img_count = 0
            seen_hashes = set()

            # Case A: MHTML images
            if raw_images:
                # Filter images: keep only real property pictures
                for img in raw_images:
                    u = img['url'].lower()
                    data = img['data']
                    sz = img['size']

                    # Discard UI icons
                    if sz < 12000:
                        continue
                    if 'needle.591' in u or 's.591' in u or 'logo' in u or 'icon' in u or 'banner' in u:
                        continue

                    # Deduplicate identical payloads
                    h = hashlib.md5(data).hexdigest()
                    if h in seen_hashes:
                        continue
                    seen_hashes.add(h)

                    img_count += 1
                    ext = '.jpg'
                    if 'png' in img['content_type'] or '.png' in u: ext = '.png'
                    elif 'webp' in img['content_type'] or '.webp' in u: ext = '.webp'

                    category = 'floor_plan' if ('layout' in u or 'floorplan' in u or '格局' in u) else 'photo'
                    out_fname = f"{category}_{img_count:02d}{ext}"
                    out_path = os.path.join(prop_img_dir, out_fname)
                    rel_path = f"物件圖片/{prop_code}_{clean_filename(comm_name)}/{out_fname}"

                    with open(out_path, 'wb') as img_f:
                        img_f.write(data)

                    cur.execute("""
                        INSERT INTO property_images (
                            property_id, image_category, file_name, file_relpath, file_size, original_url, sort_order
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s);
                    """, (property_id, category, out_fname, rel_path, sz, img['url'], img_count))
                    total_images_saved += 1

            # Case B: Sinyi HTML files (download images if not inlined)
            elif fn.endswith('.html') and '信義' in fn:
                sinyi_case_m = re.search(r'res\.sinyi\.com\.tw/buy/([A-Za-z0-9]+)/', raw_html)
                if sinyi_case_m:
                    case_code = sinyi_case_m.group(1)
                    # Sinyi images are lettered A, B, C, D, ...
                    for char in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                        img_url = f"https://res.sinyi.com.tw/buy/{case_code}/bigimg/{char}.JPG"
                        try:
                            req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                            with urllib.request.urlopen(req, context=ssl_ctx, timeout=3) as resp:
                                data = resp.read()
                                if len(data) > 10000:
                                    img_count += 1
                                    out_fname = f"photo_{img_count:02d}.jpg"
                                    out_path = os.path.join(prop_img_dir, out_fname)
                                    rel_path = f"物件圖片/{prop_code}_{clean_filename(comm_name)}/{out_fname}"
                                    with open(out_path, 'wb') as img_f:
                                        img_f.write(data)
                                    cur.execute("""
                                        INSERT INTO property_images (
                                            property_id, image_category, file_name, file_relpath, file_size, original_url, sort_order
                                        ) VALUES (%s, %s, %s, %s, %s, %s, %s);
                                    """, (property_id, 'photo', out_fname, rel_path, len(data), img_url, img_count))
                                    total_images_saved += 1
                        except Exception:
                            # Break when letter sequence ends
                            break

            print(f"[{prop_code}] {comm_name} ({platform}) -> Status: {status}, Extracted {img_count} images.")

        print(f"\n==========================================")
        print(f"ALL PROPERTIES MIGRATED SUCCESSFULLY!")
        print(f"Total properties: {len(all_files)}")
        print(f"Total house images extracted to disk: {total_images_saved}")
        print(f"==========================================")

    conn.close()

if __name__ == '__main__':
    run_import()
