import os
import sys
import mimetypes
import psycopg

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_CONFIG = "host=localhost port=5432 user=postgres password=postgres dbname=taichung_house"

def migrate_images():
    print("==================================================")
    print("🚀 開始將實體圖檔遷移至 PostgreSQL (BYTEA)...")
    print("==================================================")

    conn = psycopg.connect(DB_CONFIG, autocommit=True)
    cur = conn.cursor()

    cur.execute("SELECT id, file_relpath, file_name FROM property_images ORDER BY id;")
    rows = cur.fetchall()
    total = len(rows)
    print(f"資料庫中共有 {total} 筆圖片記錄需要填充二進位資料。")

    success_count = 0
    fail_count = 0
    total_bytes = 0

    batch = []
    batch_size = 100

    for idx, (img_id, file_relpath, file_name) in enumerate(rows, start=1):
        full_path = os.path.join(BASE_DIR, file_relpath)
        if not os.path.exists(full_path):
            print(f"❌ 找不到實體檔案: {full_path}")
            fail_count += 1
            continue

        try:
            with open(full_path, 'rb') as f:
                data = f.read()

            ext = os.path.splitext(file_name)[1].lower()
            if ext == '.webp':
                mime = 'image/webp'
            elif ext in ('.jpg', '.jpeg'):
                mime = 'image/jpeg'
            elif ext == '.png':
                mime = 'image/png'
            elif ext == '.gif':
                mime = 'image/gif'
            else:
                mime = mimetypes.guess_type(file_name)[0] or 'application/octet-stream'

            batch.append((data, mime, img_id))
            total_bytes += len(data)
            success_count += 1

            if len(batch) >= batch_size:
                with conn.cursor() as b_cur:
                    for b_data, b_mime, b_id in batch:
                        b_cur.execute(
                            "UPDATE property_images SET image_data = %s, mime_type = %s WHERE id = %s;",
                            (b_data, b_mime, b_id)
                        )
                batch = []
                print(f"已處理進度: {idx}/{total} ({idx/total*100:.1f}%) ...")

        except Exception as e:
            print(f"⚠️ 讀取或寫入失敗 ID #{img_id} ({e})")
            fail_count += 1

    if batch:
        with conn.cursor() as b_cur:
            for b_data, b_mime, b_id in batch:
                b_cur.execute(
                    "UPDATE property_images SET image_data = %s, mime_type = %s WHERE id = %s;",
                    (b_data, b_mime, b_id)
                )

    print("\n--------------------------------------------------")
    print("🔍 驗證資料庫寫入結果...")
    cur.execute("SELECT COUNT(*), COUNT(image_data), SUM(OCTET_LENGTH(image_data)) FROM property_images;")
    v_total, v_has_data, v_bytes = cur.fetchone()
    print(f"總筆數: {v_total} 筆")
    print(f"成功填充 BYTEA: {v_has_data} 筆")
    print(f"寫入總位元組數: {v_bytes:,} bytes (約 {v_bytes / (1024*1024):.2f} MB)")
    print(f"成功率: {v_has_data / v_total * 100:.2f}%")
    print("--------------------------------------------------")

    conn.close()
    if v_has_data == v_total and v_total > 0:
        print("✅ 全數圖片已 100% 成功入庫至 PostgreSQL！")
        return True
    else:
        print("❌ 有部分圖片未成功寫入，請檢查！")
        return False

if __name__ == '__main__':
    migrate_images()
