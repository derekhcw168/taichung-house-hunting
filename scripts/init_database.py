import psycopg
import sys

def init_db():
    print("Connecting to PostgreSQL server...")
    admin_conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=postgres", autocommit=True)
    with admin_conn.cursor() as cur:
        cur.execute("SELECT 1 FROM pg_database WHERE datname = 'taichung_house'")
        if not cur.fetchone():
            cur.execute('CREATE DATABASE taichung_house ENCODING "UTF8";')
            print("Successfully created database 'taichung_house'.")
        else:
            print("Database 'taichung_house' already exists.")
    admin_conn.close()

    print("Connecting to 'taichung_house' database...")
    conn = psycopg.connect("host=localhost port=5432 user=postgres password=postgres dbname=taichung_house", autocommit=True)
    with conn.cursor() as cur:
        # Create tables
        cur.execute("""
        CREATE TABLE IF NOT EXISTS communities (
            id SERIAL PRIMARY KEY,
            name VARCHAR(100) UNIQUE NOT NULL,
            district VARCHAR(50),
            address VARCHAR(255),
            total_units INTEGER,
            builder VARCHAR(100),
            completion_year INTEGER,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS community_earthquake_records (
            id SERIAL PRIMARY KEY,
            community_id INTEGER REFERENCES communities(id) ON DELETE SET NULL,
            original_community_name VARCHAR(150) NOT NULL,
            district VARCHAR(50),
            damage_level VARCHAR(50),
            legal_status TEXT,
            repair_status TEXT,
            rebuilt_name VARCHAR(100),
            notes TEXT,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS properties (
            id SERIAL PRIMARY KEY,
            code VARCHAR(50) UNIQUE,
            community_name VARCHAR(100) NOT NULL,
            community_id INTEGER REFERENCES communities(id) ON DELETE SET NULL,
            city VARCHAR(50) DEFAULT '臺中市',
            district VARCHAR(50),
            address VARCHAR(255),
            title VARCHAR(255) NOT NULL,
            price_total NUMERIC(10, 2),
            unit_price NUMERIC(10, 2),
            rooms INTEGER,
            living_rooms INTEGER,
            bathrooms INTEGER,
            balconies INTEGER,
            layout_raw VARCHAR(50),
            floor_info VARCHAR(50),
            current_floor INTEGER,
            total_floors INTEGER,
            age NUMERIC(5, 1),
            total_area NUMERIC(10, 2),
            indoor_area NUMERIC(10, 2),
            attached_area NUMERIC(10, 2),
            indoor_total NUMERIC(10, 2),
            public_area NUMERIC(10, 2),
            land_area NUMERIC(10, 2),
            public_ratio NUMERIC(5, 2),
            public_ratio_has_parking BOOLEAN DEFAULT FALSE,
            public_ratio_desc TEXT,
            parking_type VARCHAR(50),
            parking_desc TEXT,
            orientation VARCHAR(50),
            management_fee VARCHAR(50),
            decision_status VARCHAR(50) DEFAULT 'pending',
            original_source_platform VARCHAR(50),
            original_file_name VARCHAR(255),
            original_url TEXT,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS property_sources (
            id SERIAL PRIMARY KEY,
            property_id INTEGER REFERENCES properties(id) ON DELETE CASCADE,
            platform VARCHAR(50) NOT NULL,
            original_case_id VARCHAR(100),
            original_title VARCHAR(255),
            original_url TEXT,
            local_file_relpath VARCHAR(500),
            agent_name VARCHAR(100),
            agent_phone VARCHAR(50),
            agent_store VARCHAR(150),
            raw_metadata JSONB,
            description_text TEXT,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS property_images (
            id SERIAL PRIMARY KEY,
            property_id INTEGER REFERENCES properties(id) ON DELETE CASCADE,
            image_category VARCHAR(50),
            file_name VARCHAR(255) NOT NULL,
            file_relpath VARCHAR(500) NOT NULL,
            file_size INTEGER,
            width INTEGER,
            height INTEGER,
            original_url TEXT,
            sort_order INTEGER DEFAULT 0,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS inspection_notes (
            id SERIAL PRIMARY KEY,
            property_id INTEGER REFERENCES properties(id) ON DELETE SET NULL,
            community_name VARCHAR(100),
            realtor_company VARCHAR(50),
            visit_date DATE,
            visit_time VARCHAR(20),
            status_conclusion VARCHAR(50),
            pros TEXT,
            cons TEXT,
            suggested_offer_low NUMERIC(10, 2),
            suggested_offer_high NUMERIC(10, 2),
            renovation_estimate TEXT,
            google_maps_url TEXT,
            notes TEXT,
            created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS buyer_preferences (
            id SERIAL PRIMARY KEY,
            category VARCHAR(50),
            item_name VARCHAR(100) NOT NULL,
            is_mandatory BOOLEAN DEFAULT FALSE,
            priority INTEGER DEFAULT 1,
            condition_detail TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_properties_community ON properties(community_name);
        CREATE INDEX IF NOT EXISTS idx_properties_status ON properties(decision_status);
        CREATE INDEX IF NOT EXISTS idx_properties_price ON properties(price_total);
        CREATE INDEX IF NOT EXISTS idx_properties_district ON properties(district);
        CREATE INDEX IF NOT EXISTS idx_property_sources_pid ON property_sources(property_id);
        CREATE INDEX IF NOT EXISTS idx_property_images_pid ON property_images(property_id);
        CREATE INDEX IF NOT EXISTS idx_property_sources_meta_gin ON property_sources USING gin (raw_metadata);
        """)
        print("Schema and indexes successfully initialized!")
    conn.close()

if __name__ == "__main__":
    init_db()
