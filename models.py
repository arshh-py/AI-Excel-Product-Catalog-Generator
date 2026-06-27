from dataclasses import dataclass


@dataclass(frozen=True)
class Product:
    id: int
    product_name: str
    price: str
    image_filename: str
    image_source_url: str
    title_hash: str
    created_at: str
    updated_at: str


PRODUCTS_SCHEMA = """
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_name TEXT NOT NULL,
    price TEXT,
    image_filename TEXT,
    image_source_url TEXT,
    title_hash TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
)
"""

IMAGE_CACHE_SCHEMA = """
CREATE TABLE IF NOT EXISTS image_cache (
    title_hash TEXT PRIMARY KEY,
    product_name TEXT NOT NULL,
    image_filename TEXT NOT NULL,
    image_source_url TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
)
"""

UPLOAD_BATCHES_SCHEMA = """
CREATE TABLE IF NOT EXISTS upload_batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    imported_count INTEGER NOT NULL DEFAULT 0,
    skipped_count INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL
)
"""

IMPORT_JOBS_SCHEMA = """
CREATE TABLE IF NOT EXISTS import_jobs (
    id TEXT PRIMARY KEY,
    original_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    status TEXT NOT NULL,
    total_count INTEGER NOT NULL DEFAULT 0,
    processed_count INTEGER NOT NULL DEFAULT 0,
    imported_count INTEGER NOT NULL DEFAULT 0,
    skipped_count INTEGER NOT NULL DEFAULT 0,
    current_product TEXT,
    error_message TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    completed_at TEXT
)
"""
