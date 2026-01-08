"""
Database setup and utilities for IHADIS Data Comparison Tool.
"""
import os
from pathlib import Path
from cryptography.fernet import Fernet
from flask import Flask

from .models import db, Settings


def get_encryption_key():
    """Get or generate encryption key for API keys."""
    key = os.environ.get('API_KEY_ENCRYPTION_SECRET')
    if not key:
        # Generate a new key if not set (for development)
        key = Fernet.generate_key().decode()
        print(f"Warning: Using auto-generated encryption key. Set API_KEY_ENCRYPTION_SECRET in .env for production.")
    return key.encode() if isinstance(key, str) else key


def encrypt_api_key(api_key: str) -> str:
    """Encrypt an API key for storage."""
    if not api_key:
        return ''
    fernet = Fernet(get_encryption_key())
    return fernet.encrypt(api_key.encode()).decode()


def decrypt_api_key(encrypted_key: str) -> str:
    """Decrypt an API key from storage."""
    if not encrypted_key:
        return ''
    try:
        fernet = Fernet(get_encryption_key())
        return fernet.decrypt(encrypted_key.encode()).decode()
    except Exception:
        # If decryption fails (wrong key), return empty string
        return ''


def init_db(app: Flask):
    """Initialize the database with the Flask app."""
    # Configure database
    database_url = os.environ.get('DATABASE_URL', 'sqlite:///data/app.db')

    # Fix for Render/Heroku PostgreSQL URL (postgres:// -> postgresql://)
    if database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql://', 1)

    # Ensure data directory exists for SQLite
    if database_url.startswith('sqlite:///'):
        db_path = database_url.replace('sqlite:///', '')
        # Handle relative paths - make them relative to app root
        if not os.path.isabs(db_path):
            app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            db_path = os.path.join(app_root, db_path)
            database_url = f'sqlite:///{db_path}'
        db_dir = os.path.dirname(db_path)
        if db_dir:
            Path(db_dir).mkdir(parents=True, exist_ok=True)
            print(f"Database directory created: {db_dir}")

    app.config['SQLALCHEMY_DATABASE_URI'] = database_url
    print(f"Database URL: {database_url}")
    app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

    # Initialize SQLAlchemy with app
    db.init_app(app)

    # Create tables
    with app.app_context():
        db.create_all()

        # Initialize default settings if not exists
        _init_default_settings()

    return db


def _init_default_settings():
    """Initialize default settings in the database."""
    for key, value in Settings.DEFAULTS.items():
        existing = Settings.query.filter_by(key=key).first()
        if not existing:
            setting = Settings(key=key, value=value)
            db.session.add(setting)

    try:
        db.session.commit()
    except Exception:
        db.session.rollback()


def get_db():
    """Get the database instance."""
    return db
