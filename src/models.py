"""
Database models for IHADIS Data Comparison Tool.
"""
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Project(db.Model):
    """Project model - stores project metadata and column mappings."""
    __tablename__ = 'projects'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    source_type = db.Column(db.String(50))  # 'upload' or 'sheets'
    source_ref = db.Column(db.String(500))  # filename or sheet_id
    excel_path = db.Column(db.String(500))  # local Excel file path
    sheet_name = db.Column(db.String(255))  # Selected Excel sheet name

    # Column mappings (user-selected)
    col_primary_text = db.Column(db.String(100))
    col_secondary_text = db.Column(db.String(100))
    col_arabic_text = db.Column(db.String(100))
    col_id = db.Column(db.String(100))
    col_ratio = db.Column(db.String(100))

    # Project settings
    rows_per_chunk = db.Column(db.Integer, default=500)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Relationships
    comments = db.relationship('Comment', backref='project', lazy='dynamic', cascade='all, delete-orphan')
    approvals = db.relationship('ApprovalStatus', backref='project', lazy='dynamic', cascade='all, delete-orphan')

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'source_type': self.source_type,
            'source_ref': self.source_ref,
            'excel_path': self.excel_path,
            'sheet_name': self.sheet_name,
            'col_primary_text': self.col_primary_text,
            'col_secondary_text': self.col_secondary_text,
            'col_arabic_text': self.col_arabic_text,
            'col_id': self.col_id,
            'col_ratio': self.col_ratio,
            'rows_per_chunk': self.rows_per_chunk,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }


class Comment(db.Model):
    """Comment model - stores row comments (migrated from Excel)."""
    __tablename__ = 'comments'

    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    row_id = db.Column(db.Integer, nullable=False)
    text = db.Column(db.Text)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint('project_id', 'row_id', name='unique_project_row_comment'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'row_id': self.row_id,
            'text': self.text,
            'created_at': self.created_at.isoformat() if self.created_at else None,
        }


class ApprovalStatus(db.Model):
    """Approval status model - stores cell approval states (migrated from Excel colors)."""
    __tablename__ = 'approvals'

    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('projects.id'), nullable=False)
    row_id = db.Column(db.Integer, nullable=False)
    column = db.Column(db.String(50), nullable=False)  # 'primary' or 'secondary'
    status = db.Column(db.String(20), default='pending')  # 'approved', 'rejected', 'pending'
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (
        db.UniqueConstraint('project_id', 'row_id', 'column', name='unique_project_row_column'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'row_id': self.row_id,
            'column': self.column,
            'status': self.status,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }


class ApiKey(db.Model):
    """API Key model - stores user-configurable AI provider credentials."""
    __tablename__ = 'api_keys'

    id = db.Column(db.Integer, primary_key=True)
    provider = db.Column(db.String(50), unique=True, nullable=False)  # 'google', 'claude', 'openai', 'deepseek', 'grok'
    api_key_encrypted = db.Column(db.Text, nullable=False)
    model_name = db.Column(db.String(100))
    max_tokens = db.Column(db.Integer, default=4096)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self, include_key=False):
        result = {
            'id': self.id,
            'provider': self.provider,
            'model_name': self.model_name,
            'max_tokens': self.max_tokens,
            'is_active': self.is_active,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }
        if include_key:
            result['api_key_encrypted'] = self.api_key_encrypted
        else:
            # Return masked key for display
            result['api_key_masked'] = '***' + self.api_key_encrypted[-4:] if self.api_key_encrypted else None
        return result


class Settings(db.Model):
    """Settings model - key-value store for global processing settings."""
    __tablename__ = 'settings'

    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(100), unique=True, nullable=False)
    value = db.Column(db.Text)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # Default settings
    DEFAULTS = {
        'batch_size': '5',
        'max_retries': '3',
        'retry_delay': '0',
        'save_interval': '5',
    }

    @classmethod
    def get(cls, key, default=None):
        """Get a setting value by key."""
        setting = cls.query.filter_by(key=key).first()
        if setting:
            return setting.value
        return default or cls.DEFAULTS.get(key)

    @classmethod
    def set(cls, key, value):
        """Set a setting value."""
        setting = cls.query.filter_by(key=key).first()
        if setting:
            setting.value = str(value)
        else:
            setting = cls(key=key, value=str(value))
            db.session.add(setting)
        db.session.commit()
        return setting

    @classmethod
    def get_all(cls):
        """Get all settings as a dictionary."""
        settings = cls.query.all()
        result = dict(cls.DEFAULTS)  # Start with defaults
        for setting in settings:
            result[setting.key] = setting.value
        return result

    def to_dict(self):
        return {
            'id': self.id,
            'key': self.key,
            'value': self.value,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }
