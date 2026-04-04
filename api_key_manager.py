import os
import json
import uuid
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List
from crypto_utils import encrypt_with_pin, decrypt_with_pin, has_pin_set, set_pin, verify_pin

KEYS_BASE_DIR = Path(__file__).parent / "api_keys"
KEYS_BASE_DIR.mkdir(exist_ok=True)


class APIKeyManager:
    def __init__(self):
        pass
    
    def create_project(self, name: str) -> str:
        project_id = str(uuid.uuid4())[:8]
        
        meta = {
            'id': project_id,
            'name': name,
            'created': datetime.now().isoformat(),
            'keys': {}
        }
        self._save_project_file(project_id, meta)
        return project_id
    
    def get_projects(self) -> List[Dict]:
        projects = []
        for f in KEYS_BASE_DIR.iterdir():
            if f.is_file() and f.suffix == '.json':
                meta = self._load_project_file(f.stem)
                if meta:
                    projects.append(meta)
        return sorted(projects, key=lambda x: x['created'], reverse=True)
    
    def delete_project(self, project_id: str):
        project_file = KEYS_BASE_DIR / f"{project_id}.json"
        if project_file.exists():
            project_file.unlink()
    
    def get_project_keys(self, project_id: str, pin: str = None) -> Dict[str, Optional[str]]:
        meta = self._load_project_file(project_id)
        if not meta:
            return {}
        
        keys = {}
        for provider, encrypted in meta.get('keys', {}).items():
            if not encrypted:
                keys[provider] = None
                continue
            if pin:
                keys[provider] = decrypt_with_pin(encrypted, pin)
            else:
                keys[provider] = encrypted
        return keys
    
    def set_api_key(self, project_id: str, provider: str, api_key: str, pin: str = None):
        meta = self._load_project_file(project_id)
        if not meta:
            meta = self._create_default_meta(project_id)
        
        if pin:
            encrypted = encrypt_with_pin(api_key, pin)
        else:
            encrypted = api_key
        meta['keys'][provider] = encrypted
        self._save_project_file(project_id, meta)
    
    def remove_api_key(self, project_id: str, provider: str):
        meta = self._load_project_file(project_id)
        if meta and provider in meta.get('keys', {}):
            del meta['keys'][provider]
            self._save_project_file(project_id, meta)
    
    def has_keys_configured(self, project_id: str) -> bool:
        meta = self._load_project_file(project_id)
        return bool(meta and meta.get('keys'))
    
    def export_project(self, project_id: str, destination_path: str) -> bool:
        meta = self._load_project_file(project_id)
        if not meta:
            return False
        try:
            with open(destination_path, 'w', encoding='utf-8') as f:
                json.dump(meta, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False
    
    def import_project(self, source_path: str) -> Optional[str]:
        try:
            with open(source_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            
            if 'id' not in meta or 'name' not in meta:
                return None
            
            existing = self._load_project_file(meta['id'])
            if existing:
                meta['id'] = str(uuid.uuid4())[:8]
            
            self._save_project_file(meta['id'], meta)
            return meta['id']
        except Exception:
            return None
    
    def get_project_info(self, project_id: str) -> Optional[Dict]:
        return self._load_project_file(project_id)
    
    def _create_default_meta(self, project_id: str) -> Dict:
        return {
            'id': project_id,
            'name': f'Project {project_id}',
            'created': datetime.now().isoformat(),
            'keys': {}
        }
    
    def _load_project_file(self, project_id: str) -> Optional[Dict]:
        project_file = KEYS_BASE_DIR / f"{project_id}.json"
        if not project_file.exists():
            return None
        try:
            with open(project_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None
    
    def _save_project_file(self, project_id: str, meta: dict):
        project_file = KEYS_BASE_DIR / f"{project_id}.json"
        with open(project_file, 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)


class PINManager:
    @staticmethod
    def is_pin_required() -> bool:
        return has_pin_set()
    
    @staticmethod
    def setup_pin(pin: str):
        if len(pin) != 4:
            raise ValueError("PIN must be exactly 4 characters")
        if not pin.isdigit():
            raise ValueError("PIN must be only digits")
        set_pin(pin)
    
    @staticmethod
    def validate_pin(pin: str) -> bool:
        return verify_pin(pin)
    
    @staticmethod
    def change_pin(old_pin: str, new_pin: str) -> bool:
        if not verify_pin(old_pin):
            return False
        if len(new_pin) != 4 or not new_pin.isdigit():
            raise ValueError("New PIN must be exactly 4 digits")
        set_pin(new_pin)
        return True
