import os
import hashlib
import base64
import json
from pathlib import Path
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend

PIN_FILE = Path(__file__).parent / ".pin_hash"


def hash_pin(pin: str) -> str:
    return hashlib.sha256(pin.encode()).hexdigest()


def set_pin(pin: str) -> None:
    hashed = hash_pin(pin)
    PIN_FILE.write_text(hashed)


def has_pin_set() -> bool:
    return PIN_FILE.exists()


def verify_pin(pin: str) -> bool:
    if not PIN_FILE.exists():
        return True
    stored_hash = PIN_FILE.read_text()
    return hash_pin(pin) == stored_hash


def derive_aes_key(pin: str) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=b'academic_match_salt_v1',
        iterations=100000,
        backend=default_backend()
    )
    return kdf.derive(pin.encode('utf-8'))


def encrypt_with_pin(data: str, pin: str) -> str:
    aes_key = derive_aes_key(pin)
    iv = os.urandom(16)
    cipher = Cipher(algorithms.AES(aes_key), modes.CBC(iv), backend=default_backend())
    encryptor = cipher.encryptor()
    
    padded_data = data + ' ' * (16 - len(data) % 16)
    ciphertext = encryptor.update(padded_data.encode('utf-8')) + encryptor.finalize()
    
    return base64.b64encode(iv + ciphertext).decode('utf-8')


def decrypt_with_pin(encrypted_data: str, pin: str) -> str:
    try:
        aes_key = derive_aes_key(pin)
        data = base64.b64decode(encrypted_data.encode('utf-8'))
        iv = data[:16]
        ciphertext = data[16:]
        
        cipher = Cipher(algorithms.AES(aes_key), modes.CBC(iv), backend=default_backend())
        decryptor = cipher.decryptor()
        padded_plaintext = decryptor.update(ciphertext) + decryptor.finalize()
        
        return padded_plaintext.decode('utf-8').rstrip()
    except Exception:
        return None
