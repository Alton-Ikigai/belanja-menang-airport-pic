"""Run locally: python generate_passwords.py. Never commit real passwords or hashes to GitHub."""
import hashlib
import secrets
import getpass

password=getpass.getpass('Enter new airport password: ')
if len(password)<12:
    raise SystemExit('Use at least 12 characters.')
salt=secrets.token_bytes(16)
iterations=310000
digest=hashlib.pbkdf2_hmac('sha256',password.encode(),salt,iterations)
print('Paste this hash in Streamlit Secrets for the selected airport:')
print(f'pbkdf2_sha256${iterations}${salt.hex()}${digest.hex()}')
