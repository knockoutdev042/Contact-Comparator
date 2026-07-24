"""
scripts/hash_password.py
-------------------------
Generate a bcrypt hash for a new app user's password.

Run:
    python scripts/hash_password.py

Paste the printed hash into .streamlit/secrets.toml (under [app_users])
for local dev, or into the APP_USERS_JSON environment variable for the
Render deployment. See README.md for the full walkthrough.
"""

import getpass

import bcrypt


def main():
    password = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm password: ")

    if not password:
        print("Password cannot be empty.")
        raise SystemExit(1)

    if password != confirm:
        print("Passwords do not match.")
        raise SystemExit(1)

    hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt())
    print("\nBcrypt hash (copy this):\n")
    print(hashed.decode("utf-8"))


if __name__ == "__main__":
    main()
