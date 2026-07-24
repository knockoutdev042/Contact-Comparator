"""
auth.py
-------------------------
Fixed-user-list authentication, restricted to an approved email domain.

There is no self-service signup and no external identity provider — an
administrator adds each employee ahead of time. Passwords are never
stored in plain text, only a bcrypt hash, kept in one of two places:

    - .streamlit/secrets.toml, under an [app_users] table (local dev,
      git-ignored)
    - the APP_USERS_JSON environment variable, a JSON object of
      {"email": "bcrypt_hash", ...} (production, e.g. a Render env var)

See scripts/hash_password.py to generate a hash for a new user, and
README.md for the full setup/troubleshooting guide.
"""

import json
import os

import bcrypt
import streamlit as st

ALLOWED_DOMAIN = "@lghomecomfort.ca"


def _load_users():
    users = {}

    try:
        secrets_users = st.secrets.get("app_users")
        if secrets_users:
            users.update(dict(secrets_users))
    except Exception:
        pass

    env_users = os.environ.get("APP_USERS_JSON")
    if env_users:
        try:
            users.update(json.loads(env_users))
        except (json.JSONDecodeError, TypeError):
            pass

    return users


def require_login():
    """Blocks the app behind a login form. Returns the signed-in user's email."""

    if st.session_state.get("authenticated"):
        return st.session_state["user_email"]

    users = _load_users()

    st.markdown(
        """
        <div style="text-align:center; margin-top:8vh; margin-bottom:1.5rem;">
            <div style="font-size:2.4rem;">🔒</div>
            <div style="font-size:1.4rem; font-weight:700; margin-top:0.25rem;">Sign in</div>
            <div style="color:#898781; font-size:0.9rem; margin-top:0.25rem;">
                Excluded Contacts Extractor — LG Home Comfort internal tool
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not users:
        _, center, _ = st.columns([1, 2, 1])
        with center:
            st.error("No users are configured on this deployment yet.")
        st.stop()

    _, center, _ = st.columns([1, 2, 1])
    with center:
        with st.form("login_form"):
            email = st.text_input("Email", placeholder="you@lghomecomfort.ca")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", use_container_width=True)

        if submitted:
            email_clean = email.strip().lower()

            if not email_clean.endswith(ALLOWED_DOMAIN):
                st.error(f"Access is restricted to {ALLOWED_DOMAIN} accounts.")
            elif email_clean not in users:
                st.error("Invalid email or password.")
            else:
                stored_hash = users[email_clean].encode("utf-8")
                if bcrypt.checkpw(password.encode("utf-8"), stored_hash):
                    st.session_state["authenticated"] = True
                    st.session_state["user_email"] = email_clean
                    st.rerun()
                else:
                    st.error("Invalid email or password.")

    st.stop()


def render_logout_button():
    with st.sidebar:
        st.markdown(
            f"**Signed in as**  \n{st.session_state.get('user_email', '')}"
        )
        if st.button("🚪 Log out", use_container_width=True):
            st.session_state.pop("authenticated", None)
            st.session_state.pop("user_email", None)
            st.rerun()
