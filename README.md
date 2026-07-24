# Excluded Contacts Extractor

A Streamlit web app that finds Contact Us entries not present in Salesforce.
A contact is considered matched if its Email is found in Salesforce, or —
when no email is given — if its Phone is found instead.

## Running locally

```
pip install -r requirements.txt
streamlit run app.py
```

Add at least one user (see below) before signing in.

## Authentication (fixed user list, restricted to `@lghomecomfort.ca`)

There's no self-service signup and no external identity provider — an
administrator adds each employee to a fixed list ahead of time. This was
chosen over Microsoft/Google SSO to avoid depending on the company's
identity-provider admin access, and over a database-backed signup flow
because Render's free tier has an ephemeral filesystem (any file written
while running gets wiped on restart/redeploy — a real database would be
needed for genuine self-service signup, which felt like overkill here).

How it works: passwords are never stored in plain text — only a bcrypt hash
of each password is kept, in one of two places:

- `.streamlit/secrets.toml`, under an `[app_users]` table — convenient for
  local dev, and it's git-ignored so it's never committed
- the `APP_USERS_JSON` environment variable, a JSON object of
  `{"email": "bcrypt_hash", ...}` — used in production (e.g. as a Render
  environment variable)

Both a domain check (`@lghomecomfort.ca` only, even if a non-matching email
were somehow added to the list) and a bcrypt password check run before
anyone gets in. There's no "remember me": since there's no persistent
session cookie, signing in again is needed after a hard browser refresh —
an accepted simplification for this internal tool.

### Adding a user

1. Generate a hash for their password:

   ```
   python scripts/hash_password.py
   ```

   This prompts for the password (input is hidden) and prints a bcrypt hash
   like `$2b$12$...`.

2. **Local development** — add it to `.streamlit/secrets.toml` (git-ignored;
   copy `.streamlit/secrets.toml.example` to get started):

   ```toml
   [app_users]
   "jane@lghomecomfort.ca" = "$2b$12$...paste the hash here..."
   "john@lghomecomfort.ca" = "$2b$12$...another user's hash..."
   ```

3. **Render deployment** — set the `APP_USERS_JSON` environment variable
   (Dashboard > your service > Environment) to a JSON object with one entry
   per user:

   ```json
   {"jane@lghomecomfort.ca": "$2b$12$...", "john@lghomecomfort.ca": "$2b$12$..."}
   ```

   Then redeploy (or Render will pick up the env var change automatically,
   depending on your settings).

### Removing a user / resetting a password

Delete their entry (or replace it with a freshly generated hash) from
`.streamlit/secrets.toml` or `APP_USERS_JSON`, then redeploy. There's no
self-service "forgot password" flow — an administrator has to do this.

### Troubleshooting

- **"No users are configured on this deployment yet"**: neither
  `[app_users]` in secrets.toml nor `APP_USERS_JSON` has any entries — add
  at least one user as above.
- **"Access is restricted to @lghomecomfort.ca accounts"**: the email
  entered doesn't end in `@lghomecomfort.ca` — this is checked even if that
  exact email happens to be in the user list, as a safety net.
- **"Invalid email or password"**: either the email isn't in the list, or
  the password doesn't match its stored hash — regenerate the hash with
  `scripts/hash_password.py` if unsure.

## Deploying to Render

This repo includes `render.yaml`. Create a new Web Service on Render
pointing at this branch/repo and it will pick up the build/start commands
automatically. Set `APP_USERS_JSON` in the service's Environment settings
before the first deploy (or the app will show "No users are configured").
