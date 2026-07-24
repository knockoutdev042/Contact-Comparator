# 📇 Excluded Contacts Extractor

<p>
  <img alt="Python" src="https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white">
  <img alt="Streamlit" src="https://img.shields.io/badge/Streamlit-app-FF4B4B?logo=streamlit&logoColor=white">
  <img alt="Auth" src="https://img.shields.io/badge/Auth-bcrypt%20%2B%20fixed%20user%20list-2a78d6?logo=letsencrypt&logoColor=white">
  <img alt="Deploy" src="https://img.shields.io/badge/Deploy-Render-46E3B7?logo=render&logoColor=white">
  <img alt="Status" src="https://img.shields.io/badge/Status-internal%20tool-e94560">
</p>

A Streamlit web app that finds Contact Us entries **not** present in Salesforce.
A contact counts as matched if its **Email** is found in Salesforce, or —
when no email is given — if its **Phone** is found instead.

---

## 📑 Contents

- [🚀 Running locally](#-running-locally)
- [🔐 Authentication](#-authentication-fixed-user-list-restricted-to-lghomecomfortca)
- [☁️ Deploying to Render](#️-deploying-to-render)

---

## 🚀 Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

> [!IMPORTANT]
> You need at least one user configured before you can sign in — see
> [Adding a user](#-adding-a-user) below.

---

## 🔐 Authentication (fixed user list, restricted to `@lghomecomfort.ca`)

There's no self-service signup and no external identity provider — an
🧑‍💼 **administrator adds each employee to a fixed list ahead of time.**

| Why not... | Because... |
|---|---|
| 🔑 Microsoft / Google SSO | avoids depending on the company's identity-provider admin access |
| 🗄️ Database-backed signup | Render's free tier has an **ephemeral filesystem** — anything written while running is wiped on restart/redeploy. A real database would be needed for genuine self-service signup, which felt like overkill here |

> [!NOTE]
> Passwords are **never stored in plain text** — only a 🔒 bcrypt hash of
> each password is kept, in one of two places:
>
> - `.streamlit/secrets.toml`, under an `[app_users]` table — convenient for
>   local dev, and it's **git-ignored** so it's never committed
> - the `APP_USERS_JSON` environment variable, a JSON object of
>   `{"email": "bcrypt_hash", ...}` — used in production (e.g. a Render
>   environment variable)

Both a ✅ **domain check** (`@lghomecomfort.ca` only, even if a non-matching
email were somehow added to the list) and a ✅ **bcrypt password check** run
before anyone gets in.

> [!WARNING]
> There's no "remember me". Since there's no persistent session cookie,
> signing in again is needed after a hard browser refresh — an accepted
> simplification for this internal tool.

### ➕ Adding a user

**1.** Generate a hash for their password:

```bash
python scripts/hash_password.py
```

This prompts for the password (input is hidden 👀) and prints a bcrypt hash
like `$2b$12$...`.

**2. Local development** — add it to `.streamlit/secrets.toml` (git-ignored;
copy `.streamlit/secrets.toml.example` to get started):

```toml
[app_users]
"jane@lghomecomfort.ca" = "$2b$12$...paste the hash here..."
"john@lghomecomfort.ca" = "$2b$12$...another user's hash..."
```

**3. Render deployment** — set the `APP_USERS_JSON` environment variable
(**Dashboard → your service → Environment**) to a JSON object with one entry
per user:

```json
{"jane@lghomecomfort.ca": "$2b$12$...", "john@lghomecomfort.ca": "$2b$12$..."}
```

Then redeploy (or Render will pick it up automatically, depending on your
settings). 🔄

### ➖ Removing a user / resetting a password

Delete their entry (or replace it with a freshly generated hash) from
`.streamlit/secrets.toml` or `APP_USERS_JSON`, then redeploy.

> [!TIP]
> There's no self-service "forgot password" flow — an administrator has to
> do this.

### 🛠️ Troubleshooting

<details>
<summary><strong>❌ "No users are configured on this deployment yet"</strong></summary>

Neither `[app_users]` in `secrets.toml` nor `APP_USERS_JSON` has any
entries — add at least one user as described above.

</details>

<details>
<summary><strong>🚫 "Access is restricted to @lghomecomfort.ca accounts"</strong></summary>

The email entered doesn't end in `@lghomecomfort.ca` — this is checked even
if that exact email happens to be in the user list, as a safety net.

</details>

<details>
<summary><strong>🔑 "Invalid email or password"</strong></summary>

Either the email isn't in the list, or the password doesn't match its
stored hash — regenerate the hash with `scripts/hash_password.py` if
unsure.

</details>

---

## ☁️ Deploying to Render

This repo includes `render.yaml`. Create a new Web Service on Render
pointing at this branch/repo and it will pick up the build/start commands
automatically.

> [!CAUTION]
> Set `APP_USERS_JSON` in the service's **Environment** settings *before*
> the first deploy — otherwise the app will show "No users are configured".
