# Belanja & Menang — Airport PIC Portal

This is a **separate** Streamlit app. Do not replace or edit your existing management dashboard `app.py`.

## Deploy

1. Upload `app.py`, `requirements.txt`, `generate_passwords.py`, `.gitignore`, and this README to the **new** GitHub repository.
2. Create a new app at https://share.streamlit.io/ using the new repository and `app.py`.
3. In the **new app's** Streamlit Secrets, configure:

```toml
SUPABASE_URL = "https://YOUR-PROJECT.supabase.co"
SUPABASE_READ_KEY = "YOUR_READ_ONLY_KEY"

[airport_password_hashes]
PEN = "pbkdf2_sha256$310000$SALT_HEX$HASH_HEX"
LGK = "pbkdf2_sha256$310000$SALT_HEX$HASH_HEX"
# Continue for all 18 airport codes
```

4. Run `python generate_passwords.py` **locally** once for each airport, and paste the hash into Streamlit Secrets. Do not upload actual passwords or the secrets file to GitHub.
5. The Supabase read key must be authorized to **SELECT** the `campaign_dashboard` row `id=1`. Never use the Supabase service-role key in a public client or repository. A backend-only Streamlit secret is not exposed in browser code, but least-privilege credentials are still strongly recommended.
6. Test with at least two airport accounts and verify each sees only its own records and Excel export.

## Data / permissions

- Reads `campaign_dashboard.dashboard_data` from the **existing** Supabase project, once per 60 seconds. It never writes to Supabase.
- Data is filtered to the authenticated airport **on the server** before display or Excel generation. Database-level per-airport row policies cannot be applied directly to this existing single-row JSON snapshot; a dedicated airport-scoped table/API is recommended for stronger production isolation.
- The login uses PBKDF2 password hashes. A simple per-session failed-attempt delay is included; production use should add durable, account-level rate limits, password rotation and auditing.
- Merchant-level monthly categories are available when the existing publication includes `monthly_merchant_locations`.
- The current published snapshot contains aggregated merchant entries but **not merchant-level contest sales**, so this portal does not invent those figures.
- The airport sales target is annual, distinct from campaign contest sales. Actual airport-wide sales per airport are not derived from contest receipts.
- Target figures are copied from the last supplied target schedule; update this new project's TARGETS mapping if targets change.
- The existing management app is **not modified**. If the published snapshot structure changes, the PIC portal may need a compatible reader update.
