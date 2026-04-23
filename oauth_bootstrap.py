from __future__ import annotations

import os
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/spreadsheets",
]


def main() -> None:
    client_secret_file = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET_FILE", "").strip()
    if not client_secret_file:
        raise RuntimeError("Defina GOOGLE_OAUTH_CLIENT_SECRET_FILE apontando para o JSON do OAuth client desktop/web.")

    path = Path(client_secret_file).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {path}")

    flow = InstalledAppFlow.from_client_secrets_file(str(path), SCOPES)
    creds = flow.run_local_server(port=0)

    print("\n=== COPIE ESTES VALORES PARA O .env ===\n")
    print(f"GOOGLE_CLIENT_ID={creds.client_id}")
    print(f"GOOGLE_CLIENT_SECRET={creds.client_secret}")
    print(f"GOOGLE_REFRESH_TOKEN={creds.refresh_token}")
    print("GOOGLE_TOKEN_URI=https://oauth2.googleapis.com/token")


if __name__ == "__main__":
    main()
