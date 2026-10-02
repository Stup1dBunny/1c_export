"""Explicit desktop OAuth connection; nightly runs only refresh existing consent."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from report_tools import write_json

SCOPES = ['https://www.googleapis.com/auth/drive']


def validate_client(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    client = data.get('installed')
    if not isinstance(client, dict):
        raise ValueError('Create a Desktop app OAuth client in Google Cloud')
    if (client.get('auth_uri') not in {'https://accounts.google.com/o/oauth2/auth',
                                     'https://accounts.google.com/o/oauth2/v2/auth'}
            or client.get('token_uri') != 'https://oauth2.googleapis.com/token'
            or not client.get('client_id', '').endswith('.apps.googleusercontent.com')):
        raise ValueError('The OAuth client does not use the expected Google endpoints')
    return data


def load_google_credentials(token_path):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google.auth.exceptions import RefreshError

    token_path = Path(token_path)
    if not token_path.is_file():
        raise PermissionError('Google Drive API is not connected; run the one-time OAuth setup')
    data = json.loads(token_path.read_text(encoding='utf-8'))
    if data.get('token_uri') != 'https://oauth2.googleapis.com/token':
        raise ValueError('Unexpected OAuth token endpoint')
    credentials = Credentials.from_authorized_user_info(data, SCOPES)
    if not credentials.refresh_token:
        raise PermissionError('Google connection lacks offline access; reconnect explicitly')
    if not credentials.valid:
        try:
            credentials.refresh(Request())
        except RefreshError:
            raise PermissionError('Google revoked or expired the API connection; reconnect explicitly') from None
        write_json(token_path, json.loads(credentials.to_json()))
    return credentials


def build_service(token_path):
    from googleapiclient.discovery import build
    return build('drive', 'v3', credentials=load_google_credentials(token_path), cache_discovery=False)


def connect(client_path, token_path, folder_id):
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    client = validate_client(client_path)
    if Path(token_path).exists():
        raise FileExistsError('A Google token already exists; inspect the connection before replacing it')
    print('Google will request Drive access so the program can use the existing report folder and files.')
    print('Approve only in your ordinary browser. No Google password is read by this program.', flush=True)
    flow = InstalledAppFlow.from_client_config(client, SCOPES, autogenerate_code_verifier=True)
    credentials = flow.run_local_server(
        host='127.0.0.1', port=0, open_browser=True, timeout_seconds=600,
        authorization_prompt_message='Complete the Google consent in the browser window.',
        success_message='Google Drive connection received. You can close this tab.',
        access_type='offline', prompt='consent')
    if not credentials.refresh_token:
        raise PermissionError('Google did not issue offline access; no token was saved')
    # Exercise the actual unattended refresh before recording a working connection.
    credentials.refresh(Request())
    service = build('drive', 'v3', credentials=credentials, cache_discovery=False)
    folder = service.files().get(fileId=folder_id, supportsAllDrives=True,
        fields='id,name,mimeType,trashed,capabilities(canAddChildren)').execute(num_retries=3)
    if (folder.get('trashed') or folder.get('mimeType') != 'application/vnd.google-apps.folder'
            or not folder.get('capabilities', {}).get('canAddChildren')):
        raise PermissionError('The authorized account cannot add reports to the configured folder')
    write_json(Path(token_path), json.loads(credentials.to_json()))
    print('GOOGLE_DRIVE_API_CONNECTED: configured report folder verified', flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--client', type=Path, required=True)
    parser.add_argument('--token', type=Path, required=True)
    parser.add_argument('--folder', required=True)
    args = parser.parse_args()
    connect(args.client, args.token, args.folder)
