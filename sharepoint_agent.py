import streamlit as st
import msal
import requests
import os
import math

# ======================================================
# CONFIGURATION
# ======================================================
CLIENT_ID = "e8b31b4d-fc50-4ee4-8590-266ea7df9991"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite", "User.Read"]

CACHE_FILE = "msal_cache.bin"
CHUNK_SIZE = 320 * 1024  # 320 KB (Microsoft minimum)

# ======================================================
# TOKEN CACHE (PERSISTENT)
# ======================================================
def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

token_cache = load_cache()

msal_app = msal.PublicClientApplication(
    client_id=CLIENT_ID,
    authority=AUTHORITY,
    token_cache=token_cache
)

# ======================================================
# AUTHENTICATION
# ======================================================
def get_access_token():
    accounts = msal_app.get_accounts()
    if accounts:
        result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_cache(token_cache)
            return result["access_token"]

    flow = msal_app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("Azure rejected device code flow")
        st.json(flow)
        st.stop()

    st.info("One-time Microsoft login required")
    st.code(flow["message"])

    if st.button("I have completed login"):
        result = msal_app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            save_cache(token_cache)
            st.success("Authentication successful")
            st.rerun()
        else:
            st.error("Login failed")
            st.json(result)
            st.stop()

    st.stop()

# ======================================================
# ONEDRIVE HELPERS
# ======================================================
def graph_headers(token):
    return {"Authorization": f"Bearer {token}"}

def list_folders(token, parent_id="root"):
    if parent_id == "root":
        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{parent_id}/children"

    r = requests.get(url, headers=graph_headers(token))
    r.raise_for_status()

    return [i for i in r.json().get("value", []) if "folder" in i]

def create_folder(token, parent_id, name):
    if parent_id == "root":
        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{parent_id}/children"

    payload = {
        "name": name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }

    r = requests.post(url, headers=graph_headers(token), json=payload)
    r.raise_for_status()
    return r.json()["id"]

def simple_upload(token, file, folder_id):
    headers = graph_headers(token)
    headers["Content-Type"] = file.type

    if folder_id == "root":
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file.name}:/content"
    else:
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{file.name}:/content"

    return requests.put(url, headers=headers, data=file.getvalue())

def upload_large_file(token, file, folder_id):
    if folder_id == "root":
        session_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{file.name}:/createUploadSession"
    else:
        session_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}:/{file.name}:/createUploadSession"

    session = requests.post(
        session_url,
        headers=graph_headers(token),
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}
    ).json()

    upload_url = session["uploadUrl"]
    data = file.getvalue()
    size = len(data)
    chunks = math.ceil(size / CHUNK_SIZE)

    progress = st.progress(0.0)

    for i in range(chunks):
        start = i * CHUNK_SIZE
        end = min(start + CHUNK_SIZE, size)

        headers = {
            "Content-Length": str(end - start),
            "Content-Range": f"bytes {start}-{end - 1}/{size}"
        }

        r = requests.put(upload_url, headers=headers, data=data[start:end])
        if r.status_code not in (200, 201, 202):
            st.error("Chunk upload failed")
            st.json(r.json())
            st.stop()

        progress.progress((i + 1) / chunks)

    return r

# ======================================================
# STREAMLIT UI
# ======================================================
st.set_page_config(page_title="OneDrive Upload Agent", layout="centered")
st.title("OneDrive Upload Agent (Personal Account)")

token = get_access_token()
st.success("Authenticated (silent)")

# -----------------------------
# FOLDER BROWSER
# -----------------------------
st.subheader("Select Destination Folder")

if "current_folder" not in st.session_state:
    st.session_state.current_folder = "root"
    st.session_state.breadcrumb = []

folders = list_folders(token, st.session_state.current_folder)

col1, col2 = st.columns(2)

with col1:
    if st.session_state.breadcrumb and st.button("⬅ Go Back"):
        st.session_state.current_folder = st.session_state.breadcrumb.pop()
        st.rerun()

with col2:
    new_folder = st.text_input("Create new folder")
    if new_folder:
        new_id = create_folder(token, st.session_state.current_folder, new_folder)
        st.session_state.breadcrumb.append(st.session_state.current_folder)
        st.session_state.current_folder = new_id
        st.rerun()

folder_names = ["(Use this folder)"] + [f["name"] for f in folders]
choice = st.selectbox("Folders", folder_names)

if choice != "(Use this folder)":
    chosen = next(f for f in folders if f["name"] == choice)
    st.session_state.breadcrumb.append(st.session_state.current_folder)
    st.session_state.current_folder = chosen["id"]
    st.rerun()

# -----------------------------
# FILE UPLOAD
# -----------------------------
st.subheader("Upload File")

uploaded_file = st.file_uploader("Choose a file (any size)")

if uploaded_file and st.button("Upload to Selected Folder"):
    size_mb = uploaded_file.size / (1024 * 1024)
    st.info(f"File size: {size_mb:.2f} MB")

    if uploaded_file.size <= 4 * 1024 * 1024:
        resp = simple_upload(token, uploaded_file, st.session_state.current_folder)
    else:
        resp = upload_large_file(token, uploaded_file, st.session_state.current_folder)

    if resp.status_code in (200, 201, 202):
        st.success("File uploaded successfully to OneDrive")
    else:
        st.error("Upload failed")
        st.json(resp.json())
