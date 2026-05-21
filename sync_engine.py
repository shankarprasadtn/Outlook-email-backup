import os
import sys
import logging
import datetime
import shutil
import msal
import requests
from urllib.parse import urlparse
try:
    import pythoncom
    import win32timezone
except ImportError:
    pythoncom = None
    win32timezone = None

# Set up logging to file
LOG_FILE = "sync.log"

def get_log_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, LOG_FILE)

# Custom log handler to forward logs to the GUI
class GUIHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        log_entry = self.format(record)
        if self.callback:
            self.callback(log_entry, record.levelno)

# Initialize logger
logger = logging.getLogger("OutlookSharePointSync")
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(get_log_path(), encoding='utf-8')
file_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', '%Y-%m-%d %H:%M:%S')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

def setup_gui_logging(callback):
    # Remove existing gui handlers if any
    for h in logger.handlers[:]:
        if isinstance(h, GUIHandler):
            logger.removeHandler(h)
    
    gui_handler = GUIHandler(callback)
    gui_formatter = logging.Formatter('%(asctime)s - %(message)s', '%H:%M:%S')
    gui_handler.setFormatter(gui_formatter)
    logger.addHandler(gui_handler)

def get_token_cache_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "token_cache.bin")

# MSAL Token cache management
def load_cache():
    cache = msal.SerializableTokenCache()
    cache_path = get_token_cache_path()
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "rb") as f:
                cache.deserialize(f.read())
        except Exception as e:
            logger.warning(f"Could not load token cache: {e}")
    return cache

def save_cache(cache):
    cache_path = get_token_cache_path()
    try:
        with open(cache_path, "wb") as f:
            f.write(cache.serialize())
    except Exception as e:
        logger.warning(f"Could not save token cache: {e}")

# Acquire Access Token for Graph API
def acquire_token(config, login_hint=None):
    client_id = config.get("client_id")
    tenant_id = config.get("tenant_id", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    
    # We require Mail.ReadWrite, Files.ReadWrite.All, Sites.ReadWrite.All
    scopes = ["Mail.ReadWrite", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
    
    cache = load_cache()
    app = msal.PublicClientApplication(
        client_id, 
        authority=authority,
        token_cache=cache
    )
    
    accounts = app.get_accounts()
    result = None
    
    if accounts:
        # Try silently first
        logger.info("Found cached account. Attempting silent token acquisition...")
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result:
            logger.info("Silent token acquisition successful.")
            save_cache(cache)
            return result.get("access_token")
            
    # If silent fails, we need interactive browser login
    logger.info("Interactive login required. Opening browser window...")
    try:
        # This will open a local browser window for authentication
        result = app.acquire_token_interactive(
            scopes=scopes,
            login_hint=login_hint
        )
        if "access_token" in result:
            logger.info("Interactive login successful.")
            save_cache(cache)
            return result.get("access_token")
        else:
            error_desc = result.get("error_description", result.get("error", "Unknown error"))
            logger.error(f"Failed to acquire token: {error_desc}")
            raise Exception(f"Authentication failed: {error_desc}")
    except Exception as e:
        logger.error(f"Error during interactive auth: {e}")
        raise e

# SharePoint Online API Calls using Graph API
def get_sharepoint_drive_details(token, site_url):
    logger.info(f"Resolving SharePoint site details for URL: {site_url}")
    parsed = urlparse(site_url)
    hostname = parsed.netloc
    
    # Extract site path, e.g., /sites/yoursite
    path = parsed.path.rstrip('/')
    
    # Query site details to get Site ID
    # Endpoint: GET /sites/{hostname}:{path}
    headers = {"Authorization": f"Bearer {token}"}
    site_endpoint = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{path}"
    
    logger.info(f"Graph API GET: {site_endpoint}")
    res = requests.get(site_endpoint, headers=headers)
    if res.status_code != 200:
        logger.error(f"Failed to get site details from Graph API: {res.text}")
        raise Exception(f"SharePoint Site not found or access denied (HTTP {res.status_code})")
        
    site_data = res.json()
    site_id = site_data.get("id")
    logger.info(f"SharePoint Site ID resolved: {site_id}")
    
    # Query default drive (Shared Documents) of the site
    # Endpoint: GET /sites/{site_id}/drive
    drive_endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
    res = requests.get(drive_endpoint, headers=headers)
    if res.status_code != 200:
        logger.error(f"Failed to get default drive details: {res.text}")
        raise Exception(f"SharePoint document library not found (HTTP {res.status_code})")
        
    drive_data = res.json()
    drive_id = drive_data.get("id")
    logger.info(f"SharePoint Default Drive ID resolved: {drive_id}")
    
    return site_id, drive_id

def check_file_exists_on_sharepoint(token, drive_id, folder_path, filename):
    # Standardize folder path. E.g. "Shared Documents/Backups/2026"
    # Wait, in Microsoft Graph:
    # If folder_path starts with "Shared Documents", the drive_id represents "Shared Documents".
    # So the path in the drive is relative to the root, which excludes "Shared Documents/".
    # E.g., if path is "Shared Documents/Backups/2026", relative path is "Backups/2026".
    rel_path = folder_path
    if folder_path.startswith("Shared Documents/"):
        rel_path = folder_path[len("Shared Documents/"):]
    elif folder_path == "Shared Documents":
        rel_path = ""
        
    headers = {"Authorization": f"Bearer {token}"}
    
    # Construct endpoint for checking item
    # Endpoint: GET /drives/{drive_id}/root:/{rel_path}/{filename}
    if rel_path:
        item_path = f"{rel_path.rstrip('/')}/{filename}"
    else:
        item_path = filename
        
    endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{item_path}"
    logger.info(f"Checking if file exists: {endpoint}")
    res = requests.get(endpoint, headers=headers)
    
    if res.status_code == 200:
        return True
    elif res.status_code == 404:
        return False
    else:
        logger.warning(f"Unexpected response code {res.status_code} checking file existence: {res.text}")
        # Treat as error
        raise Exception(f"Failed to verify duplicate status on SharePoint (HTTP {res.status_code})")

def upload_file_to_sharepoint(token, drive_id, folder_path, filepath, filename):
    rel_path = folder_path
    if folder_path.startswith("Shared Documents/"):
        rel_path = folder_path[len("Shared Documents/"):]
    elif folder_path == "Shared Documents":
        rel_path = ""
        
    headers = {"Authorization": f"Bearer {token}"}
    
    if rel_path:
        item_path = f"{rel_path.rstrip('/')}/{filename}"
    else:
        item_path = filename
        
    file_size = os.path.getsize(filepath)
    
    # If file is smaller than 4MB, do a simple upload
    if file_size < 4 * 1024 * 1024:
        endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{item_path}:/content"
        logger.info(f"Uploading file (simple upload, {file_size} bytes): {filename}")
        
        with open(filepath, 'rb') as f:
            file_content = f.read()
            
        headers_upload = headers.copy()
        headers_upload["Content-Type"] = "application/octet-stream"
        
        res = requests.put(endpoint, headers=headers_upload, data=file_content)
        if res.status_code in [200, 201]:
            logger.info(f"Successfully uploaded: {filename}")
            return True
        else:
            logger.error(f"Failed to upload {filename}: {res.text}")
            raise Exception(f"SharePoint upload failed (HTTP {res.status_code})")
    else:
        # Large file upload session
        endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{item_path}:/createUploadSession"
        logger.info(f"Creating upload session for large file ({file_size} bytes): {filename}")
        
        res = requests.post(endpoint, headers=headers)
        if res.status_code != 200:
            logger.error(f"Failed to create upload session: {res.text}")
            raise Exception(f"SharePoint upload session failed (HTTP {res.status_code})")
            
        session_data = res.json()
        upload_url = session_data.get("uploadUrl")
        
        # Upload file in chunks
        chunk_size = 320 * 1024 * 10  # 3.2MB chunks (must be multiple of 320KB)
        with open(filepath, 'rb') as f:
            start_byte = 0
            while start_byte < file_size:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                
                chunk_len = len(chunk)
                end_byte = start_byte + chunk_len - 1
                
                headers_chunk = {
                    "Content-Length": str(chunk_len),
                    "Content-Range": f"bytes {start_byte}-{end_byte}/{file_size}"
                }
                
                logger.info(f"Uploading chunk {start_byte}-{end_byte} of {file_size}...")
                chunk_res = requests.put(upload_url, headers=headers_chunk, data=chunk)
                
                if chunk_res.status_code in [200, 201]:
                    logger.info(f"Upload complete: {filename}")
                    return True
                elif chunk_res.status_code == 202:
                    start_byte = end_byte + 1
                else:
                    logger.error(f"Failed to upload chunk: {chunk_res.text}")
                    raise Exception(f"Large file upload failed at chunk (HTTP {chunk_res.status_code})")
                    
        return False

# Local folder (OneDrive synced SharePoint) uploads
def handle_local_sync_upload(filepath, filename, config, filter_name=None):
    local_root = config.get("local_sync_path")
    if not local_root or not os.path.exists(local_root):
        raise Exception(f"Local OneDrive Sync Root path does not exist: {local_root}")
        
    sharepoint_folder = config.get("sharepoint_folder")
    
    # Construct target path on disk
    rel_folder = sharepoint_folder
    if sharepoint_folder.startswith("Shared Documents/"):
        rel_folder = sharepoint_folder[len("Shared Documents/"):]
    elif sharepoint_folder == "Shared Documents":
        rel_folder = ""
        
    local_root_norm = os.path.normpath(local_root)
    rel_folder_norm = os.path.normpath(rel_folder)
    
    # If the user-selected path already ends with the subfolder hierarchy, use it directly
    if rel_folder_norm and local_root_norm.lower().endswith(rel_folder_norm.lower()):
        target_dir = local_root_norm
    else:
        target_dir = os.path.normpath(os.path.join(local_root_norm, rel_folder_norm))
        
    # Append spare folder structure if filter_name is spare or replacement
    if filter_name in ["spare", "replacement"]:
        if os.path.basename(target_dir) == "2026":
            target_dir = os.path.normpath(os.path.join(os.path.dirname(target_dir), "Spare 2026"))
        else:
            if not target_dir.endswith("Spare 2026"):
                target_dir = os.path.normpath(os.path.join(target_dir, "Spare 2026"))
    
    if not os.path.exists(target_dir):
        logger.info(f"Creating local directory structure: {target_dir}")
        os.makedirs(target_dir, exist_ok=True)
        
    target_filepath = os.path.normpath(os.path.join(target_dir, filename))
    
    if os.path.exists(target_filepath):
        logger.info(f"File already exists in sync folder, skipping: {filename}")
        return "skipped"
        
    logger.info(f"Copying file to OneDrive sync folder: {target_filepath}")
    shutil.copy2(filepath, target_filepath)
    logger.info(f"Successfully copied: {filename}")
    return "uploaded"

# COM Outlook Desktop Automation
def process_local_outlook(config, temp_dir, upload_callback):
    logger.info("Connecting to Microsoft Outlook desktop client...")
    
    try:
        import win32com.client
    except ImportError:
        logger.error("Python pywin32 library is not installed or available.")
        raise Exception("pywin32 is required to automate local Outlook desktop. Run using run.bat or install pywin32.")
        
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        logger.error(f"Could not connect to Outlook. Is Outlook configured and running? Error: {e}")
        raise Exception("Failed to open local Outlook client.")
        
    target_folder_name = config.get("outlook_folder")
    sender_filter = config.get("sender_filter")
    sender_filter = sender_filter.strip().lower() if sender_filter else ""
    subject_filter_raw = config.get("subject_filter", "onboarding")
    subject_filters = [f.strip().lower() for f in subject_filter_raw.split(",") if f.strip()]
    
    # Try searching for target folder recursively in default store
    logger.info(f"Searching for Outlook folder: '{target_folder_name}'...")
    
    def search_folders(current_folder, target):
        if current_folder.Name.lower() == target.lower():
            return current_folder
        try:
            for sub_folder in current_folder.Folders:
                found = search_folders(sub_folder, target)
                if found:
                    return found
        except Exception:
            pass
        return None
        
    target_folder = None
    try:
        default_store = namespace.DefaultStore
        root_folder = default_store.GetRootFolder()
        target_folder = search_folders(root_folder, target_folder_name)
    except Exception as e:
        logger.warning(f"Error searching default store: {e}. Trying alternative folders extraction...")
        
    if not target_folder:
        # Fallback to search through all folders under top-level folders
        try:
            for folder in namespace.Folders:
                target_folder = search_folders(folder, target_folder_name)
                if target_folder:
                    break
        except Exception:
            pass
            
    if not target_folder:
        logger.error(f"Outlook folder '{target_folder_name}' was not found.")
        raise Exception(f"Folder '{target_folder_name}' not found in Outlook.")
        
    logger.info(f"Found folder: {target_folder.FolderPath}. Scanning messages...")
    
    try:
        messages = target_folder.Items
        # Sort messages by received time descending (newest first)
        messages.Sort("[ReceivedTime]", True)
    except Exception as e:
        logger.error(f"Failed to retrieve messages from folder: {e}")
        raise Exception("Failed to read emails from the folder.")
        
    total_processed = 0
    total_uploaded = 0
    total_skipped = 0
    
    # Avoid modifying items in-place during iteration
    message_list = list(messages)
    logger.info(f"Total emails in folder: {len(message_list)}")
    
    for message in message_list:
        try:
            # Check Sender Email
            sender_email = ""
            try:
                sender_email = message.SenderEmailAddress
            except Exception:
                pass
                
            # If Exchange user, resolve SMTP address
            if sender_email and "@" not in sender_email:
                try:
                    sender = message.Sender
                    if sender:
                        ex_user = sender.GetExchangeUser()
                        if ex_user:
                            sender_email = ex_user.PrimarySmtpAddress
                except Exception as ex_err:
                    logger.debug(f"Exchange SMTP resolution failed: {ex_err}")
                    
            if not sender_email:
                # Try SenderAddress property
                try:
                    sender_email = message.SenderAddress
                except Exception:
                    pass
                    
            if not sender_email:
                continue
                
            sender_email_lower = sender_email.lower()
            subject_lower = message.Subject.lower() if message.Subject else ""
            
            # Check filters
            sender_match = (not sender_filter) or (sender_email_lower == sender_filter)
            
            matched_filter = None
            attachments = None
            subject_matched = False
            if sender_match:
                # 1. Check subject filters
                for sf in subject_filters:
                    if sf and sf in subject_lower:
                        matched_filter = sf
                        subject_matched = True
                        break
                
                # 2. Check attachment name filters if subject didn't match
                if not matched_filter:
                    try:
                        atts_col = message.Attachments
                        count = atts_col.Count
                        attachments_list = []
                        for idx in range(1, count + 1):
                            att = atts_col.Item(idx)
                            attachments_list.append(att)
                            att_name_lower = att.FileName.lower()
                            for sf in subject_filters:
                                if sf and sf in att_name_lower:
                                    matched_filter = sf
                        if matched_filter:
                            attachments = attachments_list
                    except Exception as att_err:
                        logger.warning(f"Failed to scan attachments for email '{message.Subject}': {att_err}")
            
            if matched_filter:
                logger.info(f"Matching email found! Subject: '{message.Subject}' | From: {sender_email} (Matched filter: '{matched_filter}')")
                
                # Get email date
                received_time = message.ReceivedTime
                # received_time can be a pywintypes.datetime object
                try:
                    # Parse pywintypes datetime to python standard datetime
                    date_str = received_time.strftime("%Y%m%d")
                except Exception:
                    # Fallback
                    date_str = datetime.datetime.now().strftime("%Y%m%d")
                    
                if attachments is None:
                    try:
                        atts_col = message.Attachments
                        count = atts_col.Count
                        attachments = [atts_col.Item(idx) for idx in range(1, count + 1)]
                    except Exception as att_err:
                        logger.error(f"Failed to read attachments: {att_err}")
                        continue
                        
                if len(attachments) == 0:
                    logger.info("No attachments in this email, skipping.")
                    continue
                    
                email_success = True
                
                for attachment in attachments:
                    orig_filename = attachment.FileName
                    orig_filename_lower = orig_filename.lower()
                    
                    # If email matched ONLY via attachment name, skip attachments whose name does not contain any of the subject filters
                    if not subject_matched:
                        matched_att_name = False
                        for sf in subject_filters:
                            if sf and sf in orig_filename_lower:
                                matched_att_name = True
                                break
                        if not matched_att_name:
                            logger.info(f"Skipping non-matching attachment: '{orig_filename}'")
                            continue
                    
                    # Bifurcation rule based on attachment filename OR email subject keywords
                    is_spare = ("spare" in orig_filename_lower) or ("replacement" in orig_filename_lower) or ("spare" in subject_lower) or ("replacement" in subject_lower)
                    
                    if is_spare:
                        if "replacement" in orig_filename_lower or "replacement" in subject_lower:
                            category = "replacement"
                        else:
                            category = "spare"
                    else:
                        category = "onboarding"
                        
                    new_filename = f"{category}_{date_str}_{orig_filename}"
                    temp_filepath = os.path.join(temp_dir, new_filename)
                    logger.info(f"Downloading attachment: {orig_filename} -> {new_filename} (Category: {category})")
                    
                    try:
                        attachment.SaveAsFile(temp_filepath)
                    except Exception as e:
                        logger.error(f"Failed to save attachment locally: {e}")
                        email_success = False
                        continue
                        
                    # Call upload callback to upload file to SharePoint (with filter name for bifurcation)
                    try:
                        status = upload_callback(temp_filepath, new_filename, filter_name=category)
                        if status == "uploaded":
                            total_uploaded += 1
                        elif status == "skipped":
                            total_skipped += 1
                    except Exception as e:
                        logger.error(f"Failed to upload renamed file to SharePoint: {e}")
                        email_success = False
                    finally:
                        # Delete temp file
                        if os.path.exists(temp_filepath):
                            try:
                                os.remove(temp_filepath)
                                logger.info(f"Cleaned up local temp file: {new_filename}")
                            except Exception as cleanup_err:
                                logger.warning(f"Failed to delete local temp file {temp_filepath}: {cleanup_err}")
                                
                if email_success:
                    total_processed += 1
                    # Mark processed email as read
                    if message.UnRead:
                        try:
                            message.UnRead = False
                            message.Save()
                            logger.info("Marked email as READ.")
                        except Exception as e:
                            logger.warning(f"Could not mark email as read: {e}")
                            
        except Exception as e:
            logger.error(f"Error processing email message: {e}")
            
    logger.info(f"Execution statistics: Emails Processed: {total_processed} | Files Uploaded: {total_uploaded} | Files Skipped (Duplicates): {total_skipped}")
    return total_processed, total_uploaded, total_skipped

# Graph API Outlook Online integration (Cloud mode)
def process_graph_outlook(token, config, temp_dir, upload_callback):
    logger.info("Connecting to Exchange Online via Microsoft Graph API...")
    headers = {"Authorization": f"Bearer {token}"}
    
    target_folder_name = config.get("outlook_folder")
    sender_filter = config.get("sender_filter")
    sender_filter = sender_filter.strip().lower() if sender_filter else ""
    subject_filter_raw = config.get("subject_filter", "onboarding")
    subject_filters = [f.strip().lower() for f in subject_filter_raw.split(",") if f.strip()]
    
    # Find folder ID by name
    # We query: GET /me/mailFolders
    logger.info(f"Querying mail folders to find '{target_folder_name}'...")
    folders_url = "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100"
    res = requests.get(folders_url, headers=headers)
    if res.status_code != 200:
        logger.error(f"Failed to load mail folders: {res.text}")
        raise Exception(f"Failed to list mail folders from Graph (HTTP {res.status_code})")
        
    folders = res.json().get("value", [])
    folder_id = None
    
    for folder in folders:
        if folder.get("displayName", "").lower() == target_folder_name.lower():
            folder_id = folder.get("id")
            break
            
    # If not found in root folders, check child folders of Inbox recursively or search (simplified search)
    if not folder_id:
        logger.info(f"Folder '{target_folder_name}' not found at root level. Querying inbox child folders...")
        # Get inbox folder ID first
        inbox_url = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox"
        inbox_res = requests.get(inbox_url, headers=headers)
        if inbox_res.status_code == 200:
            inbox_id = inbox_res.json().get("id")
            child_folders_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{inbox_id}/childFolders?$top=100"
            child_res = requests.get(child_folders_url, headers=headers)
            if child_res.status_code == 200:
                child_folders = child_res.json().get("value", [])
                for child in child_folders:
                    if child.get("displayName", "").lower() == target_folder_name.lower():
                        folder_id = child.get("id")
                        break
                        
    if not folder_id:
        logger.error(f"Folder '{target_folder_name}' not found via Graph API.")
        raise Exception(f"Outlook folder '{target_folder_name}' not found in Exchange Online mailbox.")
        
    logger.info(f"Found folder in Graph. ID: {folder_id}. Querying messages...")
    
    # Query emails in folder
    # Filter: sender address and subject containing the filter terms can be done in the OData query or locally.
    # To keep it extremely robust and matches our requirements:
    # Query all messages in the folder (newest first)
    messages_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?$orderby=receivedDateTime%20desc&$top=50"
    res = requests.get(messages_url, headers=headers)
    if res.status_code != 200:
        logger.error(f"Failed to query messages in folder: {res.text}")
        raise Exception("Failed to retrieve emails from folder via Graph.")
        
    messages = res.json().get("value", [])
    logger.info(f"Total emails fetched from folder: {len(messages)}")
    
    total_processed = 0
    total_uploaded = 0
    total_skipped = 0
    
    for message in messages:
        try:
            sender = message.get("sender", {}).get("emailAddress", {}).get("address", "")
            subject = message.get("subject", "") or ""
            
            if not sender:
                continue
                
            sender_match = (not sender_filter) or (sender.lower() == sender_filter)
            
            matched_filter = None
            attachments = None
            subject_matched = False
            if sender_match:
                # 1. Check subject filters
                for sf in subject_filters:
                    if sf and sf in subject.lower():
                        matched_filter = sf
                        subject_matched = True
                        break
                
                # 2. Check attachment name filters if subject didn't match
                if not matched_filter:
                    message_id = message.get("id")
                    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
                    att_res = requests.get(attachments_url, headers=headers)
                    if att_res.status_code == 200:
                        attachments = att_res.json().get("value", [])
                        for att in attachments:
                            att_name_lower = att.get("name", "").lower()
                            for sf in subject_filters:
                                if sf and sf in att_name_lower:
                                    matched_filter = sf
                                    break
                            if matched_filter:
                                break
            
            if matched_filter:
                message_id = message.get("id")
                logger.info(f"Matching email found! Subject: '{subject}' | From: {sender} (Matched by: '{matched_filter}')")
                
                # Format Date YYYYMMDD
                received_time_str = message.get("receivedDateTime", "")
                try:
                    # format: 2026-05-20T10:00:00Z
                    dt = datetime.datetime.strptime(received_time_str.split('.')[0].replace('Z', ''), "%Y-%m-%dT%H:%M:%S")
                    date_str = dt.strftime("%Y%m%d")
                except Exception:
                    date_str = datetime.datetime.now().strftime("%Y%m%d")
                    
                # Query attachments if we haven't already
                if attachments is None:
                    attachments_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
                    att_res = requests.get(attachments_url, headers=headers)
                    if att_res.status_code != 200:
                        logger.warning(f"Failed to query attachments for message {message_id}: {att_res.text}")
                        continue
                    attachments = att_res.json().get("value", [])
                    
                if not attachments:
                    logger.info("No attachments found in this email.")
                    continue
                    
                email_success = True
                
                for att in attachments:
                    # Verify it's a file attachment (not inline itemAttachment)
                    if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
                        logger.info(f"Skipping non-file attachment: {att.get('name')}")
                        continue
                        
                    orig_filename = att.get("name")
                    if not orig_filename:
                        continue
                    orig_filename_lower = orig_filename.lower()
                    
                    # If email matched ONLY via attachment name, skip attachments whose name does not contain any of the subject filters
                    if not subject_matched:
                        matched_att_name = False
                        for sf in subject_filters:
                            if sf and sf in orig_filename_lower:
                                matched_att_name = True
                                break
                        if not matched_att_name:
                            logger.info(f"Skipping non-matching attachment: '{orig_filename}'")
                            continue
                            
                    content_bytes_base64 = att.get("contentBytes")
                    
                    if not content_bytes_base64:
                        logger.warning(f"Empty attachment payload for {orig_filename}.")
                        continue
                        
                    # Decode attachment bytes
                    import base64
                    try:
                        file_data = base64.b64decode(content_bytes_base64)
                    except Exception as base64_err:
                        logger.error(f"Failed to decode base64 attachment content: {base64_err}")
                        email_success = False
                        continue
                        
                    # Bifurcation rule based on attachment filename OR email subject
                    # Bifurcation rule based on attachment filename OR email subject keywords
                    is_spare = ("spare" in orig_filename_lower) or ("replacement" in orig_filename_lower) or ("spare" in subject.lower()) or ("replacement" in subject.lower())
                    
                    if is_spare:
                        if "replacement" in orig_filename_lower or "replacement" in subject.lower():
                            category = "replacement"
                        else:
                            category = "spare"
                    else:
                        category = "onboarding"
                        
                    new_filename = f"{category}_{date_str}_{orig_filename}"
                    temp_filepath = os.path.join(temp_dir, new_filename)
                    
                    logger.info(f"Saving attachment: {orig_filename} -> {new_filename} (Category: {category})")
                    try:
                        with open(temp_filepath, 'wb') as f:
                            f.write(file_data)
                    except Exception as e:
                        logger.error(f"Failed to save attachment locally: {e}")
                        email_success = False
                        continue
                        
                    # Upload to SharePoint
                    try:
                        status = upload_callback(temp_filepath, new_filename, filter_name=category)
                        if status == "uploaded":
                            total_uploaded += 1
                        elif status == "skipped":
                            total_skipped += 1
                    except Exception as e:
                        logger.error(f"Failed to upload renamed file: {e}")
                        email_success = False
                    finally:
                        # Clean up temp file
                        if os.path.exists(temp_filepath):
                            try:
                                os.remove(temp_filepath)
                                logger.info(f"Cleaned up local temp file: {new_filename}")
                            except Exception as cleanup_err:
                                logger.warning(f"Failed to delete local temp file {temp_filepath}: {cleanup_err}")
                                
                if email_success:
                    total_processed += 1
                    # Mark email as read: PATCH /me/messages/{message_id} with isRead = true
                    if not message.get("isRead"):
                        logger.info("Marking email as READ...")
                        patch_url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}"
                        patch_res = requests.patch(patch_url, headers=headers, json={"isRead": True})
                        if patch_res.status_code == 200:
                            logger.info("Successfully marked email as READ.")
                        else:
                            logger.warning(f"Failed to mark email as read (HTTP {patch_res.status_code}): {patch_res.text}")
                            
        except Exception as e:
            logger.error(f"Error processing email message via Graph: {e}")
            
    logger.info(f"Execution statistics: Emails Processed: {total_processed} | Files Uploaded: {total_uploaded} | Files Skipped (Duplicates): {total_skipped}")
    return total_processed, total_uploaded, total_skipped

# Main Orchestrator Run function
def sync_run(config, status_callback=None):
    logger.info("Starting synchronization process...")
    
    # Create a local temp folder in the workspace for downloading
    import tempfile
    temp_dir = os.path.join(tempfile.gettempdir(), "outlook_sharepoint_sync_temp")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir, exist_ok=True)
        
    try:
        sync_method = config.get("sync_method", "graph")
        outlook_source = config.get("outlook_source", "local")
        
        token = None
        
        # Decide if we need an MS Access Token (if either Outlook or SharePoint is using Graph API)
        needs_token = (sync_method == "graph") or (outlook_source == "graph")
        
        if needs_token:
            if status_callback:
                status_callback("Authenticating with Microsoft 365...", logging.INFO)
            token = acquire_token(config)
            
        # Define upload callback for the file
        def upload_callback(filepath, filename, filter_name=None):
            if sync_method == "local":
                # Copy to local OneDrive synced path
                return handle_local_sync_upload(filepath, filename, config, filter_name=filter_name)
            else:
                # Microsoft Graph API Cloud upload
                if not token:
                    raise Exception("Access token not acquired.")
                
                # Fetch Site ID and Drive ID
                site_url = config.get("sharepoint_site")
                folder_path = config.get("sharepoint_folder")
                if filter_name in ["spare", "replacement"]:
                    if folder_path.rstrip("/").endswith("/2026"):
                        folder_path = folder_path.rstrip("/")[:-5] + "/Spare 2026"
                    else:
                        if not folder_path.rstrip("/").endswith("/Spare 2026"):
                            folder_path = folder_path.rstrip("/") + "/Spare 2026"
                site_id, drive_id = get_sharepoint_drive_details(token, site_url)
                
                # Check duplicate
                exists = check_file_exists_on_sharepoint(token, drive_id, folder_path, filename)
                if exists:
                    logger.info(f"File already exists on SharePoint, skipping: {filename}")
                    return "skipped"
                    
                # Upload
                upload_file_to_sharepoint(token, drive_id, folder_path, filepath, filename)
                return "uploaded"
                
        # Run Outlook scan
        if outlook_source == "local":
            if status_callback:
                status_callback("Scanning local Outlook mail folder...", logging.INFO)
            
            com_initialized = False
            if pythoncom is not None:
                try:
                    pythoncom.CoInitialize()
                    com_initialized = True
                except Exception as ce:
                    logger.warning(f"Failed to CoInitialize COM: {ce}")
            
            try:
                stats = process_local_outlook(config, temp_dir, upload_callback)
            finally:
                if com_initialized:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
        else:
            if status_callback:
                status_callback("Scanning Exchange Online mailbox...", logging.INFO)
            stats = process_graph_outlook(token, config, temp_dir, upload_callback)
            
        logger.info("Synchronization process completed successfully.")
        return True, stats
        
    except Exception as e:
        logger.exception("An error occurred during synchronization:")
        return False, str(e)
    finally:
        # Clean up temp folder
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.warning(f"Failed to remove temp directory {temp_dir}: {e}")
