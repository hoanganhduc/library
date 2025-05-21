#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Standard library imports
import argparse
import concurrent.futures
import json
import os
import html
import io
import sys
import re
from datetime import datetime

# Third-party imports
import pdfkit
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from pyzotero import zotero
import io
import tempfile
from googleapiclient.http import MediaIoBaseDownload
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import tempfile
import os
import random
import sqlite3

# Google Drive API imports
# Remark: Create a service account in Google Console and share Zotero folder with the service account email. If you don't share it, you won't be able to access the files.
def authenticate_google_drive(service_account_file):
    """
    Authenticate to Google Drive using a service account.
    
    Args:
        service_account_file (str): Path to the service account key JSON file or JSON string
        
    Returns:
        google.auth.credentials.Credentials: Google API credentials
    """
    
    # Use full Drive access for reading and downloading files
    SCOPES = [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/drive.metadata.readonly'
    ]
    
    creds = None
    service_info = {}
    
    try:
        # Check if input is a JSON string (starts with '{' and ends with '}')
        if service_account_file.strip().startswith('{') and service_account_file.strip().endswith('}'):
            # Parse JSON string directly
            service_info = json.loads(service_account_file)
            
            # Create credentials from parsed JSON
            service_json_io = io.StringIO(service_account_file)
            creds = service_account.Credentials.from_service_account_info(
                service_info, scopes=SCOPES)
            
            print("Authenticated using provided JSON string")
        else:
            # Treat as file path
            if not os.path.exists(service_account_file):
                print(f"Error: Service account key file not found: {service_account_file}")
                return None
                
            # Read the file and load JSON for getting email
            with open(service_account_file, 'r') as f:
                service_info = json.load(f)
            
            # Create credentials from file
            creds = service_account.Credentials.from_service_account_file(
                service_account_file, scopes=SCOPES)
                
            print(f"Authenticated using service account file: {service_account_file}")
        
        # Get service account email for logging
        service_email = service_info.get('client_email', 'unknown-service-account')
        print(f"Authenticated as service account: {service_email}")
            
    except Exception as e:
        print(f"Error authenticating with service account: {str(e)}")
        return None
    
    return creds

def test_google_drive_access(google_creds, verbose=False):
    """
    Test access to Google Drive using Google credentials.
    
    Args:
        google_creds: Google API credentials object
        verbose (bool): Whether to display verbose output
        
    Returns:
        tuple: (success, message) where success is a boolean indicating if the test was successful,
                and message contains additional information
    """
    if verbose:
        print("Testing Google Drive access using provided credentials")
    
    try:
        if not google_creds:
            return False, "No credentials provided. Authentication failed."
            
        # Build the Drive API client
        drive_service = build('drive', 'v3', credentials=google_creds)
        
        # Try to get account information and file count
        about = drive_service.about().get(fields="user,storageQuota").execute()
        files = drive_service.files().list(
            pageSize=1, 
            fields="files(id,name),nextPageToken"
        ).execute()
        
        # Get service account email from credentials or user info
        service_email = 'Unknown'
        if hasattr(google_creds, 'service_account_email'):
            service_email = google_creds.service_account_email
        elif 'user' in about and 'emailAddress' in about['user']:
            service_email = about['user']['emailAddress']
        
        storage_used = int(about.get('storageQuota', {}).get('usage', 0)) / (1024 * 1024)  # Convert to MB
        storage_total = int(about.get('storageQuota', {}).get('limit', 0)) / (1024 * 1024 * 1024)  # Convert to GB
        
        # Count files (this may take a while for large accounts, so we estimate)
        file_count = "at least 1" if files.get('files') else "0"
        if 'nextPageToken' in files:
            file_count = "more than 100"  # Just an indication that there are many files
            
        # Format the success message
        message = (
            f"Successfully connected to Google Drive!\n"
            f"Service Account: {service_email}\n"
            f"Storage used: {storage_used:.2f} MB / {storage_total:.2f} GB\n"
            f"Files: {file_count}"
        )
        
        return True, message
        
    except Exception as e:
        error_message = f"Error accessing Google Drive: {str(e)}"
        if verbose:
            print(error_message)
        return False, error_message

def search_file_in_drive(drive_service, query, max_results=10, folder_name=None, include_shared=True):
    """
    Search for files in Google Drive based on a query.
    
    Args:
        drive_service: Google Drive service instance
        query (str): Search query string
        max_results (int): Maximum number of results to return
        folder_name (str, optional): Name of folder to search within (default: None, searches all of Drive)
        include_shared (bool): Whether to include files shared with the user (default: True)
        
    Returns:
        list: List of file metadata matching the query
    """
    results = []
    page_token = None
    
    # If folder name is specified, find its ID and modify the query
    folder_id = None
    if folder_name:
        # Search for the folder (include both owned and shared folders)
        folder_query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        folder_response = drive_service.files().list(
            q=folder_query,
            spaces='drive',
            fields='files(id, name)',
            pageSize=1).execute()
        
        folders = folder_response.get('files', [])
        if folders:
            folder_id = folders[0]['id']
            # Modify query to search within the specific folder
            query = f"{query} and '{folder_id}' in parents"
    
    # Search in both owned and shared files
    while True:
        search_params = {
            'q': query,
            'spaces': 'drive',
            'fields': 'nextPageToken, files(id, name, webViewLink)',
            'pageToken': page_token,
            'pageSize': max_results
        }
        
        response = drive_service.files().list(**search_params).execute()
        results.extend(response.get('files', []))
        
        # If we need to specifically search in shared files and we haven't reached max results
        if include_shared and len(results) < max_results and not folder_id:
            # Create a separate query for shared files
            shared_query = f"{query} and sharedWithMe=true"
            shared_response = drive_service.files().list(
                q=shared_query,
                spaces='drive',
                fields='files(id, name, webViewLink)',
                pageSize=max_results - len(results)
            ).execute()
            
            # Add any unique shared files to results
            shared_files = shared_response.get('files', [])
            existing_ids = {file['id'] for file in results}
            for file in shared_files:
                if file['id'] not in existing_ids:
                    results.append(file)
                    existing_ids.add(file['id'])
        
        page_token = response.get('nextPageToken', None)
        
        if page_token is None or len(results) >= max_results:
            break
            
    return results[:max_results]

def get_drive_url_by_filename(google_creds, filename, exact_match=True, folder_name=None, return_all=False, verbose=False):
    """
    Find a file in Google Drive by name and return its URL using provided Google credentials.
    
    Args:
        google_creds: Already authenticated Google credentials object
        filename (str): Name of the file to search for
        exact_match (bool): If True, match exact filename, otherwise partial match
        folder_name (str, optional): Name of folder to search within (None searches all of Drive)
        return_all (bool): If True, return all matching files, not just the first one
        verbose (bool): Whether to display progress messages
        
    Returns:
        Union[str, List[str], None]: URL(s) of the file(s) if found, None otherwise
    """
    try:
        if verbose:
            print(f"Searching for file: {filename} in Google Drive")
            
        # Check if credentials are valid
        if not google_creds:
            if verbose:
                print("No valid Google credentials provided")
            return None
            
        # Build the Drive API client
        drive_service = build('drive', 'v3', credentials=google_creds)
        
        # Escape single quotes in filename for query
        safe_filename = filename.replace("'", "\\'")
        
        # Construct the search query based on the filename
        if exact_match:
            query = f"name = '{safe_filename}' and trashed = false"
        else:
            query = f"name contains '{safe_filename}' and trashed = false"
            
        # Search for the file, possibly in a specific folder
        results = search_file_in_drive(drive_service, query, max_results=10 if return_all else 1, folder_name=folder_name)
        
        if verbose:
            print(f"Found {len(results)} matching files")
            
        # Return based on return_all parameter
        if not results:
            return None
        elif return_all:
            return [item.get('webViewLink') for item in results if 'webViewLink' in item]
        else:
            return results[0].get('webViewLink')
            
    except Exception as e:
        print(f"Error accessing Google Drive: {str(e)}", file=sys.stderr)
        return None

def print_progress(message, verbose=True, level=1, file=sys.stdout):
    """
    Print progress messages to track script execution.
    
    Args:
        message (str): The progress message to display
        verbose (bool): Whether to display the message (default: True)
        level (int): Importance level of the message (higher = more important)
        file: File object to write to (default: sys.stdout)
    """
    if verbose:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {message}", file=file)
        file.flush()  # Ensure the message is written immediately

def connect_to_zotero(library_id, library_type, api_key):
    """Create and return a Zotero connection."""
    return zotero.Zotero(library_id, library_type, api_key)

def get_collections_from_sqlite(sqlite_path, verbose=False):
    """Get collections from a Zotero SQLite database file."""
    try:
        conn = sqlite3.connect(sqlite_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("SELECT collectionID, collectionName, key FROM collections ORDER BY collectionName")
        rows = cur.fetchall()
        conn.close()
        
        if rows:
            if verbose:
                print_progress(f"Found {len(rows)} collections in SQLite database", verbose)
            return [{'data': {'name': row['collectionName'], 'key': row['key']}} for row in rows]
    except Exception as e:
        if verbose:
            print_progress(f"Error querying SQLite database: {e}", verbose, file=sys.stderr)
    
    return []

def get_local_collections(verbose=False):
    """Get collections from local Zotero SQLite database."""
    local_paths = [
        os.path.expanduser("~/Zotero/zotero.sqlite"),
        os.path.expanduser("~/Library/Application Support/Zotero/zotero.sqlite"),
        os.path.expanduser("~/.zotero/zotero.sqlite"),
        "./zotero.sqlite"
    ]
    
    for path in local_paths:
        if os.path.exists(path):
            if verbose:
                print_progress(f"Found local zotero.sqlite at {path}", verbose)
            collections = get_collections_from_sqlite(path, verbose)
            if collections:
                return collections
            break  # Only try first available database
    
    return []

def get_gdrive_collections(google_creds, verbose=False):
    """Get collections from Google Drive SQLite database."""
    if not google_creds:
        return []
        
    try:
        if verbose:
            print_progress("Searching for zotero.sqlite in Google Drive...", verbose)
            
        drive_service = build('drive', 'v3', credentials=google_creds)
        drive_url = get_drive_url_by_filename(google_creds, "zotero.sqlite", exact_match=True, verbose=verbose)
        
        if drive_url:
            file_id = extract_file_id_from_drive_url(drive_url)
            if file_id:
                temp_path = download_file_from_drive(drive_service, file_id, verbose=verbose)
                
                if temp_path and os.path.exists(temp_path):
                    collections = get_collections_from_sqlite(temp_path, verbose)
                    
                    # Clean up temp file
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
                    
                    return collections
    except Exception as e:
        if verbose:
            print_progress(f"Error accessing Google Drive: {e}", verbose, file=sys.stderr)
    
    return []

def list_collections(zot, verbose=False):
    """
    List all collections in the Zotero library with priority:
    1. Local Zotero database
    2. Google Drive Zotero database
    3. Online Zotero API
    
    Args:
        zot: Zotero API client instance
        verbose (bool): Whether to display progress messages
    
    Returns:
        list: List of collections
    """
    # Step 1: Try local database
    collections = get_local_collections(verbose)
    if collections:
        return collections
    
    # Step 2: Try Google Drive
    try:
        # Get google_creds from global scope if available
        google_creds = globals().get('google_creds', None)
        collections = get_gdrive_collections(google_creds, verbose)
        if collections:
            return collections
    except Exception:
        pass
    
    # Step 3: Fall back to online API
    if verbose:
        print_progress("Fetching collections from online Zotero API...", verbose)
    collections = zot.collections()
    if verbose and collections:
        print_progress(f"Found {len(collections)} collections via online API", verbose)
    
    return collections

def get_items_from_sqlite(db_path, collection=None, item_type=None, verbose=False):
    """Get items from a Zotero SQLite database file."""
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        
        # Base query for retrieving items
        sql = """
            SELECT items.itemID, items.key, itemTypes.typeName, itemDataValues.value as title
            FROM items
            JOIN itemTypes ON items.itemTypeID = itemTypes.itemTypeID
            LEFT JOIN itemData ON items.itemID = itemData.itemID AND itemData.fieldID = (
                SELECT fieldID FROM fields WHERE fieldName = 'title'
            )
            LEFT JOIN itemDataValues ON itemData.valueID = itemDataValues.valueID
            WHERE items.itemID NOT IN (SELECT itemID FROM deletedItems)
            AND itemTypes.typeName NOT IN ('note', 'attachment')
        """
        params = []
        
        # Add collection filter if specified
        if collection:
            sql += """ 
                AND items.itemID IN (
                    SELECT itemID FROM collectionItems 
                    WHERE collectionID = (SELECT collectionID FROM collections WHERE key = ?)
                )
            """
            params.append(collection)
            
        # Add item type filter if specified
        if item_type:
            sql += " AND itemTypes.typeName = ?"
            params.append(item_type)
            
        # Execute query
        cur.execute(sql, params)
        rows = cur.fetchall()
        
        # Convert to format similar to Zotero API
        items = []
        for row in rows:
            items.append({
                'key': row['key'],
                'data': {
                    'title': row.get('title', 'Unknown'),
                    'itemType': row['typeName']
                }
            })
        
        conn.close()
        
        if verbose:
            print_progress(f"Found {len(items)} items in SQLite database", verbose)
        return items
        
    except Exception as e:
        if verbose:
            print_progress(f"Error querying SQLite database: {e}", verbose, file=sys.stderr)
        return []

def get_items_from_local_db(collection=None, item_type=None, verbose=False):
    """Get items from local Zotero SQLite database."""
    local_paths = [
        os.path.expanduser("~/Zotero/zotero.sqlite"),
        os.path.expanduser("~/Library/Application Support/Zotero/zotero.sqlite"),
        os.path.expanduser("~/.zotero/zotero.sqlite"),
        "./zotero.sqlite"
    ]
    
    for path in local_paths:
        if os.path.exists(path):
            if verbose:
                print_progress(f"Found local zotero.sqlite at {path}", verbose)
            items = get_items_from_sqlite(path, collection, item_type, verbose)
            if items:
                return items
            break  # Only try first available database
    
    return []

def get_items_from_gdrive(google_creds, collection=None, item_type=None, verbose=False):
    """Get items from Google Drive SQLite database."""
    if not google_creds:
        return []
        
    try:
        if verbose:
            print_progress("Searching for zotero.sqlite in Google Drive...", verbose)
            
        drive_service = build('drive', 'v3', credentials=google_creds)
        drive_url = get_drive_url_by_filename(google_creds, "zotero.sqlite", exact_match=True, verbose=verbose)
        
        if drive_url:
            file_id = extract_file_id_from_drive_url(drive_url)
            if file_id:
                temp_path = download_file_from_drive(drive_service, file_id, verbose=verbose)
                
                if temp_path and os.path.exists(temp_path):
                    items = get_items_from_sqlite(temp_path, collection, item_type, verbose)
                    
                    # Clean up temp file
                    try:
                        os.remove(temp_path)
                    except Exception:
                        pass
                    
                    return items
    except Exception as e:
        if verbose:
            print_progress(f"Error accessing Google Drive: {e}", verbose, file=sys.stderr)
    
    return []

def get_items(zot, collection=None, item_type=None, verbose=True):
    """
    Retrieve all items based on filters with priority:
    1. Local Zotero database
    2. Google Drive Zotero database
    3. Online Zotero API
    """
    # Step 1: Try local database
    items = get_items_from_local_db(collection, item_type, verbose)
    if items:
        return items
    
    # Step 2: Try Google Drive
    try:
        # Get google_creds from global scope if available
        google_creds = globals().get('google_creds', None)
        if google_creds:
            items = get_items_from_gdrive(google_creds, collection, item_type, verbose)
            if items:
                return items
    except Exception as e:
        if verbose:
            print_progress(f"Error accessing Google Drive: {e}", verbose, file=sys.stderr)
    
    # Step 3: Fall back to online API
    if verbose:
        if collection:
            print_progress(f"Fetching items from collection {collection} via online API...", verbose)
        elif item_type:
            print_progress(f"Fetching items with type '{item_type}' via online API...", verbose)
        else:
            print_progress("Fetching all library items via online API...", verbose)
    
    try:
        if collection:
            items = zot.everything(zot.collection_items(collection))
        else:
            if item_type:
                items = zot.everything(zot.items(itemType=item_type))
            else:
                items = zot.everything(zot.items())
                
        # Filter out notes and attachments
        filtered_items = [item for item in items if item['data'].get('itemType') not in ['note', 'attachment']]
        
        # Remove relations field from each item
        for item in filtered_items:
            if 'relations' in item['data']:
                del item['data']['relations']
                
        if verbose:
            print_progress(f"Retrieved {len(filtered_items)} items via online API", verbose)
            
        return filtered_items
        
    except Exception as e:
        if verbose:
            print_progress(f"Error retrieving items from online API: {e}", verbose, file=sys.stderr)
        return []

def get_attachment_paths(zot, item, google_creds=None, verbose=False):
    """
    Get attachment paths for a given item and their Google Drive URLs if available.
    Supports various file types including PDF, DJVU, EPUB, AZW3, MOBI and more.
    
    Args:
        zot: Zotero API client instance
        item: Zotero item to get attachments for
        google_creds: Google API credentials object (already authenticated)
        verbose (bool): Whether to display progress messages
    
    Returns:
        list: A list of dictionaries with keys 'local_path' and 'drive_url' (None if not found)
    """
    if not item:
        return []
    
    attachment_info = []
    
    # Try to get attachments from local database first
    try:
        local_paths = [
            os.path.expanduser("~/Zotero/zotero.sqlite"),
            os.path.expanduser("~/Library/Application Support/Zotero/zotero.sqlite"),
            os.path.expanduser("~/.zotero/zotero.sqlite"),
            "./zotero.sqlite"
        ]
        
        for db_path in local_paths:
            if os.path.exists(db_path):
                conn = sqlite3.connect(db_path)
                conn.row_factory = sqlite3.Row
                cur = conn.cursor()
                
                # Get attachments from SQLite database
                cur.execute("""
                    SELECT att.itemID, att.key, att.contentType, att.path, items.key as parentKey, att.filename
                    FROM itemAttachments AS att
                    JOIN items ON att.itemID = items.itemID
                    WHERE att.parentItemID = (SELECT itemID FROM items WHERE key = ?)
                    AND att.contentType IN (
                        'application/pdf', 'image/vnd.djvu', 'video/mp4',
                        'application/vnd.ms-powerpoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                        'application/epub+zip', 'application/vnd.amazon.ebook',
                        'application/x-mobi8-ebook', 'application/x-mobipocket-ebook',
                        'application/vnd.comicbook+zip', 'application/x-cbr',
                        'application/x-fictionbook+xml', 'text/plain'
                    )
                """, (item['key'],))
                
                rows = cur.fetchall()
                conn.close()
                
                if rows:
                    if verbose:
                        print_progress(f"Found {len(rows)} attachments in local database", verbose)
                    
                    for row in rows:
                        attachment_id = row['key']
                        filename = row['filename']
                        local_path = f"storage/{attachment_id}/{filename}"
                        
                        # Initialize with local path only
                        info = {'local_path': local_path, 'drive_url': None}
                        
                        # If Google credentials are provided, search in Drive
                        if google_creds:
                            try:
                                drive_url = get_drive_url_by_filename(
                                    google_creds, 
                                    filename, 
                                    exact_match=True,
                                    verbose=verbose
                                )
                                if drive_url:
                                    info['drive_url'] = drive_url
                            except Exception as e:
                                if verbose:
                                    print_progress(f"Error searching Google Drive for {filename}: {e}", verbose, file=sys.stderr)
                        
                        attachment_info.append(info)
                    
                    return attachment_info
                break  # Only try first available database
    
    except Exception as e:
        if verbose:
            print_progress(f"Error querying local database for attachments: {e}", verbose, file=sys.stderr)
    
    # Fall back to API if no local attachments found
    try:
        attachments = zot.children(item['key'])
        
        for attachment in attachments:
            # Check if it's an attachment of supported type
            if attachment['data'].get('itemType') == 'attachment' and 'contentType' in attachment['data']:
                content_type = attachment['data']['contentType']
                if content_type in [
                    'application/pdf', 
                    'image/vnd.djvu',
                    'video/mp4',
                    'application/vnd.ms-powerpoint',
                    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    'application/epub+zip',
                    'application/vnd.amazon.ebook',
                    'application/x-mobi8-ebook',
                    'application/x-mobipocket-ebook',
                    'application/vnd.comicbook+zip',
                    'application/x-cbr',
                    'application/x-fictionbook+xml',
                    'text/plain'
                ]:
                    # Get the file information
                    if 'key' in attachment and 'filename' in attachment['data']:
                        attachment_id = attachment['key']
                        filename = attachment['data']['filename']
                        local_path = f"storage/{attachment_id}/{filename}"
                        
                        # Initialize with local path only
                        info = {'local_path': local_path, 'drive_url': None}
                        
                        # If Google credentials are provided, search in Drive
                        if google_creds:
                            if verbose:
                                print_progress(f"Searching for {filename} in Google Drive", verbose)
                            try:
                                # Search using Google credentials
                                drive_url = get_drive_url_by_filename(
                                    google_creds, 
                                    filename, 
                                    exact_match=True,
                                    folder_name=None, 
                                    verbose=verbose
                                )
                                if drive_url:
                                    info['drive_url'] = drive_url
                                    if verbose:
                                        print_progress(f"Found Google Drive URL for {filename}", verbose)
                            except Exception as e:
                                if verbose:
                                    print_progress(f"Error searching Google Drive for {filename}: {e}", verbose, file=sys.stderr)
                        
                        attachment_info.append(info)
    
    except Exception as e:
        print(f"Error getting attachments for item {item.get('data', {}).get('title', 'Unknown')}: {e}", file=sys.stderr)
    
    return attachment_info

def extract_doi(item):
    """
    Extract DOI from a Zotero item by examining URL and Extra fields.
    
    Args:
        item: Zotero item dictionary
        
    Returns:
        str: DOI if found, None otherwise
    """
    if not item or 'data' not in item:
        return None
    
    doi = None
    
    # Check if DOI is directly available in the DOI field
    if 'DOI' in item['data'] and item['data']['DOI']:
        doi = item['data']['DOI']
        return doi
    
    # Check URL field for DOI
    if 'url' in item['data'] and item['data']['url']:
        url = item['data']['url']
        
        # Check for DOI in doi.org URLs
        if 'doi.org/' in url:
            doi_part = url.split('doi.org/')[-1]
            # Remove any trailing parameters or anchors
            doi = doi_part.split('#')[0].split('?')[0]
            return doi
        
        # Look for DOI pattern in URL
        doi_match = re.search(r'(10\.\d{4,}(?:\.\d+)*\/(?:(?!["&\'])\S)+)', url)
        if doi_match:
            return doi_match.group(0)
    
    # Check Extra field for DOI
    if 'extra' in item['data'] and item['data']['extra']:
        extra = item['data']['extra']
        
        # Look for lines starting with DOI:
        for line in extra.split('\n'):
            line = line.strip()
            if line.lower().startswith('doi:'):
                doi = line[4:].strip()
                return doi
        
        # Try regex for DOI pattern in Extra field
        doi_match = re.search(r'(10\.\d{4,}(?:\.\d+)*\/(?:(?!["&\'])\S)+)', extra)
        if doi_match:
            return doi_match.group(0)
    
    return doi

def format_item_text(item, zot, google_creds=None, verbose=False):
    """Format a single item for text output with proper Unicode support."""
    output = []
    # Use Unicode strings for all text content
    output.append(f"Title: {item['data'].get('title', 'Unknown')}")
    output.append(f"Type: {item['data'].get('itemType', 'Unknown')}")
    
    # Format authors
    if 'creators' in item['data'] and item['data']['creators']:
        authors = []
        for creator in item['data']['creators']:
            if 'lastName' in creator and 'firstName' in creator:
                authors.append(f"{creator['lastName']}, {creator['firstName']}")
            elif 'name' in creator:
                authors.append(creator['name'])
        if authors:
            output.append(f"Authors: {'; '.join(authors)}")
    
    if 'date' in item['data'] and item['data']['date']:
        output.append(f"Date: {item['data']['date']}")
    
    # Type-specific fields
    item_type = item['data'].get('itemType')
    if item_type == 'book':
        if 'publisher' in item['data'] and item['data']['publisher']:
            output.append(f"Publisher: {item['data']['publisher']}")
        if 'place' in item['data'] and item['data']['place']:
            output.append(f"Place: {item['data']['place']}")
        if 'ISBN' in item['data'] and item['data']['ISBN']:
            output.append(f"ISBN: {item['data']['ISBN']}")
    elif item_type == 'journalArticle':
        if 'publicationTitle' in item['data'] and item['data']['publicationTitle']:
            output.append(f"Journal: {item['data']['publicationTitle']}")
        if 'volume' in item['data'] and item['data']['volume']:
            output.append(f"Volume: {item['data']['volume']}")
        if 'issue' in item['data'] and item['data']['issue']:
            output.append(f"Issue: {item['data']['issue']}")
        if 'pages' in item['data'] and item['data']['pages']:
            output.append(f"Pages: {item['data']['pages']}")
    elif item_type == 'manuscript':
        # Add arXiv URL for manuscripts
        if 'url' in item['data'] and item['data']['url'] and 'arxiv.org' in item['data']['url']:
            output.append(f"arXiv URL: {item['data']['url']}")
        # Check if there's an arXiv ID in extra field
        if 'extra' in item['data'] and item['data']['extra']:
            extra = item['data']['extra']
            if 'arXiv:' in extra:
                for line in extra.split('\n'):
                    if line.strip().startswith('arXiv:'):
                        arxiv_id = line.strip().replace('arXiv:', '').strip()
                        output.append(f"arXiv ID: {arxiv_id}")
                        if 'url' not in item['data'] or 'arxiv.org' not in item['data']['url']:
                            output.append(f"arXiv URL: https://arxiv.org/abs/{arxiv_id}")
    
    # Extract DOI using the extract_doi function
    doi = extract_doi(item)
    if doi:
        output.append(f"DOI: {doi}")
    
    # Add attachment paths and Google Drive URLs
    attachments = get_attachment_paths(zot, item, google_creds, verbose)
    if attachments:
        output.append("Attachments:")
        for attachment in attachments:
            local_path = attachment.get('local_path', 'Unknown')
            drive_url = attachment.get('drive_url')
            
            if drive_url:
                output.append(f"  - {local_path} (Drive: {drive_url})")
            else:
                output.append(f"  - {local_path}")
    
    # Join all lines with Unicode newlines and ensure the result is Unicode
    return "\n".join(output)

def format_item_html(item, zot, google_creds=None, verbose=False):
    """Format a single item for HTML output."""
    # Use html.escape for all text content to handle Unicode properly
    
    # Start with basic item info
    html_parts = [f"<div class='item {html.escape(item['data'].get('itemType', ''))}'>"
            f"<h2>{html.escape(item['data'].get('title', 'Unknown'))}</h2>"]
    
    html_parts.append(f"<p><strong>Type:</strong> {html.escape(item['data'].get('itemType', 'Unknown'))}</p>")
    
    # Format authors
    if 'creators' in item['data'] and item['data']['creators']:
        authors = []
        for creator in item['data']['creators']:
            if 'lastName' in creator and 'firstName' in creator:
                authors.append(f"{html.escape(creator['lastName'])}, {html.escape(creator['firstName'])}")
            elif 'name' in creator:
                authors.append(html.escape(creator['name']))
        if authors:
            html_parts.append(f"<p><strong>Authors:</strong> {html.escape('; '.join(authors))}</p>")
    
    if 'date' in item['data'] and item['data']['date']:
        html_parts.append(f"<p><strong>Date:</strong> {html.escape(item['data']['date'])}</p>")
    
    # Type-specific fields
    item_type = item['data'].get('itemType')
    if item_type == 'book':
        if 'publisher' in item['data'] and item['data']['publisher']:
            html_parts.append(f"<p><strong>Publisher:</strong> {html.escape(item['data']['publisher'])}</p>")
        if 'place' in item['data'] and item['data']['place']:
            html_parts.append(f"<p><strong>Place:</strong> {html.escape(item['data']['place'])}</p>")
        if 'ISBN' in item['data'] and item['data']['ISBN']:
            html_parts.append(f"<p><strong>ISBN:</strong> {html.escape(item['data']['ISBN'])}</p>")
    elif item_type == 'journalArticle':
        if 'publicationTitle' in item['data'] and item['data']['publicationTitle']:
            html_parts.append(f"<p><strong>Journal:</strong> {html.escape(item['data']['publicationTitle'])}</p>")
        if 'volume' in item['data'] and item['data']['volume']:
            html_parts.append(f"<p><strong>Volume:</strong> {html.escape(item['data']['volume'])}</p>")
        if 'issue' in item['data'] and item['data']['issue']:
            html_parts.append(f"<p><strong>Issue:</strong> {html.escape(item['data']['issue'])}</p>")
        if 'pages' in item['data'] and item['data']['pages']:
            html_parts.append(f"<p><strong>Pages:</strong> {html.escape(item['data']['pages'])}</p>")
    elif item_type == 'manuscript':
        # Add arXiv URL for manuscripts
        if 'url' in item['data'] and item['data']['url'] and 'arxiv.org' in item['data']['url']:
            html_parts.append(f"<p><strong>arXiv URL:</strong> <a href='{html.escape(item['data']['url'])}' target='_blank'>{html.escape(item['data']['url'])}</a></p>")
        # Check if there's an arXiv ID in extra field
        if 'extra' in item['data'] and item['data']['extra']:
            extra = item['data']['extra']
            if 'arXiv:' in extra:
                for line in extra.split('\n'):
                    if line.strip().startswith('arXiv:'):
                        arxiv_id = line.strip().replace('arXiv:', '').strip()
                        html_parts.append(f"<p><strong>arXiv ID:</strong> {html.escape(arxiv_id)}</p>")
                        if 'url' not in item['data'] or 'arxiv.org' not in item['data']['url']:
                            arxiv_url = f"https://arxiv.org/abs/{arxiv_id}"
                            html_parts.append(f"<p><strong>arXiv URL:</strong> <a href='{html.escape(arxiv_url)}' target='_blank'>{html.escape(arxiv_url)}</a></p>")
    
    # Extract DOI using extract_doi function
    doi = extract_doi(item)
    if doi:
        doi_escaped = html.escape(doi)
        doi_url = f"https://doi.org/{doi_escaped}"
        html_parts.append(f"<p><strong>DOI:</strong> <a href='{doi_url}' target='_blank'>{doi_escaped}</a></p>")
    
    # Add attachment paths with Google Drive links
    attachments = get_attachment_paths(zot, item, google_creds, verbose)
    if attachments:
        html_parts.append("<p><strong>Attachments:</strong></p>")
        html_parts.append("<ul>")
        for attachment in attachments:
            local_path = html.escape(attachment.get('local_path', 'Unknown'))
            drive_url = attachment.get('drive_url')
            
            if drive_url:
                html_parts.append(f"<li>{local_path} - <a href='{html.escape(drive_url)}' target='_blank'>View on Google Drive</a></li>")
            else:
                html_parts.append(f"<li>{local_path}</li>")
        html_parts.append("</ul>")
    
    html_parts.append("</div>")
    return "\n".join(html_parts)

def generate_text_output(items, zot, collection_name=None, google_creds=None, verbose=False):
    """Generate complete text document from items with proper Unicode support."""
    
    if verbose:
        print_progress("Starting text output generation", verbose)
    
    current_date = datetime.now().strftime("%Y-%m-%d")
    title = f"Zotero Items - {current_date}".title()
    if collection_name:
        title = f"Zotero Collection: {collection_name} - {current_date}".title()
        
    header = [
        title,
        "=" * len(title),
        ""  # Empty line after header
    ]
    
    if verbose:
        print_progress(f"Preparing to format {len(items)} items simultaneously", verbose)
    
    # Helper function to format a single item
    def format_single_item(idx, item):
        try:
            item_header = f"{collection_name} #{idx+1}"
            item_content = format_item_text(item, zot, google_creds, verbose)
            return f"{item_header}\n{item_content}\n---"
        except Exception as e:
            error_msg = f"Error formatting item {idx+1}: {e}"
            print_progress(error_msg, verbose, file=sys.stderr)
            return f"{error_msg}\n---"
    
    # Process items in parallel with ThreadPoolExecutor
    formatted_items = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Create and submit all tasks
        future_to_idx = {}
        for i, item in enumerate(items):
            future = executor.submit(format_single_item, i, item)
            future_to_idx[future] = i
        
        # Process results as they complete
        completed = 0
        for future in concurrent.futures.as_completed(future_to_idx):
            completed += 1
            if verbose and (completed % 10 == 0 or completed == len(items)):
                print_progress(f"Completed {completed}/{len(items)} items", verbose)
            
            # Store results with their index for later sorting
            idx = future_to_idx[future]
            try:
                formatted_items.append((idx, future.result()))
            except Exception as e:
                error_msg = f"Error processing item {idx+1}: {e}"
                print_progress(error_msg, verbose, file=sys.stderr)
                formatted_items.append((idx, f"{error_msg}\n---"))
    
    # Sort by original index to maintain order
    formatted_items.sort(key=lambda x: x[0])
    ordered_items = [text for _, text in formatted_items]
    
    if verbose:
        print_progress("Text output generation complete", verbose)
    
    # Ensure Unicode output    
    return "\n".join(header + ordered_items)

def generate_html_header(title, notice=None):
    """
    Generate the HTML header section with styles and KaTeX support.
    
    Args:
        title (str): The title for the HTML document
        notice (str, optional): Custom notice message. If None, uses default.
    """
    # Use default notice if none is provided
    default_notice = "This document was automatically generated from a Zotero library. Items are listed for personal reference only. All references, articles, and other content remain the property of their respective copyright holders. This document is not for redistribution. Last updated on " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "."
    notice_text = notice if notice is not None else default_notice
    
    return [
        "<!DOCTYPE html>",
        "<html>",
        "<head>",
        f"<title>{title}</title>",
        "<!-- KaTeX CSS -->",
        "<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css' integrity='sha384-GvrOXuhMATgEsSwCs4smul74iXGOixntILdUW9XmUC6+HX0sLNAK3q71HotJqlAn' crossorigin='anonymous'>",
        "<!-- KaTeX JS -->",
        "<script defer src='https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js' integrity='sha384-cpW21h6RZv/phavutF+AuVYrr+dA8xD9zs6FwLpaCct6O9ctzYFfFr4dgmgccOTx' crossorigin='anonymous'></script>",
        "<!-- KaTeX auto-render extension -->",
        "<script defer src='https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js' integrity='sha384-+VBxd3r6XgURycqtZ117nYw44OOcIax56Z4dCRWbxyPt0Koah1uHoK0o4+/RRE05' crossorigin='anonymous'></script>",
        "<script>",
        "document.addEventListener('DOMContentLoaded', function() {",
        "  renderMathInElement(document.body, {",
        "    delimiters: [",
        "      {left: '$$', right: '$$', display: true},",
        "      {left: '$', right: '$', display: false}",
        "    ],",
        "    throwOnError: false",
        "  });",
        "});",
        "</script>",
        "<style>",
        "body { font-family: Arial, sans-serif; margin: 40px; }",
        ".item { margin-bottom: 30px; border-bottom: 1px solid #ccc; padding-bottom: 20px; }",
        ".item-number { font-weight: bold; color: #7f8c8d; margin-bottom: 5px; }",
        "h1 { color: #2c3e50; }",
        "h2 { color: #3498db; }",
        ".notice { font-style: italic; background-color: #f8f9fa; padding: 10px; border-left: 3px solid #3498db; margin-bottom: 20px; }",
        ".coffee-button { position: absolute; top: 20px; right: 20px; }",
        ".coffee-button img { height: 40px; border: none; }",
        ".search-container { margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; }",
        "#searchInput { width: 300px; padding: 8px; font-size: 16px; border: 1px solid #ccc; border-radius: 4px; }",
        "#searchBtn { padding: 8px 15px; background-color: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; margin-left: 10px; }",
        "#searchBtn:hover { background-color: #2980b9; }",
        "#searchCount { margin-left: 15px; font-style: italic; }",
        ".highlight { background-color: yellow; font-weight: bold; }",
        ".hidden { display: none; }",
        "</style>",
        "</head>",
        "<body>",
        "<div class='coffee-button'>",
        "<a href='https://www.buymeacoffee.com/hoanganhduc' target='_blank'>",
        "<img src='https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png' alt='Buy Me A Coffee'>",
        "</a>",
        "</div>",
        f"<h1>{title}</h1>",
        f"<div class='notice'>{notice_text}</div>"
    ]

def generate_search_container():
    """Generate the search box HTML."""
    return [
        "<div class='search-container'>",
        "<input type='text' id='searchInput' placeholder='Search within this page...' />",
        "<button id='searchBtn'>Search</button>",
        "<span id='searchCount'></span>",
        "</div>"
    ]

def format_single_item(idx, item, collection_name, zot, google_creds, verbose):
    """Format a single item for HTML output."""
    try:
        item_number = f"<div class='item-number'>{collection_name} #{idx+1}</div>"
        item_content = format_item_html(item, zot, google_creds, verbose)
        return item_number + "\n" + item_content
    except Exception as e:
        error_msg = f"Error formatting item {idx+1}: {e}"
        print_progress(error_msg, verbose, file=sys.stderr)
        return f"<div class='item-error'>{error_msg}</div>"

def generate_items_html(items, collection_name, zot, google_creds, verbose):
    """Generate HTML for all items using parallel processing."""
    if verbose:
        print_progress(f"Preparing to format {len(items)} items simultaneously", verbose)
    
    # Process items in parallel with ThreadPoolExecutor
    formatted_items = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Create and submit all tasks
        future_to_idx = {}
        for i, item in enumerate(items):
            future = executor.submit(format_single_item, i, item, collection_name, zot, google_creds, verbose)
            future_to_idx[future] = i
        
        # Process results as they complete
        completed = 0
        for future in concurrent.futures.as_completed(future_to_idx):
            completed += 1
            if verbose and (completed % 10 == 0 or completed == len(items)):
                print_progress(f"Completed {completed}/{len(items)} items", verbose)
            
            # Store results with their index for later sorting
            idx = future_to_idx[future]
            try:
                formatted_items.append((idx, future.result()))
            except Exception as e:
                error_msg = f"Error processing item {idx+1}: {e}"
                print_progress(error_msg, verbose, file=sys.stderr)
                formatted_items.append((idx, f"<div class='item-error'>{error_msg}</div>"))
    
    # Sort by original index to maintain order
    formatted_items.sort(key=lambda x: x[0])
    return [html_content for _, html_content in formatted_items]

def generate_search_script():
    """Generate the JavaScript code for search functionality."""
    return [
        "<script>",
        "document.addEventListener('DOMContentLoaded', function() {",
        "  const searchInput = document.getElementById('searchInput');",
        "  const searchBtn = document.getElementById('searchBtn');",
        "  const searchCount = document.getElementById('searchCount');",
        "  const items = document.querySelectorAll('.item');",
        "",
        "  function performSearch() {",
        "    const searchTerm = searchInput.value.toLowerCase().trim();",
        "    if (searchTerm === '') {",
        "      // Show all items if search is empty",
        "      items.forEach(item => {",
        "        item.classList.remove('hidden');",
        "        // Remove any existing highlights",
        "        const highlighted = item.querySelectorAll('.highlight');",
        "        highlighted.forEach(el => {",
        "          const parent = el.parentNode;",
        "          parent.replaceChild(document.createTextNode(el.textContent), el);",
        "          parent.normalize();",
        "        });",
        "      });",
        "      searchCount.textContent = '';",
        "      return;",
        "    }",
        "",
        "    let matchCount = 0;",
        "",
        "    // Process each item",
        "    items.forEach(item => {",
        "      const text = item.textContent.toLowerCase();",
        "      const hasMatch = text.includes(searchTerm);",
        "      ",
        "      // Show/hide based on match",
        "      if (hasMatch) {",
        "        item.classList.remove('hidden');",
        "        matchCount++;",
        "        ",
        "        // Highlight matches (only in text nodes)",
        "        highlightText(item, searchTerm);",
        "      } else {",
        "        item.classList.add('hidden');",
        "      }",
        "    });",
        "",
        "    // Update count display",
        "    searchCount.textContent = `Found ${matchCount} matching items`;",
        "  }",
        "",
        "  function highlightText(element, searchTerm) {",
        "    // Remove any existing highlights first",
        "    const highlighted = element.querySelectorAll('.highlight');",
        "    highlighted.forEach(el => {",
        "      const parent = el.parentNode;",
        "      parent.replaceChild(document.createTextNode(el.textContent), el);",
        "      parent.normalize();",
        "    });",
        "",
        "    // Function to recursively process text nodes",
        "    function processNode(node) {",
        "      // Only process text nodes",
        "      if (node.nodeType === 3) {",
        "        const text = node.nodeValue.toLowerCase();",
        "        const index = text.indexOf(searchTerm.toLowerCase());",
        "        ",
        "        // If search term found in this text node",
        "        if (index >= 0) {",
        "          const before = node.nodeValue.substring(0, index);",
        "          const match = node.nodeValue.substring(index, index + searchTerm.length);",
        "          const after = node.nodeValue.substring(index + searchTerm.length);",
        "          ",
        "          const beforeNode = document.createTextNode(before);",
        "          const matchNode = document.createElement('span');",
        "          matchNode.classList.add('highlight');",
        "          matchNode.textContent = match;",
        "          const afterNode = document.createTextNode(after);",
        "          ",
        "          const parent = node.parentNode;",
        "          parent.replaceChild(afterNode, node);",
        "          parent.insertBefore(matchNode, afterNode);",
        "          parent.insertBefore(beforeNode, matchNode);",
        "          ",
        "          // Process the 'after' part too for multiple occurrences",
        "          processNode(afterNode);",
        "        }",
        "      } else if (node.nodeType === 1 && node.childNodes && !/(script|style)/i.test(node.tagName)) {",
        "        // Process children of element nodes",
        "        Array.from(node.childNodes).forEach(child => processNode(child));",
        "      }",
        "    }",
        "",
        "    // Start processing from the item element",
        "    processNode(element);",
        "  }",
        "",
        "  // Event listeners",
        "  searchBtn.addEventListener('click', performSearch);",
        "  searchInput.addEventListener('keyup', function(event) {",
        "    if (event.key === 'Enter') {",
        "      performSearch();",
        "    }",
        "  });",
        "});",
        "</script>",
        "</body>",
        "</html>"
    ]

def generate_html_output(items, zot, collection_name=None, google_creds=None, verbose=False, notice=None):
    """Generate complete HTML document from items."""
    if verbose:
        print_progress("Starting HTML output generation", verbose)
    
    current_date = datetime.now().strftime("%Y-%m-%d")
    title = f"Zotero Items - {current_date}".title()
    if collection_name:
        title = f"Zotero Collection: {collection_name} - {current_date}".title()
    
    # Build HTML components
    html_parts = []
    html_parts.extend(generate_html_header(title, notice))  # Pass the notice parameter
    html_parts.extend(generate_search_container())
    
    # Process items
    formatted_items = generate_items_html(items, collection_name, zot, google_creds, verbose)
    html_parts.extend(formatted_items)
    
    # Add search functionality
    html_parts.extend(generate_search_script())
    
    if verbose:
        print_progress("HTML output generation complete", verbose)
    
    return "\n".join(html_parts)

def generate_pdf_output(html_content, output_file, verbose=False):
    """Generate PDF from HTML content using pdfkit."""
    if verbose:
        print_progress("Starting PDF generation...", verbose)
        html_size_kb = len(html_content) / 1024
        print_progress(f"Using pdfkit to process approximately {html_size_kb:.1f} KB of HTML content", verbose)
    
    try:
        # Configure pdfkit options
        options = {
            'quiet': not verbose,
            'encoding': "UTF-8",
        }
        pdfkit.from_string(html_content, output_file, options=options)
        
        # Get the file size of the generated PDF
        if os.path.exists(output_file):
            pdf_size_kb = os.path.getsize(output_file) / 1024
            print_progress(f"PDF successfully generated ({pdf_size_kb:.1f} KB) and saved to {output_file}", verbose)
        else:
            print_progress("PDF generation seemed to complete but output file not found", verbose, file=sys.stderr)
    
    except Exception as e:
        print_progress(f"Error generating PDF with pdfkit: {str(e)}", verbose, file=sys.stderr)
        sys.exit(1)

def display_collections(collections, output_format, output_file=None, verbose=False):
    """Display collections in the specified format."""
    if not collections:
        print("No collections found.")
        return
    
    print_progress("Displaying collections...", verbose)
    
    if output_format == 'text':
        print_progress(f"Formatting {len(collections)} collections as text", verbose)
        for i, collection in enumerate(collections):
            if verbose and (i % 10 == 0 or i == len(collections) - 1):
                print_progress(f"Processing collection {i+1} of {len(collections)}", verbose)
            print(f"Name: {collection['data']['name']}")
            print(f"Key: {collection['data']['key']}")
            print("---")
    elif output_format in ['html', 'pdf']:
        print_progress(f"Formatting {len(collections)} collections as HTML", verbose)
        html = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            "<title>Zotero Collections</title>",
            "<style>",
            "body { font-family: Arial, sans-serif; margin: 40px; }",
            ".collection { margin-bottom: 20px; }",
            "h1 { color: #2c3e50; }",
            "</style>",
            "</head>",
            "<body>",
            "<h1>Zotero Collections</h1>"
        ]
        
        for i, collection in enumerate(collections):
            if verbose and (i % 10 == 0 or i == len(collections) - 1):
                print_progress(f"Processing collection {i+1} of {len(collections)}", verbose)
            html.append("<div class='collection'>")
            html.append(f"<p><strong>Name:</strong> {collection['data']['name']}</p>")
            html.append(f"<p><strong>Key:</strong> {collection['data']['key']}</p>")
            html.append("</div>")
        
        html.extend(["</body>", "</html>"])
        html_content = "\n".join(html)
        
        if output_format == 'html':
            if output_file:
                print_progress(f"Saving HTML output to {output_file}", verbose)
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                print(f"HTML output saved to {output_file}")
            else:
                print_progress("Displaying HTML output", verbose)
                print(html_content)
        else:  # pdf
            if not output_file:
                output_file = "zotero_collections.pdf"
            print_progress(f"Generating PDF output to {output_file}", verbose)
            generate_pdf_output(html_content, output_file, verbose)
            print(f"PDF output saved to {output_file}")
    
    print_progress("Collection display complete", verbose)

def display_items(items, output_format, output_file=None, collection_name=None, zot=None, verbose=False, google_creds=None, notice=None):
    """Display items in the specified format."""
    if not items:
        print("No items found.")
        return
    
    print_progress("Displaying items...", verbose)
    
    if output_format == 'text':
        print_progress("Generating text output...", verbose)
        text_content = generate_text_output(items, zot, collection_name, google_creds, verbose)
        if output_file:
            print_progress(f"Saving text output to {output_file}", verbose)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(text_content)
            print(f"Text output saved to {output_file}")
        else:
            print_progress("Displaying text output to console", verbose)
            print(text_content)
    elif output_format == 'html':
        print_progress("Generating HTML output...", verbose)
        html_content = generate_html_output(items, zot, collection_name, google_creds, verbose, notice)
        if output_file:
            print_progress(f"Saving HTML output to {output_file}", verbose)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"HTML output saved to {output_file}")
        else:
            print_progress("Displaying HTML output to console", verbose)
            print(html_content)
    elif output_format == 'pdf':
        print_progress("Generating PDF output...", verbose)
        html_content = generate_html_output(items, zot, collection_name, google_creds, verbose, notice)
        if not output_file:
            output_file = "zotero_items.pdf"
            print_progress(f"No output file specified, using default: {output_file}", verbose)
        generate_pdf_output(html_content, output_file, verbose)
        print(f"PDF output saved to {output_file}")
    
    print_progress("Item display complete", verbose)

def search_sqlite_db(sqlite_path, queries, item_type, max_results, verbose):
    """
    Search for papers in a Zotero SQLite database.
    
    Args:
        sqlite_path (str): Path to the SQLite database file
        queries (list): List of search terms
        item_type (str, optional): Filter by item type
        max_results (int): Maximum results per query
        verbose (bool): Whether to display progress messages
        
    Returns:
        list: List of matching Zotero items
    """
    results = []
    seen_keys = set()
    try:
        conn = sqlite3.connect(sqlite_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        for query in queries:
            sql = """
                SELECT items.itemID, items.key, itemType.typeName, itemDataValues.value as title
                FROM items
                JOIN itemTypes as itemType ON items.itemTypeID = itemType.itemTypeID
                JOIN itemData ON items.itemID = itemData.itemID
                JOIN fields ON itemData.fieldID = fields.fieldID
                JOIN itemDataValues ON itemData.valueID = itemDataValues.valueID
                WHERE fields.fieldName = 'title' AND itemDataValues.value LIKE ?
                AND items.itemID NOT IN (SELECT itemID FROM deletedItems)
                AND itemType.typeName NOT IN ('note', 'attachment')
            """
            params = [f"%{query}%"]
            if item_type:
                sql += " AND itemType.typeName = ?"
                params.append(item_type)
            sql += " LIMIT ?"
            params.append(max_results)
            if verbose:
                print_progress(f"Searching local zotero.sqlite for '{query}'...", verbose)
            for row in cur.execute(sql, params):
                key = row['key']
                if key not in seen_keys:
                    results.append({
                        'key': key,
                        'data': {
                            'title': row['title'],
                            'itemType': row['typeName']
                        }
                    })
                    seen_keys.add(key)
        conn.close()
    except Exception as e:
        print_progress(f"Error searching SQLite database: {e}", verbose, file=sys.stderr)
    return results

def search_local_sqlite(queries, item_type, max_results, verbose):
    """
    Find and search local Zotero SQLite databases.
    
    Args:
        queries (list): List of search terms
        item_type (str, optional): Filter by item type
        max_results (int): Maximum results per query
        verbose (bool): Whether to display progress messages
        
    Returns:
        list: List of matching Zotero items
    """
    local_paths = [
        os.path.expanduser("~/Zotero/zotero.sqlite"),
        os.path.expanduser("~/Library/Application Support/Zotero/zotero.sqlite"),
        os.path.expanduser("~/.zotero/zotero.sqlite"),
        "./zotero.sqlite"
    ]
    
    for path in local_paths:
        if os.path.exists(path):
            if verbose:
                print_progress(f"Found local zotero.sqlite at {path}", verbose)
            results = search_sqlite_db(path, queries, item_type, max_results, verbose)
            if results:
                return results
            break  # Only use first found
    
    return []

def search_drive_sqlite(queries, item_type, max_results, verbose):
    """
    Download and search Zotero SQLite database from Google Drive.
    
    Args:
        queries (list): List of search terms
        item_type (str, optional): Filter by item type
        max_results (int): Maximum results per query
        verbose (bool): Whether to display progress messages
        
    Returns:
        list: List of matching Zotero items
    """
    google_creds = globals().get('google_creds', None)
    if not google_creds:
        if verbose:
            print_progress("No Google Drive credentials available for search", verbose)
        return []
        
    drive_url = get_drive_url_by_filename(google_creds, "zotero.sqlite", exact_match=True, verbose=verbose)
    if not drive_url:
        return []
        
    file_id = extract_file_id_from_drive_url(drive_url)
    if not file_id:
        return []
        
    drive_service = build('drive', 'v3', credentials=google_creds)
    temp_path = download_file_from_drive(drive_service, file_id, verbose=verbose)
    
    if not temp_path or not os.path.exists(temp_path):
        return []
        
    results = search_sqlite_db(temp_path, queries, item_type, max_results, verbose)
    
    # Clean up temp file
    try:
        os.remove(temp_path)
    except Exception:
        pass
        
    return results

def search_zotero_api(zot, queries, item_type, max_results, verbose):
    """
    Search for papers using the Zotero API.
    
    Args:
        zot: Zotero API client instance
        queries (list): List of search terms
        item_type (str, optional): Filter by item type
        max_results (int): Maximum results per query
        verbose (bool): Whether to display progress messages
        
    Returns:
        list: List of matching Zotero items
    """
    if verbose:
        print_progress("Searching online Zotero library via API...", verbose)
        
    all_results = []
    all_item_keys = set()
    
    try:
        for query in queries:
            if verbose and len(queries) > 1:
                print_progress(f"Searching for title: '{query}'", verbose)
                
            params = {'q': query, 'limit': max_results}
            if item_type:
                params['itemType'] = item_type
                
            results = zot.items(params=params)
            
            for item in results:
                if (item['data'].get('itemType') not in ['attachment', 'note'] and
                    item['key'] not in all_item_keys):
                    all_results.append(item)
                    all_item_keys.add(item['key'])
                    
        if verbose:
            print_progress(f"Found {len(all_results)} unique matching papers", verbose)
            
        return all_results
        
    except Exception as e:
        print_progress(f"Error searching online: {e}", verbose, file=sys.stderr)
        return []

def find_papers_by_title(zot, title_query, item_type=None, max_results=10, verbose=False):
    """
    Search for papers in Zotero library by title or partial title.
    Priority:
      1. Search local zotero.sqlite if available.
      2. If not, search for zotero.sqlite in Google Drive, download and search.
      3. If not, search online Zotero library via API.
    
    Args:
        zot: Zotero API client instance
        title_query (str or list): Title(s) or partial title(s) to search for
        item_type (str, optional): Filter by item type (e.g., 'journalArticle')
        max_results (int): Maximum number of results to return per search term
        verbose (bool): Whether to display progress messages
        
    Returns:
        list: List of matching Zotero items
    """
    # Prepare queries as list
    if isinstance(title_query, str):
        title_queries = [title_query]
    else:
        title_queries = title_query

    # Try each search method in order
    for search_method in [search_local_sqlite, search_drive_sqlite, 
                         lambda q, it, mr, v: search_zotero_api(zot, q, it, mr, v)]:
        results = search_method(title_queries, item_type, max_results, verbose)
        if results:
            return results
            
    return []

def download_file_from_drive(drive_service, file_id, output_path=None, verbose=False):
    """
    Download a file from Google Drive by its ID.
    
    Args:
        drive_service: Google Drive service instance
        file_id (str): ID of the file to download
        output_path (str, optional): Path where to save the file (if None, uses temp file)
        verbose (bool): Whether to display progress messages
        
    Returns:
        str: Path to the downloaded file, or None if download failed
    """
    try:
        if verbose:
            print_progress(f"Downloading file ID: {file_id}...", verbose)
            
        # Get file metadata to get the filename
        file_metadata = drive_service.files().get(fileId=file_id).execute()
        file_name = file_metadata.get('name', 'unknown_file')
        
        # Create a BytesIO object for the download
        request = drive_service.files().get_media(fileId=file_id)
        file_buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(file_buffer, request)
        
        # Download the file
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if verbose and status.progress() > 0:
                print_progress(f"Download progress: {int(status.progress() * 100)}%", verbose)
        
        # Determine output path
        if output_path is None:
            suffix = '.' + file_name.split('.')[-1] if '.' in file_name else ''
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            output_path = temp_file.name
            temp_file.close()
        
        # Write the file content to disk
        with open(output_path, 'wb') as f:
            f.write(file_buffer.getvalue())
            
        if verbose:
            print_progress(f"File downloaded to {output_path}", verbose)
            
        return output_path
        
    except Exception as e:
        print_progress(f"Error downloading file: {e}", verbose, file=sys.stderr)
        return None

def extract_file_id_from_drive_url(drive_url):
    """
    Extract the file ID from a Google Drive URL.
    
    Args:
        drive_url (str): Google Drive URL
        
    Returns:
        str: File ID, or None if not found
    """
    if not drive_url:
        return None
        
    # Format: https://drive.google.com/file/d/{fileid}/view
    file_match = re.search(r'/file/d/([^/]+)', drive_url)
    if file_match:
        return file_match.group(1)
        
    # Format: https://drive.google.com/open?id={fileid}
    id_match = re.search(r'[?&]id=([^&]+)', drive_url)
    if id_match:
        return id_match.group(1)
    
    # Format: https://docs.google.com/document/d/{fileid}/edit
    docs_match = re.search(r'/document/d/([^/]+)', drive_url)
    if docs_match:
        return docs_match.group(1)
    
    return None

def send_email_with_attachments(username, app_password, to_emails, subject, body, attachments=None, verbose=False):
    """
    Send an email with attachments using Gmail.
    
    Args:
        username (str): Gmail username (email address)
        app_password (str): Gmail app password
        to_emails (list): List of recipient email addresses
        subject (str): Email subject
        body (str): Email body
        attachments (list, optional): List of file paths to attach
        verbose (bool): Whether to display progress messages
        
    Returns:
        bool: True if email sent successfully, False otherwise
    """
    # Convert single email to list if needed
    if isinstance(to_emails, str):
        to_emails = [to_emails]
    
    if verbose:
        print_progress(f"Preparing email to {', '.join(to_emails)}", verbose)
    
    try:
        # Create a multipart message
        msg = MIMEMultipart()
        msg['From'] = username
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject
        
        # Add body to email
        msg.attach(MIMEText(body, 'plain'))
        
        # Add attachments
        if attachments:
            for file_path in attachments:
                if os.path.exists(file_path):
                    try:
                        with open(file_path, 'rb') as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                            
                        # Encode file in ASCII characters and add header
                        encoders.encode_base64(part)
                        filename = os.path.basename(file_path)
                        part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                        msg.attach(part)
                        
                        if verbose:
                            print_progress(f"Attached: {filename}", verbose)
                    except Exception as e:
                        print_progress(f"Error attaching {file_path}: {e}", verbose, file=sys.stderr)
        
        # Connect to Gmail's SMTP server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(username, app_password)
        
        # Send email
        text = msg.as_string()
        server.sendmail(username, to_emails, text)
        server.quit()
        
        if verbose:
            print_progress("Email sent successfully", verbose)
        return True
        
    except Exception as e:
        print_progress(f"Error sending email: {e}", verbose, file=sys.stderr)
        return False

def send_paper_by_email(zot, google_creds, title_query, gmail_username, gmail_app_password, recipients, 
                        subject=None, body=None, delete_after_send=True, verbose=False, max_papers=5):
    """
    Search for papers by title(s), download attachments from Google Drive, and email them.
    If total file size exceeds 20MB, sends Google Drive links instead of attachments.
    
    Args:
        zot: Zotero API client instance
        google_creds: Google API credentials object
        title_query (str or list): Title(s) or partial title(s) of the paper(s)
        gmail_username (str): Gmail username/email
        gmail_app_password (str): Gmail app password
        recipients (list or str): Email address(es) to send to
        subject (str, optional): Email subject (default: constructed from paper title)
        body (str, optional): Email body (default: constructed from paper details)
        delete_after_send (bool): Whether to delete downloaded files after sending
        verbose (bool): Whether to display progress messages
        max_papers (int): Maximum number of papers to include
        
    Returns:
        bool: True if any papers were found and email sent successfully
    """
    # Convert inputs to lists
    if isinstance(title_query, str):
        title_queries = [title_query]
    else:
        title_queries = title_query
        
    if isinstance(recipients, str):
        recipient_list = [r.strip() for r in recipients.split(',')]
    else:
        recipient_list = recipients
    
    if verbose:
        query_desc = ", ".join([f"'{q}'" for q in title_queries])
        print_progress(f"Searching for papers: {query_desc}...", verbose)
    
    # Find the papers in Zotero
    papers = find_papers_by_title(zot, title_queries, verbose=verbose)
    if not papers:
        print_progress(f"No papers found matching titles", verbose, file=sys.stderr)
        return False
    
    # Limit number of papers
    if len(papers) > max_papers:
        if verbose:
            print_progress(f"Found {len(papers)} papers, limiting to {max_papers}", verbose)
        papers = papers[:max_papers]
    
    if not google_creds:
        print_progress("No Google credentials provided", verbose, file=sys.stderr)
        return False
    
    # Set up Google Drive service
    drive_service = build('drive', 'v3', credentials=google_creds)
    
    # Process papers and download attachments
    downloaded_files = []
    paper_info_list = []
    total_size = 0
    
    for paper in papers:
        paper_title = paper['data'].get('title', 'Unknown')
        if verbose:
            print_progress(f"Processing '{paper_title}'", verbose)
        
        # Get paper metadata
        authors = []
        if 'creators' in paper['data']:
            for creator in paper['data']['creators']:
                if 'lastName' in creator and 'firstName' in creator:
                    authors.append(f"{creator['lastName']}, {creator['firstName']}")
                elif 'name' in creator:
                    authors.append(creator['name'])
        
        # Get DOI
        doi = extract_doi(paper)
        
        # Get attachments
        attachments = get_attachment_paths(zot, paper, google_creds, verbose)
        
        paper_info = {
            'title': paper_title,
            'authors': authors,
            'doi': doi,
            'attachments': []
        }
        
        # Process each attachment
        for attachment_info in attachments:
            drive_url = attachment_info.get('drive_url')
            if not drive_url:
                continue
                
            file_id = extract_file_id_from_drive_url(drive_url)
            if not file_id:
                continue
            
            local_path = attachment_info.get('local_path', 'Unknown')
            filename = os.path.basename(local_path)
            
            # Add attachment info
            attachment_item = {
                'filename': filename,
                'drive_url': drive_url,
                'path': None,
                'size': 0,
                'success': False
            }
            
            # Download the file
            try:
                downloaded_file = download_file_from_drive(drive_service, file_id, verbose=verbose)
                if downloaded_file:
                    downloaded_files.append(downloaded_file)
                    attachment_item['path'] = downloaded_file
                    attachment_item['size'] = os.path.getsize(downloaded_file)
                    attachment_item['success'] = True
                    total_size += attachment_item['size']
            except Exception as e:
                if verbose:
                    print_progress(f"Error downloading {filename}: {e}", verbose, file=sys.stderr)
            
            paper_info['attachments'].append(attachment_item)
        
        if paper_info['attachments']:
            paper_info_list.append(paper_info)
    
    # If no attachments were found, exit
    if not paper_info_list:
        print_progress("No attachments found", verbose, file=sys.stderr)
        return False
    
    # Check size limits (20MB)
    size_limit = 20 * 1024 * 1024
    exceed_size_limit = total_size > size_limit
    
    # Prepare email content
    if not subject:
        if len(paper_info_list) == 1:
            subject = f"Paper: {paper_info_list[0]['title']}"
        else:
            subject = f"Papers: {paper_info_list[0]['title']} and {len(paper_info_list)-1} more"
    
    if not body:
        if exceed_size_limit:
            body_lines = ["Here are links to the requested paper(s) (exceeds 20MB email limit):"]
        else:
            body_lines = ["Here are the requested paper(s):"]
            
        for paper_info in paper_info_list:
            body_lines.append(f"\n## {paper_info['title']}")
            
            if paper_info.get('authors'):
                body_lines.append(f"Authors: {'; '.join(paper_info['authors'])}")
            
            if paper_info.get('doi'):
                body_lines.append(f"DOI: {paper_info['doi']} (https://doi.org/{paper_info['doi']})")
            
            body_lines.append("Attachments:")
            for attachment in paper_info['attachments']:
                if exceed_size_limit or not attachment.get('success'):
                    drive_url = attachment.get('drive_url', 'No link available')
                    body_lines.append(f"- {attachment['filename']} - {drive_url}")
                else:
                    body_lines.append(f"- {attachment['filename']}")
        
        body_lines.append("\n\nThis email was sent automatically via the <a href='https://hoanganhduc.github.io/library/zotero/list-zotero-collection.py'>list-zotero-collection.py</a> script.")
        body = "\n".join(body_lines)
    
    # Send email
    if exceed_size_limit:
        # Send links only
        email_result = send_email_with_attachments(
            gmail_username, gmail_app_password, recipient_list,
            subject, body, attachments=None, verbose=verbose
        )
    else:
        # Send with attachments
        valid_attachment_paths = [
            att['path'] for paper in paper_info_list 
            for att in paper['attachments'] 
            if att.get('success') and att.get('path')
        ]
        
        email_result = send_email_with_attachments(
            gmail_username, gmail_app_password, recipient_list,
            subject, body, valid_attachment_paths, verbose=verbose
        )
    
    # Clean up
    if delete_after_send and downloaded_files:
        if verbose:
            print_progress("Cleaning up temporary files...", verbose)
        for file_path in downloaded_files:
            try:
                os.remove(file_path)
            except Exception:
                pass
    
    return email_result

def parse_arguments():
    """Parse and return command line arguments."""
    parser = argparse.ArgumentParser(description='List items from a Zotero collection or send papers by email.')
    
    # Core Zotero authentication options
    parser.add_argument('-k', '--api-key', required=True, help='Your Zotero API key')
    parser.add_argument('-t', '--library-type', choices=['user', 'group'], default='user',
                        help='Type of library (user or group)')
    parser.add_argument('-l', '--library-id', required=True, help='Your user or group ID')
    
    # Item/Collection selection options
    parser.add_argument('-c', '--collection', help='Collection ID (optional)')
    parser.add_argument('-i', '--item-type', help='Filter by item type (e.g., book, journalArticle)')
    parser.add_argument('-L', '--list-collections', action='store_true', 
                        help='List all collections instead of items')
    
    # Output options
    parser.add_argument('-o', '--output-format', choices=['text', 'html', 'pdf'], default='text',
                        help='Output format (default: text)')
    parser.add_argument('-f', '--output-file', help='Output file name (for html and pdf)')
    parser.add_argument('-v', '--verbose', action='store_true', 
                        help='Display progress information during execution')
    parser.add_argument('-n', '--notice', 
                        help='Custom copyright notice message for HTML/PDF output')
    
    # Google Drive integration
    parser.add_argument('-s', '--service-account-file', 
                        help='Path to Google service account JSON file or JSON string')
    
    # Search functionality
    parser.add_argument('-S', '--search', action='append', dest='searches',
                        help='Search for papers by title (can be specified multiple times)')
    
    # Email functionality
    email_group = parser.add_argument_group('Email options', 'Options for sending papers by email')
    email_group.add_argument('-e', '--send-email', action='store_true', help='Send papers by email (requires --search)')
    email_group.add_argument('-u', '--gmail-username', help='Gmail username/email for sending papers')
    email_group.add_argument('-p', '--gmail-app-password', help='Gmail app password for sending papers')
    email_group.add_argument('-R', '--recipient', action='append', dest='recipients',
                        help='Email recipient (can be specified multiple times)')
    email_group.add_argument('-j', '--email-subject', help='Email subject (optional, defaults to paper title)')
    email_group.add_argument('-b', '--email-body', help='Email body (optional)')
    email_group.add_argument('-K', '--keep-files', action='store_true', 
                            help='Keep downloaded files after sending email (default: delete them)')
    
    return parser.parse_args()

def main():
    # For functions that access google_creds from global scope
    global google_creds
    google_creds = None
    
    try:
        # Parse arguments
        args = parse_arguments()
        
        # Connect to Zotero
        print_progress("Connecting to Zotero...", args.verbose)
        zot = connect_to_zotero(args.library_id, args.library_type, args.api_key)
        print_progress("Connection established successfully", args.verbose)

        # Set up Google Drive credentials if provided
        if args.service_account_file:
            print_progress("Authenticating with Google Drive...", args.verbose)
            google_creds = authenticate_google_drive(args.service_account_file)
            
            # Test Google Drive access
            if google_creds:
                success, message = test_google_drive_access(google_creds, verbose=args.verbose)
                if success:
                    print_progress("Google Drive access verified", args.verbose)
                else:
                    print_progress("Google Drive access failed", args.verbose, file=sys.stderr)
            else:
                print_progress("Google Drive authentication failed", args.verbose, file=sys.stderr)
        
        # Handle email sending if requested
        if args.send_email:
            if not args.searches:
                print_progress("Error: --search parameter is required", args.verbose, file=sys.stderr)
                sys.exit(1)
                
            if not args.gmail_username or not args.gmail_app_password or not args.recipients:
                print_progress("Error: Gmail credentials and recipients required", args.verbose, file=sys.stderr)
                sys.exit(1)
                
            result = send_paper_by_email(
                zot,
                google_creds,
                args.searches,
                args.gmail_username,
                args.gmail_app_password,
                args.recipients,
                args.email_subject,
                args.email_body,
                not args.keep_files,
                args.verbose
            )
            
            if result:
                print_progress("Email sent successfully", args.verbose)
            else:
                print_progress("Failed to send email", args.verbose, file=sys.stderr)
                sys.exit(1)
                
        # Handle regular search (without email)
        elif args.searches:
            print_progress(f"Searching for papers...", args.verbose)
            papers = find_papers_by_title(zot, args.searches, args.item_type, verbose=args.verbose)
            if papers:
                print_progress(f"Found {len(papers)} papers", args.verbose)
                search_desc = f"Search results for {', '.join(args.searches)}"
                display_items(papers, args.output_format, args.output_file, 
                             search_desc, zot, 
                             args.verbose, google_creds, args.notice)
            else:
                print_progress("No papers found matching your search", args.verbose, file=sys.stderr)
                sys.exit(1)
                
        # List collections or items
        elif args.list_collections:
            print_progress("Fetching collections...", args.verbose)
            collections = list_collections(zot)
            display_collections(collections, args.output_format, args.output_file, args.verbose)
        else:
            print_progress("Fetching items...", args.verbose)
            items = get_items(zot, args.collection, args.item_type, args.verbose)
            
            # Get collection name if a collection ID was provided
            collection_name = None
            if args.collection:
                try:
                    collection = zot.collection(args.collection)
                    collection_name = collection.get('data', {}).get('name')
                except Exception:
                    pass
                
            display_items(items, args.output_format, args.output_file, 
                         collection_name, zot, args.verbose, google_creds, args.notice)
    
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()