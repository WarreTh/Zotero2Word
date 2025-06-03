import sys
import warnings
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional
from pyzotero import zotero
from bs4 import BeautifulSoup
from collections import defaultdict
from collections.abc import Mapping
import os

def safe_get(d, key, default=None):
    if isinstance(d, dict):
        return d.get(key, default)
    return default

class ZItem:
    """Represents a Zotero item (entry or standalone note)."""
    def __init__(self, meta: dict):
        self.key: str = str(safe_get(meta, "key", ""))
        self.meta: dict = meta
        data = safe_get(meta, "data", {})
        self.item_type: str = str(safe_get(data, "itemType", ""))
        self.title: str = str(safe_get(data, "title", ""))
        self.date_added: str = str(safe_get(data, "dateAdded", ""))
        self.child_notes: List[str] = []
        self.snapshots: List[Path] = []
        self.standalone_note_content: Optional[str] = str(safe_get(data, "note", "")) if self.item_type == "note" else None
        if self.item_type == "note" and not self.title:
            soup = BeautifulSoup(self.standalone_note_content or "", "html.parser")
            first_meaningful_text = ""
            for i in range(1, 7):
                header = soup.find(f'h{i}')
                if header and header.get_text(strip=True):
                    first_meaningful_text = header.get_text(strip=True)
                    break
            if not first_meaningful_text:
                first_meaningful_text = soup.get_text(separator="\n", strip=True).split("\n")[0]
            self.title = (first_meaningful_text[:70] + "..." if len(first_meaningful_text) > 70 else first_meaningful_text) or "(Untitled Note)"
        elif not self.title:
            self.title = "(Untitled Item)"
        self.creators: List[dict] = list(safe_get(data, "creators", []) or [])
        self.date: str = str(safe_get(data, "date", ""))
        self.tags: List[str] = [str(safe_get(t, "tag", "")) for t in (safe_get(data, "tags", []) or []) if safe_get(t, "tag", None) is not None]
    @staticmethod
    def is_note_empty_html(note_html: Optional[str]) -> bool:
        if not note_html:
            return True
        return not BeautifulSoup(note_html, "html.parser").get_text(strip=True)
    def get_displayable_notes(self) -> List[str]:
        if self.item_type == "note":
            return [self.standalone_note_content] if self.standalone_note_content and not self.is_note_empty_html(self.standalone_note_content) else []
        return [cn for cn in self.child_notes if cn and not self.is_note_empty_html(cn)]
    def has_displayable_notes(self) -> bool:
        return bool(self.get_displayable_notes())

def connect_local(CONFIG) -> zotero.Zotero:
    storage_dir = safe_get(CONFIG, "STORAGE_DIR")
    if not storage_dir or not storage_dir.exists() or not storage_dir.is_dir():
        print(f"❌ Storage directory not found or not a directory: {storage_dir}")
        print("Please check 'STORAGE_DIR' in your config.py.")
        sys.exit(1)
    try:
        zot_instance = zotero.Zotero(safe_get(CONFIG, "LIBRARY_ID"), safe_get(CONFIG, "LIBRARY_TYPE"), api_key=None, local=True)
        zot_instance.top(limit=1)
        return zot_instance
    except Exception as e:
        print(f"❌ Could not connect to local Zotero: {e}")
        sys.exit(1)

def is_image_file(filepath):
    """Check if a file is an image by extension."""
    image_exts = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".webp"]
    if not filepath:
        return False
    return str(filepath).lower().endswith(tuple(image_exts))

def get_attachment_path(attachment_meta: dict, CONFIG) -> Optional[str]:
    """
    Resolves the local file path for a Zotero attachment (image, PDF, HTML, etc).
    Returns the file path as a string, or None if not found.
    """
    data = safe_get(attachment_meta, "data", {})
    link_mode = safe_get(data, "linkMode")
    key = safe_get(data, "key")
    filename = safe_get(data, "filename")
    storage_dir = str(safe_get(CONFIG, "STORAGE_DIR"))
    # Error checking for storage_dir
    if not storage_dir or not os.path.isdir(storage_dir):
        warnings.warn(f"[ERROR] STORAGE_DIR is not set or does not exist: {storage_dir}")
        return None
    # For imported files (local attachments)
    if (link_mode == 1 or link_mode == "imported_file" or link_mode == "imported_url") and key and storage_dir:
        folder = os.path.join(storage_dir, key)
        if filename:
            path = os.path.join(folder, filename)
            if os.path.exists(path):
                return path
            else:
                warnings.warn(f"[ERROR] Attachment file does not exist: {path}")
        # If filename is missing, use the first file in the folder
        if os.path.isdir(folder):
            files = os.listdir(folder)
            if files:
                return os.path.join(folder, files[0])
            else:
                warnings.warn(f"[ERROR] No files found in expected attachment folder: {folder}")
        else:
            warnings.warn(f"[ERROR] Expected attachment folder does not exist: {folder}")
    # For linked files (rare)
    raw_path = safe_get(data, "path")
    if raw_path:
        if raw_path.startswith("storage:"):
            path = os.path.join(storage_dir, raw_path.replace("storage:", "", 1))
            if os.path.exists(path):
                return path
            else:
                warnings.warn(f"[ERROR] Linked storage path does not exist: {path}")
        elif os.path.isabs(raw_path):
            if os.path.exists(raw_path):
                return raw_path
            else:
                warnings.warn(f"[ERROR] Absolute linked path does not exist: {raw_path}")
        else:
            warnings.warn(f"[ERROR] Could not resolve attachment path: {raw_path}")
    return None

def populate_item_children(z_item: ZItem, zot_instance: zotero.Zotero, CONFIG):
    if z_item.item_type == "note":
        return
    image_exts = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff']
    z_item.meta['attachments'] = []  # Ensure this field exists
    verbose = CONFIG.get("VERBOSE_LOGGING", True)
    try:
        children_meta = zot_instance.children(z_item.key, limit=None)
        for child_meta in children_meta:
            child_data = safe_get(child_meta, "data", {})
            if verbose:
                print(f"[Zotero2Word] Child: itemType={child_data.get('itemType')}, title={child_data.get('title')}, path={child_data.get('path')}, url={child_data.get('url')}", file=sys.stderr)
            if safe_get(child_data, "itemType", "") == "note":
                note_content_html = safe_get(child_data, "note", "")
                if note_content_html and not ZItem.is_note_empty_html(note_content_html):
                    z_item.child_notes.append(note_content_html)
            elif safe_get(child_data, "itemType", "") == "attachment":
                path = get_attachment_path(child_meta if isinstance(child_meta, dict) else {}, CONFIG)
                if path and os.path.exists(path) and path.lower().endswith(".html"):
                    z_item.snapshots.append(Path(path))
                # --- Add image attachments to meta['attachments'] ---
                if path and os.path.exists(path) and is_image_file(path):
                    z_item.meta['attachments'].append(child_meta)
                    if verbose:
                        print(f"[Zotero2Word] Found image attachment: {path}", file=sys.stderr)
                elif 'url' in child_data:
                    url = child_data['url']
                    ext = os.path.splitext(url)[1].lower()
                    if ext in image_exts:
                        z_item.meta['attachments'].append(child_meta)
                        if verbose:
                            print(f"[Zotero2Word] Found image URL attachment: {url}", file=sys.stderr)
    except Exception as e:
        warnings.warn(f"Error fetching children for item {z_item.key} ('{z_item.title}'): {e}")

def build_zotero_item_tree(zot_instance: zotero.Zotero, CONFIG) -> Tuple[Dict[Tuple[str, ...], List[ZItem]], Dict[str, ZItem]]:
    items_map: Dict[str, ZItem] = {}
    all_items_meta = zot_instance.everything(zot_instance.items(itemType="-attachment"))
    for item_meta in all_items_meta:
        data = safe_get(item_meta, "data", {})
        if safe_get(data, "itemType", "") == "note" and safe_get(data, "parentItem", None):
            continue
        z_item = ZItem(item_meta)
        items_map[z_item.key] = z_item
    for z_item in items_map.values():
        populate_item_children(z_item, zot_instance, CONFIG)
    collections_data: Dict[Tuple[str, ...], List[ZItem]] = defaultdict(list)
    processed_item_keys_in_collections: set[str] = set()
    all_zotero_collections = zot_instance.collections()
    collections_by_key_map: Dict[str, Dict[str, Any]] = {str(safe_get(c, "key", "")): c if isinstance(c, dict) else {} for c in all_zotero_collections if safe_get(c, "key", None) is not None}
    parent_to_child_collection_keys_map: Dict[Optional[str], List[str]] = defaultdict(list)
    root_collection_keys: List[str] = []
    for c_meta in all_zotero_collections:
        data = safe_get(c_meta, "data", {})
        parent_key = safe_get(data, "parentCollection", None)
        this_key = str(safe_get(c_meta, "key", ""))
        if not parent_key:
            root_collection_keys.append(this_key)
        else:
            parent_to_child_collection_keys_map[parent_key].append(this_key)
    def walk_collections_recursively(collection_key: str, current_path_tuple: Tuple[str, ...]):
        collection_meta = collections_by_key_map.get(collection_key, {})
        data = safe_get(collection_meta, "data", {})
        collection_name = str(safe_get(data, "name", ""))
        if not collection_name:
            return
        new_path_tuple = current_path_tuple + (collection_name,)
        try:
            collection_item_refs = zot_instance.collection_items(collection_key, itemType="-attachment -note", limit=None)
            for item_ref in collection_item_refs:
                item_key = str(safe_get(item_ref, "key", ""))
                z_item_obj = items_map.get(item_key)
                if z_item_obj:
                    collections_data[new_path_tuple].append(z_item_obj)
                    processed_item_keys_in_collections.add(item_key)
                else:
                    warnings.warn(f"Item {item_key} in collection '{collection_name}' not found in main items_map. Skipping.")
        except Exception as e:
            warnings.warn(f"Error fetching items for collection {collection_key} ('{collection_name}'): {e}")
        for sub_collection_key in parent_to_child_collection_keys_map.get(collection_key, []):
            walk_collections_recursively(sub_collection_key, new_path_tuple)
    for root_key in root_collection_keys:
        walk_collections_recursively(root_key, tuple())
    unfiled_items_list: List[ZItem] = [
        z_item_obj for item_key, z_item_obj in items_map.items() 
        if item_key not in processed_item_keys_in_collections
    ]
    if unfiled_items_list:
        collections_data[("Unfiled Items",)].extend(unfiled_items_list)
    return collections_data, items_map
