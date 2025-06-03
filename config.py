from pathlib import Path

CONFIG = {
    "LIBRARY_ID":      "0",
    "LIBRARY_TYPE":    "user",        # 'user' of 'group'
    "STORAGE_DIR":     Path("~/Zotero/storage/").expanduser(),
    "OUTPUT_DOCX":     Path("zotero_library_export.docx"),
    "EMBED_METADATA":  True,
    "MAX_IMG_WIDTH":   6.0,   # inch
    "VERBOSE_LOGGING": False,  # Enable/disable verbose debug logs
}
