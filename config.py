from pathlib import Path

CONFIG = {
    "LIBRARY_ID":      "0",
    "LIBRARY_TYPE":    "user",        # 'user' of 'group'
    "STORAGE_DIR":     Path("~/Zotero/storage/").expanduser(),
    "OUTPUT_DOCX":     Path("zotero_library_export.docx"),
    "EMBED_METADATA":  True,
    "MAX_IMG_WIDTH":   6.0,   # inch
    "VERBOSE_LOGGING": False,  # Enable/disable verbose debug logs
    "ENABLE_IMAGES":   True,  # Set to False to disable adding/parsing/downloading images
    "ENABLE_WEBPAGES": True,  # Set to False to disable adding/parsing HTML snapshots/webpages
    # Styling configuration for document output
    "STYLING": {
        # Main font for normal text
        "FONT_NAME": "Calibri",
        # Normal text size (in points)
        "NORMAL_TEXT_SIZE": 11,
        # Heading sizes (in points)
        "HEADING_SIZES": {
            "h1": 16,
            "h2": 14,
            "h3": 12,
            "h4": 11,
            "h5": 10,
            "h6": 10
        },
        # Divider (horizontal rule) thickness (in points)
        "DIVIDER_SIZE": 2,
        # Code block font and size
        "CODE_BLOCK_FONT_NAME": "Courier New",
        "CODE_BLOCK_FONT_SIZE": 10,
        # Code block background color (hex)
        "CODE_BLOCK_BG_COLOR": "F0F0F0",
        # Text color (hex)
        "TEXT_COLOR": "000000",
        # Hyperlink color (hex)
        "HYPERLINK_COLOR": "0563C1",
        # Small/metadata text size (in points)
        "SMALL_TEXT_SIZE": 8
    }
}
