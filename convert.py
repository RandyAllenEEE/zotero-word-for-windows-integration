import win32com.client
import re
import json
import sys
import os

def convert_citations():
    print("Connecting to Word...")
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception as e:
        print(f"Error connecting to Word: {e}")
        print("Make sure Word is open and a document is active.")
        return

    try:
        doc = word.ActiveDocument
    except Exception:
        print("No active document found.")
        return

    print(f"Scanning document: {doc.Name}")
    
    # Define regex for "[citation](zotero://select/items/KEY)"
    # Matches: [anything](zotero://.../items/8CHARKEY)
    pattern = re.compile(r"\[.*?\]\(zotero://.*?/items/([A-Z0-9]{8})\)")
    
    # Get all text
    content = doc.Content.Text
    matches = list(pattern.finditer(content))
    
    if not matches:
        print("No plain text citations found.")
        return

    print(f"Found {len(matches)} citations. Processing backwards...")

    # Process backwards to keep indices valid
    count = 0
    for match in reversed(matches):
        start, end = match.span()
        item_key = match.group(1)
        
        # Word Range
        rng = doc.Range(Start=start, End=end)
        
        # Construct Minimal Zotero CSL Data
        # We assume libraryID 1 (My Library) for simplicity or allow Zotero to resolve.
        # Zotero usually needs specific URIs.
        # Format: http://zotero.org/users/local/{RANDOM}/items/{KEY}
        # But for local usage, usually library 0 or local is implicit?
        # Let's try the standard export format.
        
        csl_data = {
            "citationID": f"CIT_{item_key}_{start}",
            "properties": {
                "formattedCitation": "[Loading...]",
                "plainCitation": "[Loading...]",
                "noteIndex": 0
            },
            "citationItems": [
                {
                    "id": item_key,
                    "uris": [f"http://zotero.org/users/local/0/items/{item_key}"],
                    "itemData": {
                        "id": item_key,
                        "type": "article-journal",
                        "title": "Loading...",
                        "author": [{"family": "Loading", "given": "..."}]
                    }
                }
            ],
            "schema": "https://github.com/citation-style-language/schema/raw/master/schemas/input/csl-data.json"
        }
        
        json_str = json.dumps(csl_data)
        
        # Delete original text
        rng.Text = ""
        
        # Insert Field
        # Type 81 is wdFieldAddin
        field = doc.Fields.Add(Range=rng, Type=81, Text=f" ZOTERO_ITEM CSL_CITATION {json_str} ")
        
        count += 1
        print(f"Converted {item_key} at position {start}")

    print("-" * 30)
    print(f"Successfully converted {count} citations.")
    print("Now, go to the Zotero Ribbon in Word and click 'Refresh' to update them.")

if __name__ == "__main__":
    if "win32com" not in sys.modules:
        try:
            import win32com.client
        except ImportError:
            print("Error: 'pywin32' library is missing.")
            print("Please install it by running: pip install pywin32")
            input("Press Enter to exit...")
            sys.exit(1)
            
    convert_citations()
