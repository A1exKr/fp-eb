import win32com.client
import os

def inspect_advanced(file_path):
    try:
        app = win32com.client.Dispatch("InDesign.Application")
        doc = app.Open(os.path.abspath(file_path))
        print(f"Document: {doc.Name}")
        
        # Check all frames in the document, not just on pages (could be on layers/spreads)
        print("\n--- Identifying Potential Data Mapping Targets (Labeled or Named Frames) ---")
        found_target = False
        for frame in doc.AllPageItems:
            if frame.Type == 1131707510: # idPageItemTypes.idTextFrameType
                # Check for script labels or specific names
                label = frame.Label
                name = frame.Name
                if label or name:
                    print(f"Frame ID {frame.Id}: Label='{label}', Name='{name}', Sample='{frame.Contents[:30]}...'")
                    found_target = True
        
        if not found_target:
            print("No frames with labels or names found.")

        print("\n--- Placeholders ---")
        # Search for text bracketed by curly braces or all-caps placeholders
        for page in doc.Pages:
            for frame in page.TextFrames:
                content = frame.Contents
                if content and ("[" in content or "{" in content or "<" in content):
                    print(f"Possible Placeholder on Page {page.Name}: {content[:100]}...")

        doc.Close(1852776480)
    except Exception as e:
        print(f"Error: {e}")

file_path = r"c:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\examples\INDD\cp\I. Commercial.indd"
inspect_advanced(file_path)
