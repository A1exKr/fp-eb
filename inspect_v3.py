import win32com.client
import os

def inspect_v3(file_path):
    try:
        app = win32com.client.Dispatch("InDesign.Application")
        doc = app.Open(os.path.abspath(file_path))
        print(f"Document: {doc.Name}")
        
        print("\n--- Scanning All Text Frames for Names/Labels ---")
        for i, page in enumerate(doc.Pages):
            for j, frame in enumerate(page.TextFrames):
                label = frame.Label
                name = frame.Name
                content = frame.Contents if frame.Contents else ""
                
                # Check for bracketed placeholders like [CLIENT NAME]
                import re
                placeholders = re.findall(r"\[.*?\]", content)
                
                if label or name or placeholders:
                    print(f"Page {page.Name} | Frame {j}: Name='{name}', Label='{label}', Content='{content[:40]}...'")
                    if placeholders:
                        print(f"  -> Found Placeholders: {placeholders}")
        
        doc.Close(1852776480)
    except Exception as e:
        print(f"Error: {e}")

file_path = r"c:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\examples\INDD\cp\I. Commercial.indd"
inspect_v3(file_path)
