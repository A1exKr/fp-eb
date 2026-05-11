import win32com.client
import os

def check_master_simple(file_path):
    try:
        app = win32com.client.Dispatch("InDesign.Application")
        doc = app.Open(os.path.abspath(file_path))
        print(f"Document: {doc.Name}")
        
        print("\n--- Master Spread Frames ---")
        for master in doc.MasterSpreads:
            print(f"Master: {master.Name}")
            # Try a safer way to get text frames
            for frame in master.TextFrames:
                print(f"  Frame: {frame.Contents[:40]}...")

        print("\n--- Summary of High-Impact Page Content ---")
        # Pages are 1-indexed in ID but let's be safe
        target_indices = [1, 2, 3, 7]
        for p_idx in target_indices:
            try:
                page = doc.Pages.Item(p_idx)
                print(f"Page {page.Name}:")
                for frame in page.TextFrames:
                    content = frame.Contents if frame.Contents else "<Empty>"
                    print(f"  - {content[:80].replace('\r', ' ')}...")
            except:
                print(f"Page index {p_idx} not found.")

        doc.Close(1852776480)
    except Exception as e:
        print(f"Error: {e}")

file_path = r"c:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\examples\INDD\cp\I. Commercial.indd"
check_master_simple(file_path)
