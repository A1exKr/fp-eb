import win32com.client
import os
import sys

def inspect_indd(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    try:
        # Use DispatchEx to ensure a fresh session if needed, but Dispatch is usually fine
        app = win32com.client.Dispatch("InDesign.Application")
    except Exception as e:
        print(f"Error: Could not dispatch InDesign.Application. {e}")
        return

    try:
        # InDesign COM: Open(fileName, [showingWindow])
        # Force a normal path
        file_path = os.path.abspath(file_path)
        doc = app.Open(file_path)
        
        print(f"Document: {doc.Name}")
        print(f"Page Count: {doc.Pages.Count}")
        
        for i, page in enumerate(doc.Pages):
            print(f"\nPage {i+1} (Name: {page.Name}):")
            text_frames = page.TextFrames
            print(f"  Text Frames: {text_frames.Count}")
            for j, frame in enumerate(text_frames):
                # Using getattr to avoid potential property errors in different ID versions
                label = getattr(frame, 'Label', '<No Label>')
                name = getattr(frame, 'Name', '<No Name>')
                try:
                    content = frame.Contents
                    content_sample = (content[:50].replace('\r', ' ').replace('\n', ' ') if content else '<Empty>')
                except:
                    content_sample = '<Error getting contents>'
                
                print(f"    Frame {j+1}: Label='{label}', Name='{name}', Content='{content_sample}...'")
        
        doc.Close(1852776480) # idSaveOptions.idNo is '1852776480' (which is 'no  ') but usually 1852776480; or 2 for JS enum. 
        # In COM, it's safer to use the constant or just 1852776480 (idNo)
    except Exception as e:
        print(f"Error during inspection: {e}")

file_path = r"c:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\examples\INDD\cp\I. Commercial.indd"
inspect_indd(file_path)
