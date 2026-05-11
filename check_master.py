import win32com.client
import os

def check_master_info(file_path):
    try:
        app = win32com.client.Dispatch("InDesign.Application")
        doc = app.Open(os.path.abspath(file_path))
        print(f"Document: {doc.Name}")
        
        print("\n--- Master Page Analysis ---")
        for master in doc.MasterSpreads:
            print(f"Master Page: {master.Name}")
            for item in master.AllPageItems:
                if item.Type == 1131707510: # TextFrame
                    content = item.Contents if item.Contents else ""
                    if content.strip():
                        print(f"  Frame on Master: {content[:50]}...")

        print("\n--- Summary of High-Impact Pages ---")
        pages_to_check = [1, 2, 3, 7] # Cover, Revision, Intro, Total Payment
        for pnum in pages_to_check:
            page = doc.Pages.Item(pnum)
            print(f"Page {pnum}:")
            for frame in page.TextFrames:
                print(f"  - {frame.Contents[:100]}...")

        doc.Close(1852776480)
    except Exception as e:
        print(f"Error: {e}")

file_path = r"c:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\examples\INDD\cp\I. Commercial.indd"
check_master_info(file_path)
