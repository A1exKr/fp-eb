import sys
import json
import os
from pathlib import Path

# Add project root to sys.path
project_root = r'C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp'
sys.path.append(project_root)

# Import the specific service
from app.services.indd_exporter import export_proposal_to_indd

# Load the JSON data
json_path = r'C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp\data\proposals\f1173599-a4e3-45dd-af25-1f400b0bf10b.json'
with open(json_path, 'r', encoding='utf-8') as f:
    data = json.load(f)

payload = data['payload']
proposal_id = 'f1173599-a4e3-45dd-af25-1f400b0bf10b'

# Call the function
try:
    output_path = export_proposal_to_indd(proposal_id, payload)
    print(f'OUTPUT_PATH:{output_path}')
    
    # Inspect via COM
    import win32com.client
    app = win32com.client.Dispatch('InDesign.Application')
    doc = app.Open(str(output_path))
    
    page_count = doc.Pages.Count
    print(f'PAGE_COUNT:{page_count}')
    
    # Text sample from first page
    text_samples = []
    first_page = doc.Pages.Item(1)
    # Broadly look for text frames
    for i in range(1, first_page.TextFrames.Count + 1):
        tf = first_page.TextFrames.Item(i)
        content = tf.Contents
        if content and len(content.strip()) > 0:
            text_samples.append(content.strip()[:50])
            if len(text_samples) >= 3:
                break
    
    print(f'TEXT_SAMPLES:{json.dumps(text_samples)}')
    # Note: doc.Close(6) would be SaveOptions.no but common practice is 1833319200 (idSaveOptions.idNo) or just closing
    doc.Close(1833319200) 
except Exception as e:
    print(f'ERROR:{str(e)}')
    import traceback
    traceback.print_exc()
