import json, sys, os
from pathlib import Path

root = r'C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp'
sys.path.append(root)

try:
    from app.services.indd_exporter import export_proposal_to_indd
    
    json_path = os.path.join(root, 'data', 'proposals', 'f1173599-a4e3-45dd-af25-1f400b0bf10b.json')
    with open(json_path, 'r', encoding='utf-8') as f:
        wrapper = json.load(f)
    
    payload = wrapper['payload']
    proposal_id = 'f1173599-a4e3-45dd-af25-1f400b0bf10b'
    
    output_path = export_proposal_to_indd(proposal_id, payload)
    print(f'SUCCESS: {output_path}')
    
    import win32com.client
    app = win32com.client.Dispatch('InDesign.Application')
    doc = app.Open(str(output_path))
    print(f'PAGES: {len(doc.Pages)}')
    doc.Close(1627389540)
except Exception as e:
    print(f'ERROR: {str(e)}')
    sys.exit(1)
