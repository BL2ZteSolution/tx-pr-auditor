"""Real shared MW subprocesses plus the Auditor pipeline, using synthetic workbooks."""
import hashlib
import json
import os
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from tests.shared_fixture import approval_context
from tests.test_skill_contract_integration import first_headers, audit, contract


def write_epms(path):
    mw = Path(approval_context()['root'])
    profile=json.loads((mw/'config/du_profiles/tx_mini_pr_v1.yaml').read_text(encoding='utf-8'))
    fields={}
    for name,cfg in profile['field_mapping'].items():
        candidates=[c for c in cfg.get('source_candidates',[]) if c.get('mapping_status')=='APPROVED']
        if candidates: fields[name]=candidates[0]['fingerprint']
    values={'site_code':'AUDIT-SYNTH-001','site_name':'Synthetic test site','du_key':'DU-AUDIT-001','region':'Central','state':'Selangor','tx_sow_raw':'MW Swap','subcontractor_tss':'GCI','subcontractor_ti':'GCI','subcontractor_planning':'GCI','antenna_size_ne':'0.6m','antenna_size_fe':'0.6m'}
    book=Workbook();sheet=book.active;sheet.title='data'
    for col,(name,fp) in enumerate(fields.items(),1):
        for row,key in enumerate(['field_code','wbs_stage','task_name','display_header'],1):sheet.cell(row,col,fp[key])
        sheet.cell(5,col,values.get(name,''))
    book.save(path);book.close()


def write_final(path):
    headers=first_headers(audit.FINAL_PO_FIELD_MAP)
    values={'dispatch_date':datetime(2026,8,1),'dispatch_order_number':'DO-SYNTH-001','po_line_number':1,'request_number':'REQ-SYNTH-001','project_name':'Malaysia_CelcomDigi_Project','business_domain':'TSS survey','submitted_subcontractor':'GCI','logical_site_name':'Wireless RAN/TX Mini Project/AUDIT-SYNTH-001','du':'DU-AUDIT-001','site_code':'AUDIT-SYNTH-001','submitted_item_code':'350000589343','submitted_item_description':'MW only site Technical site survey(Hop)','submitted_unit':'Hop','submitted_quantity':1,'settlement_quantity':1}
    book=Workbook();sheet=book.active;sheet.title=audit.FINAL_PO_SHEET_NAME
    sheet.append(list(headers.values()));sheet.append([values.get(field,'') for field in headers]);book.save(path);book.close()


class RealSharedIntegrationTests(unittest.TestCase):
    def test_actual_tss_ti_subprocesses_and_audit(self):
        for extension in ['.xlsx','.xlsm']:
            with self.subTest(extension=extension), tempfile.TemporaryDirectory() as tmp:
                root=Path(tmp); epms=root/('epms'+extension);final=root/'final.xlsx';write_epms(epms);write_final(final)
                before=hashlib.sha256(epms.read_bytes()).hexdigest();context=approval_context()
                envelope={'schemaVersion':'1.0','jobId':'REAL-SHARED-AUDIT','skill':{'skillId':'tx-pr-auditor','version':'1.2.0'},'dependencies':[{k:context[k] for k in ['skillId','version','packageSha256']}], 'parameters':{'annotateEcc':True},'files':[{'name':name,'path':path.name,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()} for name,path in [('final_po',final),('epms',epms)]]}
                manifest=root/'input.json';manifest.write_text(json.dumps(envelope),encoding='utf-8')
                with patch.dict(os.environ,{'TX_PR_AUDITOR_MW_DEPENDENCY':json.dumps(context)}):
                    self.assertEqual(contract.run(manifest),0,(root/'result.json').read_text(encoding='utf-8'))
                result=json.loads((root/'result.json').read_text(encoding='utf-8'))
                self.assertEqual(result['summary']['metrics']['classifications'],{'Normal':1})
                dep=result['summary']['metrics']['entitlement'];self.assertEqual(dep['dependencyVersion'],'4.1.0');self.assertEqual(dep['dependencyPackageSha256'],context['packageSha256'])
                self.assertEqual([r['scope'] for r in dep['scopes']],['TSS','TI'])
                self.assertTrue(all(r['workbookCount']>0 for r in dep['scopes']))
                self.assertEqual(hashlib.sha256(epms.read_bytes()).hexdigest(),before)
                for item in result['outputs']:
                    self.assertEqual(hashlib.sha256((root/item['path']).read_bytes()).hexdigest(),item['sha256'])
                for scope in ['tss','ti']:
                    child=json.loads((root/f'temp/create-pr-cd/{scope}/skill-input.json').read_text(encoding='utf-8'))
                    self.assertEqual(child['skill']['version'],'4.1.0');self.assertTrue(child['files'][0]['path'].endswith('.xlsx'))

    def test_child_failure_and_existing_cancellation_are_safe(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);source=root/'invalid.xlsx';book=Workbook();book.save(source);book.close()
            context=approval_context()
            with patch.dict(os.environ,{'TX_PR_AUDITOR_MW_DEPENDENCY':json.dumps(context)}):
                with self.assertRaises(contract.ContractError):
                    contract.run_entitlement_scope({'jobId':'BAD'},root,source,root/'cancel','TSS',0,100)
