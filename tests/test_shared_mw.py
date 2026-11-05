import hashlib
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from tests.shared_fixture import approval_context
from tests.test_skill_contract import contract
from shared_mw import verify_shared_mw, SharedDependencyError

class SharedMWTests(unittest.TestCase):
    def test_real_shared_approval_and_no_nested_checkout(self):
        context = approval_context()
        with patch.dict(os.environ, {"TX_PR_AUDITOR_MW_DEPENDENCY": json.dumps(context)}):
            result = verify_shared_mw()
            self.assertEqual(result["version"], context["version"])
            self.assertEqual(result["packageSha256"], context["packageSha256"])
            self.assertEqual(result["root"], Path(context["root"]))
        self.assertFalse((Path(__file__).parents[1] / "dependencies/create-pr-cd").exists())
        self.assertEqual(contract.ENTITLEMENT_SCOPES, ("TSS", "TI"))

    def test_missing_approval_fails_without_nested_fallback(self):
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaises(SharedDependencyError):
                verify_shared_mw()

    def test_invalid_approvals_fail_closed(self):
        for change in ({"version": "4.0.0"}, {"root": "relative"}, {"root": "C:/nonexistent-mw"},
                       {"packageSha256": "0"*64}, {"runtimeFiles": []},
                       {"runtimeFiles": ["../outside.py", "skill.json", "src/main.py", "Info/input/pr_model.xlsx"]}):
            with self.subTest(change=change), patch.dict(os.environ, {"TX_PR_AUDITOR_MW_DEPENDENCY": json.dumps({**approval_context(), **change})}):
                with self.assertRaises(SharedDependencyError):
                    verify_shared_mw()

    def test_files_are_rehashed_and_commit_mismatch_rejected(self):
        context = approval_context()
        with patch.dict(os.environ, {"TX_PR_AUDITOR_MW_DEPENDENCY": json.dumps({**context,"commit":"0"*40})}):
            with self.assertRaises(SharedDependencyError) as error:
                verify_shared_mw()
            self.assertEqual(error.exception.code, "CREATE_PR_CD_DEPENDENCY_COMMIT_MISMATCH")
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); (root/'src').mkdir(); (root/'Info/input').mkdir(parents=True)
            (root/'skill.json').write_text(json.dumps({'skillId':'create-pr-cd','version':context['version'],'runtime':{'entrypoint':'src/main.py'}}))
            (root/'src/main.py').write_text('print(1)\n'); (root/'Info/input/pr_model.xlsx').write_bytes(b'model')
            files=['skill.json','src/main.py','Info/input/pr_model.xlsx']; digest=hashlib.sha256()
            for name in sorted(files): digest.update((name+'\0'+hashlib.sha256((root/name).read_bytes().replace(b'\r\n',b'\n')).hexdigest()+'\n').encode())
            approval={**context,'root':str(root),'runtimeFiles':files,'packageSha256':digest.hexdigest()}
            with patch.dict(os.environ, {'TX_PR_AUDITOR_MW_DEPENDENCY':json.dumps(approval)}):
                verify_shared_mw()
                (root/'src/main.py').write_text('print(2)\n')
                with self.assertRaises(SharedDependencyError) as error: verify_shared_mw()
                self.assertEqual(error.exception.code,'CREATE_PR_CD_DEPENDENCY_FINGERPRINT_MISMATCH')

    def test_old_frozen_dependency_blocks_before_entitlement(self):
        from tests.test_skill_contract import SkillContractTests
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); manifest=SkillContractTests().make_request(root)
            document=json.loads(manifest.read_text());document['dependencies']=[{'skillId':'create-pr-cd','version':'4.0.0','packageSha256':'0'*64}]
            manifest.write_text(json.dumps(document))
            with patch.dict(os.environ, {'TX_PR_AUDITOR_MW_DEPENDENCY':json.dumps(approval_context())}), patch.object(contract,'run_entitlement_generation') as generate:
                self.assertNotEqual(contract.run(manifest),0);generate.assert_not_called()
            result=json.loads((root/'result.json').read_text())
            self.assertEqual(result['error']['code'],'CREATE_PR_CD_DEPENDENCY_CHANGED')

    def test_prestart_child_cancellation_and_planning_rejection(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);cancel=root/'cancel';cancel.write_text('cancel')
            with self.assertRaises(contract.CancelledError), patch.object(contract.subprocess,'Popen') as launch:
                contract.run_entitlement_scope({'jobId':'CANCEL'},root,root/'input.xlsx',cancel,'TSS',0,100)
            launch.assert_not_called()
            cancel.unlink()
            with self.assertRaises(contract.ContractError) as error:
                contract.run_entitlement_scope({'jobId':'PLANNING'},root,root/'input.xlsx',cancel,'PLANNING',0,100)
            self.assertEqual(error.exception.code,'CREATE_PR_CD_SCOPE_INVALID')
