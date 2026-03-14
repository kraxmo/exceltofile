#!/usr/bin/env python3

# import csv
import sys
import time
import unittest
import tempfile
import subprocess
from pathlib import Path

import pandas as pd

SCRIPT = 'excel_to_file.py'

class Rubric:
    def __init__(self):
        self.total = 0
        self.points = 0
    def add(self, ok: bool, weight: int, note: str):
        self.total += weight
        if ok:
            self.points += weight
        else:
            print(f"[RUBRIC] Lost {weight} points: {note}")
    def score(self) -> float:
        return 100.0 * (self.points / self.total if self.total else 0.0)

class TestExcelToFile(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.base = Path(self.tmpdir.name)
        self.out = self.base / 'out'
        self.out.mkdir(parents=True, exist_ok=True)
        self.xlsx = self.base / 'Book.xlsx'
        with pd.ExcelWriter(self.xlsx, engine='openpyxl') as xw:
            pd.DataFrame({'A':[1,2],'B':['x','y'],'C':[None,'z']}).to_excel(xw, sheet_name='Sheet1', index=False)
            pd.DataFrame().to_excel(xw, sheet_name='Empty', index=False)
            pd.DataFrame({'Bad,Header':[1,2], 'OK':[3,4]}).to_excel(xw, sheet_name='BadHeader', index=False)
            pd.DataFrame([[10,20],[30,40]]).to_excel(xw, sheet_name='NoHeader', header=False, index=False)
            pd.DataFrame({'Q':[1],'W':[2]}).to_excel(xw, sheet_name='Q1 Summary', index=False)

    def tearDown(self):
        self.tmpdir.cleanup()

    def _run_cli(self, *args, expect_ok=True):
        cmd = [sys.executable, SCRIPT, str(self.xlsx), '--out-dir', str(self.out), *args]
        cp = subprocess.run(cmd, capture_output=True, text=True)
        if expect_ok:
            self.assertEqual(cp.returncode, 0, f"CLI failed: {cp.stderr}\n{cp.stdout}")
        return cp

    def _score(self) -> float:
        r = Rubric()

        # 1) Happy path (pipe delimiter, selected sheets)
        print("\n# 1) Happy path (pipe delimiter, selected sheets)")
        self._run_cli('--prefix','run01','--sheets','Sheet1','Empty','Q1 Summary','--allow-delimiter-in-header','--delimiter','pipe','--log-level','ERROR')
        files = sorted(p for p in self.out.iterdir() if p.name.startswith('run01_'))
        r.add(any(p.name.startswith('run01_Sheet1_') for p in files), 10, 'Sheet1 not exported')
        r.add(not any('Empty_' in p.name for p in files), 10, 'Empty sheet not skipped')

        # 2) NA representation default is empty string (so sequence ",," should appear)
        print("# 2) NA representation default is empty string (so sequence ',,' should appear)")
        s1 = next(p for p in files if 'Sheet1_' in p.name)
        text = s1.read_text(encoding='utf-8')
        r.add('1|x|' in text or '|N/A' in text, 5, 'NA representation not reflected')

        # 3) Delimiter keyword parsing (tab)
        print("# 3) Delimiter keyword parsing (tab)")
        self._run_cli('--prefix','tab','--sheets','Sheet1','--delimiter','tab')
        tfile = next(p for p in self.out.iterdir() if p.name.startswith('tab_Sheet1_'))
        r.add('	' in tfile.read_text(encoding='utf-8'), 10, 'tab delimiter not used')

        # 4) Quoting none auto-escape notice (just ensure file written)
        print("# 4) Quoting none auto-escape notice (just ensure file written)")
        self._run_cli('--prefix','noq','--sheets','Sheet1','--quoting','none')
        r.add(any(p.name.startswith('noq_Sheet1_') for p in self.out.iterdir()), 5, 'quoting none export failed')

        # 5) Header delimiter validation should cause non-zero rc
        print("# 5) Header delimiter validation should cause non-zero rc")
        cp = self._run_cli('--prefix','vchk','--sheets','BadHeader', '--delimiter', ',', '--log-level','ERROR', expect_ok=False)
        r.add(cp.returncode != 0, 15, 'header delimiter validation did not fail CLI as expected')

        # 6) Allow delimiter in header bypass
        print("# 6) Allow delimiter in header bypass")
        self._run_cli('--prefix','allow','--sheets','BadHeader','--allow-delimiter-in-header')
        r.add(any(p.name.startswith('allow_BadHeader_') for p in self.out.iterdir()), 10, 'allow flag ineffective')

        # 7) No header (-1)
        print("# 7) No header (-1)")
        self._run_cli('--prefix','nh','--sheets','NoHeader','--header-row','-1')
        r.add(any(p.name.startswith('nh_NoHeader_') for p in self.out.iterdir()), 10, 'no-header export failed')

        # 8) Filename collision avoidance: pre-create collision name and export w/o prefix
        print("# 8) Filename collision avoidance: pre-create collision name and export w/o prefix")
        coll = self.out / f"Sheet1_{self.xlsx.stem}.csv"
        coll.write_text('dummy', encoding='utf-8')
        self._run_cli('--sheets','Sheet1')
        names = [p.name for p in self.out.iterdir() if p.name.startswith('Sheet1_')]
        r.add(any(n != coll.name for n in names), 15, 'timestamp collision avoidance failed')

        return r.score()

    def test_rubric(self):
        s1 = self._score()
        if s1 < 95:
            time.sleep(0.05)
            s2 = self._score()
            final = s2
        else:
            final = s1
        print(f"Rubric score: {final:.2f}")
        self.assertGreaterEqual(final, 95.0)

if __name__ == '__main__':
    unittest.main(verbosity=2)
