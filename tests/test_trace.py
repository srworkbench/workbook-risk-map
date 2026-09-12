import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from audit import build
from demo_trace import run
from trace import select_cell, shortest_path, selection


class TraceTests(unittest.TestCase):
    def test_qualified_cell_normalization_and_refusal(self):
        sheets=[{'name':"O'Brien"},{'name':'Main'}]
        self.assertEqual(select_cell("'o''brien'!$b$2",sheets),"'O''Brien'!B2")
        for bad in ['A1','Main!A1:B2','Missing!A1','Main!XFE1','Main!A1048577',"'[book]Main'!A1",'Main!A0']:
            with self.subTest(bad=bad),self.assertRaises(ValueError):select_cell(bad,sheets)

    def test_breadth_first_shortest_route_and_deterministic_tie(self):
        graph={'A':['C','B'],'B':['D'],'C':['D','E'],'D':['F'],'E':['F']}
        self.assertEqual(shortest_path(graph,'A','F'),['A','B','D','F'])
        graph['A'].append('F')
        self.assertEqual(shortest_path(graph,'A','F'),['A','F'])

    def test_cycle_and_unreachable_do_not_recurse(self):
        graph={str(i):[str((i+1)%1500)] for i in range(1500)}
        path=shortest_path(graph,'0','1499');self.assertEqual(len(path),1500)
        self.assertEqual(shortest_path(graph,'0','missing'),[])
        self.assertEqual(shortest_path(graph,'0','0'),['0'])

    def test_selection_states_have_distinct_meanings(self):
        plan={'sheets':[{'name':'Main'}],'dependents':{"'Main'!A1":["'Main'!B1"]},'coverage':{'unresolved_items':1}}
        self.assertEqual(selection(plan,'Main!A1')['path_status'],'not_requested')
        same=selection(plan,'Main!A1','Main!A1')
        self.assertEqual(same['path_status'],'same_cell');self.assertEqual(same['edge_count'],0)
        absent=selection(plan,'Main!A1','Main!Z1')
        self.assertEqual(absent['path_status'],'no_parsed_path');self.assertIsNone(absent['edge_count'])
        self.assertIn('does not prove independence',absent['interpretation'])

    def test_normal_input_hidden_path_and_unresolved_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'demo';proof=run(out)
            self.assertFalse(proof['source_is_error_seed']);self.assertEqual(proof['reachable_formulas'],4)
            for folder in ['trace','unresolved']:
                directory=out/folder
                self.assertTrue((directory/'path.png').exists())
                for name,digest in json.loads((directory/'READY.json').read_text())['files'].items():
                    self.assertEqual(hashlib.sha256((directory/name).read_bytes()).hexdigest(),digest)
            plan=json.loads((out/'trace/audit.json').read_text())
            path=plan['selection']['path']
            for a,b in zip(path,path[1:]):self.assertIn(a,plan['formulas'][b]['dependencies'])
            self.assertEqual(plan['selection']['unresolved_items_in_workbook'],1)
            before=(out/'trace/audit.json').read_bytes()
            with self.assertRaises(FileExistsError):build(out/'model.xlsx',out/'trace','Inputs!B2','Checks!A1')
            self.assertEqual(before,(out/'trace/audit.json').read_bytes())

    def test_long_path_keeps_full_json_and_bounds_image(self):
        from openpyxl import Workbook
        from PIL import Image
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);wb=Workbook();wb.active.title="O'Brien";wb.active['A1']=1
            for i in range(2,32):wb.active.cell(i,1,f'=A{i-1}+1')
            wb.save(root/'long.xlsx')
            plan=build(root/'long.xlsx',root/'report',"'O''Brien'!A1","'O''Brien'!A31")
            self.assertEqual(len(plan['selection']['path']),31)
            with Image.open(root/'report/path.png') as im:self.assertEqual(im.size,(1000,2094))

    def test_cli_selection_and_missing_source_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'demo';run(out);new=Path(tmp)/'cli'
            cmd=[sys.executable,'audit.py',str(out/'model.xlsx'),'--out',str(new)]
            result=subprocess.run(cmd+['--to','Checks!A1'],capture_output=True,text=True)
            self.assertEqual(result.returncode,2);self.assertIn('--to requires --from-cell',result.stderr)
            self.assertFalse(new.exists())
            result=subprocess.run(cmd+['--from-cell','Inputs!B2','--to','Checks!A1'],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(json.loads((new/'audit.json').read_text())['selection']['edge_count'],3)


if __name__=='__main__':unittest.main()
