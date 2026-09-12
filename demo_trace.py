"""Trace a normal input through a hidden sheet to a chosen result."""
import argparse
import hashlib
import json
from pathlib import Path
from openpyxl import Workbook
from audit import build


def run(output):
    output.mkdir(mode=0o700)
    wb=Workbook();wb.remove(wb.active)
    wb.properties.creator=wb.properties.lastModifiedBy='SR Workbench'
    for name in ['Inputs','Model','Summary','Checks','Notes']:wb.create_sheet(name)
    wb['Model'].sheet_state='hidden';wb['Checks'].sheet_state='veryHidden'
    wb['Inputs']['B2']=125
    wb['Model']['B2']="='Inputs'!B2*2";wb['Model']['C2']="='Inputs'!B2*3"
    wb['Summary']['B2']="='Model'!B2+'Model'!C2"
    wb['Checks']['A1']="='Summary'!B2>0"
    wb['Notes']['A1']='=INDIRECT("Inputs!B2")'
    source=output/'model.xlsx';wb.save(source)
    digest=hashlib.sha256(source.read_bytes()).hexdigest()
    plan=build(source,output/'trace','Inputs!B2','Checks!A1')
    selected=plan['selection']
    assert selected['path']==["'Inputs'!B2","'Model'!B2","'Summary'!B2","'Checks'!A1"]
    assert len(selected['reachable_formulas'])==4
    other=build(source,output/'unresolved','Inputs!B2','Notes!A1')
    assert other['selection']['path_status']=='no_parsed_path'
    assert other['coverage']['unresolved_items']==1
    assert hashlib.sha256(source.read_bytes()).hexdigest()==digest
    proof={'verified_demo':True,'source_unchanged':True,'reachable_formulas':4,
           'path':selected['path'],'edge_count':3,'unresolved_target_status':'no_parsed_path',
           'unresolved_items':1,'source_is_error_seed':False}
    assert not any(i['cell']=="'Inputs'!B2" for i in plan['issues'])
    (output/'demo-proof.json').write_text(json.dumps(proof,indent=2)+'\n')
    return proof


if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--out',type=Path,required=True)
    print(json.dumps(run(p.parse_args().out)))
