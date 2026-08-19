#!/usr/bin/env python3
from pathlib import Path
import csv, json, math, sys

BASE=Path(__file__).resolve().parent
DATA=BASE/'data'; OUT=BASE/'outputs'

def rows(path):
    with path.open(encoding='utf-8-sig',newline='') as f: return list(csv.DictReader(f))

def close(a,b,tol=5e-3):
    return abs(float(a)-float(b)) <= tol

checks=[]
def check(name, cond, detail):
    checks.append((name,bool(cond),detail))

clean=rows(DATA/'clean_analysis_data_public.csv')
control=[float(r['Acceptance control']) for r in clean]
failure=[float(r['Acceptance failure']) for r in clean]
change=[float(r['Acceptance change']) for r in clean]
check('Acceptance control mean', close(sum(control)/len(control),5.31125,1e-9), sum(control)/len(control))
check('Acceptance failure mean', close(sum(failure)/len(failure),4.969375,1e-9), sum(failure)/len(failure))
check('Acceptance mean change', close(sum(change)/len(change),-0.341875,1e-9), sum(change)/len(change))

rec=rows(DATA/'failure_recognition_coding_public.csv')
hw=[r for r in rec if r['Failure type']=='Hardware']; inter=[r for r in rec if r['Failure type']=='Interaction']
early=[r for r in rec if r['Timing']=='Early']; late=[r for r in rec if r['Timing']=='Late']
count=lambda rs: sum(r['Recognized']=='Yes' for r in rs)
check('Hardware recognition', count(hw)==14, count(hw))
check('Interaction recognition', count(inter)==5, count(inter))
check('Early recognition', count(early)==7, count(early))
check('Late recognition', count(late)==12, count(late))
check('Total recognition', count(rec)==19, count(rec))

anova=rows(OUT/'table_E7_two_way_anova.csv')
by={r['Source']:r for r in anova}
check('ANOVA failure type F', close(by['Failure type']['F'],0.3258945091,1e-9), by['Failure type']['F'])
check('ANOVA failure timing F', close(by['Failure timing']['F'],0.3138253964,1e-9), by['Failure timing']['F'])
check('ANOVA interaction F', close(by['Failure type × timing']['F'],1.9094531895,1e-9), by['Failure type × timing']['F'])

failed=[x for x in checks if not x[1]]
for name,ok,detail in checks:
    print(('PASS' if ok else 'FAIL')+f' | {name}: {detail}')
if failed:
    print(f'\n{len(failed)} verification check(s) failed.')
    sys.exit(1)
print(f'\nAll {len(checks)} verification checks passed.')
