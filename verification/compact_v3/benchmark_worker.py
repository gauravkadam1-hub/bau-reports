"""Same synthetic records through original, v2, and v3 workbook writers."""
from __future__ import annotations
import argparse, csv, json, sys, time
from datetime import date, datetime, timedelta
from pathlib import Path
from zipfile import ZipFile

parser=argparse.ArgumentParser()
parser.add_argument('--mode',choices=['original','v2','v3','v3-uncompressed'],required=True)
parser.add_argument('--shape',choices=['sparse','dense'],required=True)
parser.add_argument('--rows',type=int,default=20000)
parser.add_argument('--workspace',type=Path,required=True)
a=parser.parse_args()
root=a.workspace.resolve()
project=root/({'original':'original','v2':'v2','v3':'updated','v3-uncompressed':'updated'}[a.mode])/'bau-reports-main'
sys.path.insert(0,str(project))
if a.mode=="original": sys.path.insert(0,str(project/"engines/onboarding"))
profile=json.loads((root/'v2/bau-reports-main/engines/onboarding/onboarding_layout.json').read_text())
heads=profile['headers']; start=date(2026,10,1);end=start+timedelta(days=19)
dc=[f'Synthetic {i:02} DC' for i in range(8)];xd=[f'Synthetic {i:02} XD' for i in range(4)]
records=[]
for i in range(a.rows):
    day=start+timedelta(days=(i//12)%20);site=(dc+xd)[i%12]
    values={'ID':f'{i:08}', 'Name':f'Synthetic Employee {i:07}', 'Status':'Onboarded - Active',
            'Role':'GC','Registered Facility':site, 'Onboarded Date':day.isoformat(),
            'Approved By':f'Synthetic HR {i%7:02}', 'Onboarded by Aadhaar (Y/N)':'N' if i%4==0 else 'Y',
            'Hyperverge Status':'red' if i%5==0 else 'green', 'Phone Number':f'0{9000000000+i}',
            'Present Count Since DOJ':str(i%21)}
    if a.shape=='dense':
        values.update({'Sub Status':'Active','Role Type':'Worker','Department':'Operations',
            'Registered Date':(day-timedelta(days=2)).isoformat(),
            'Rejection Date':'','Rejection Reason':'','Vendor':f'Synthetic Staffing {i%50}',
            'Lead Source':['Referral','Walk-in','Agency'][i%3],
            'Date of Birth':f'199{i%10}-0{(i%9)+1}-15', 'Auto Verification Status':'Completed',
            'Auto Verification Comments':f'Synthetic verification reference TEST-{i:08}; no real person.',
            'Winback Count':str(i%3), 'Aadhaar Number':f'TEST-{i:012}',
            'Father Name':f'Synthetic Parent {i:07}',
            'Permanent Address':f'{i+1} Test Street, Synthetic Area {i%300}, Fictional City {i%40}, TEST-{i%100000:06}',
            'Gender':['Female','Male'][i%2]})
    records.append([values.get(h) or None for h in heads])
out=root/'verification/size';out.mkdir(exist_ok=True)
folder=out/f'{a.shape}_{a.rows}_{a.mode}';folder.mkdir(exist_ok=True)
if a.mode=='original':
    source=out/f'{a.shape}_{a.rows}.csv'
    with source.open('w',newline='',encoding='utf-8') as f:
        w=csv.writer(f);w.writerow(heads);w.writerows(records)
    from engines.onboarding.report_generator import generate_report
    t=time.perf_counter()
    output,*_=generate_report(source,selected_start_date=start,selected_end_date=end,
                              selected_dc=dc,selected_xd=xd,output_dir=folder,generation_date=date(2026,10,24))
    elapsed=time.perf_counter()-t
else:
    from engines.onboarding.onboarding_common import PreparedReport,Counts,load_profile
    from engines.onboarding import onboarding_writer as writer
    counts={};manual={};red={}
    for r in records:
        day=date.fromisoformat(r[9]);r[9]=datetime.combine(day,datetime.min.time())
        key=(r[7],day);v=Counts(1,int(r[15]=='Y'),int(r[15]=='N'),int(r[20]=='red'))
        counts[key]=counts.get(key,Counts())+v
        if v.manual: manual[r[14]]=manual.get(r[14],0)+1
        if v.red: red[r[14]]=red.get(r[14],0)+1
    sort=lambda s:(-counts.get((s,end),Counts()).red,-counts.get((s,end),Counts()).manual)
    report=PreparedReport([start+timedelta(days=i) for i in range(20)],sorted(dc,key=sort),sorted(xd,key=sort),
                          records,len(records),counts,sorted(manual),sorted(red),manual,red)
    if a.mode=='v3-uncompressed':writer.compact_workbook=lambda _:None
    output=folder/'report.xlsx';t=time.perf_counter();writer.write_report(output,load_profile(),report)
    elapsed=time.perf_counter()-t
with ZipFile(output) as z:
    parts=[{'name':i.filename,'compressed':i.compress_size,'raw':i.file_size} for i in z.infolist()]
result={'mode':a.mode,'shape':a.shape,'rows':a.rows,'path':str(output), 'bytes':output.stat().st_size,
        'elapsed_seconds':round(elapsed,3),'parts':parts,
        'timing_scope':'original includes CSV read; other modes measure writer/compression/validation only, NOT comparable end-to-end pipelines'}
(out/f'{a.shape}_{a.rows}_{a.mode}.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps({k:v for k,v in result.items() if k not in ['parts','timing_scope']}),flush=True)
