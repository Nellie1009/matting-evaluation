"""Generate formula cells for a new matting analysis sheet; no XLSX writes."""
import argparse,json,re
from pathlib import Path

def col(n):
 s=''
 while n:n,k=divmod(n-1,26);s=chr(65+k)+s
 return s

def extreme_colors(values,reverse=False):
 numeric=[x for x in values if isinstance(x,(int,float)) and not isinstance(x,bool)]
 if not numeric or max(numeric)==min(numeric):return ['white']*len(values)
 lo,hi=min(numeric),max(numeric)
 return [('red' if reverse else 'green') if x==hi else ('green' if reverse else 'red') if x==lo else 'white' for x in values]

def build(config):
 required=('sheet','first_row','last_row','columns','models','tasks','labels')
 if any(k not in config for k in required):raise ValueError('Missing required mapping field')
 c=config['columns'];models=config['models'];tasks=config['tasks'];labels=config['labels'];n=len(tasks);m=len(models)
 if not n or not m or len(set(models))!=m or len(set(tasks))!=n or len(labels)!=n:raise ValueError('Expected unique, nonempty tasks/models and one label pair per task')
 if any(k not in c for k in ('task','model','status','total','scores')) or len(c.get('scores',[]))!=6:raise ValueError('Missing source columns or six score columns')
 if any(not re.fullmatch('[A-Z]+',x) for x in [c['task'],c['model'],c['status'],c['total']]+c['scores']+[c[k] for k in ('complexity','fusion') if k in c]):raise ValueError('Column mappings must use uppercase Excel letters')
 first,last=config['first_row'],config['last_row']
 if not isinstance(first,int) or not isinstance(last,int) or first<1 or last<first:raise ValueError('Invalid source row bounds')
 if any(len(x)!=2 or any(v not in ('低','中','高') for v in x) for x in labels):raise ValueError('Invalid difficulty labels')
 sheet="'"+config['sheet'].replace("'","''")+"'!"
 def source(k):return f'{sheet}${k}${first}:${k}${last}'
 tr,mr,sr,total=[source(c[k]) for k in ('task','model','status','total')]
 cells={};groups=[];details=[]
 def value(a,v):cells[a]={'value':v}
 def formula(a,f):cells[a]={'formula':'='+f}
 def average(a,r):formula(a,f'IF(COUNT({r})=0,"",AVERAGE({r}))')
 def highlight(r,group,**kwargs):groups.append({'range':r,'group':group,**kwargs})
 def task_mean(a,score,task,model):
  criteria=f'{tr},{task},{mr},{model},{sr},"已评"'
  formula(a,f'IF(COUNTIFS({criteria})=0,"",SUMIFS({score},{criteria})/COUNTIFS({criteria}))')
 value('B2','任务');value('C2','任务平均分');end=col(m+3);avg=n+3;cover=n+4;rank=n+5
 for j,model in enumerate(models):value(f'{col(j+4)}2',model)
 for i,task in enumerate(tasks):
  r=i+3;value(f'B{r}',task);average(f'C{r}',f'D{r}:{end}{r}')
  for j in range(m):
   k=col(j+4);task_mean(f'{k}{r}',total,f'$B{r}',f'{k}$2')
  highlight(f'D{r}:{end}{r}',f'task:{task}')
 for r,text in [(avg,'平均得分'),(cover,f'已评任务数（共{n}）'),(rank,'综合排名')]:value(f'B{r}',text)
 average(f'C{avg}',f'D{avg}:{end}{avg}')
 for j in range(m):
  k=col(j+4);average(f'{k}{avg}',f'{k}3:{k}{n+2}');formula(f'{k}{cover}',f'COUNT({k}3:{k}{n+2})')
  formula(f'{k}{rank}',f'IF({k}{cover}<>{n},"",1+COUNTIFS($D${cover}:${end}${cover},{n},$D${avg}:${end}${avg},">"&{k}{avg}))')
 highlight(f'D{avg}:{end}{avg}','model-means',eligible_range=f'D{cover}:{end}{cover}',required_count=n)
 highlight(f'D{rank}:{end}{rank}','ranking',reverse=True)
 metrics=['透明背景（0/4）','主体识别（0-4）','抠图精细度（0-4）','色彩还原（0-4）','避免AI修改（0-4）','避免形变/裁切（0/4）']
 for r,text in [(rank+1,'已评运行数'),(rank+2,'缺失任务数'),(rank+3,'待评运行数')]:value(f'B{r}',text)
 for j in range(m):
  k=col(j+4);formula(f'{k}{rank+1}',f'COUNTIFS({mr},{k}$2,{sr},"已评")');formula(f'{k}{rank+2}',f'{n}-{k}{cover}');formula(f'{k}{rank+3}',f'COUNTIFS({mr},{k}$2,{sr},"<>已评")')
 row=rank+6
 for j,model in enumerate(models):
  value(f'B{row}',model);header=row+1;start=row+2;finish=start+n-1;meanrow=finish+1;value(f'B{header}','任务')
  for d,metric in enumerate(metrics):value(f'{col(d+3)}{header}',metric)
  for i,task in enumerate(tasks):
   r=start+i;value(f'B{r}',task)
   for d in range(6):task_mean(f'{col(d+3)}{r}',source(c['scores'][d]),f'$B{r}',f'$B${row}')
  value(f'B{meanrow}','平均得分')
  for d in range(6):
   k=col(d+3);average(f'{k}{meanrow}',f'{k}{start}:{k}{finish}');highlight(f'{k}{start}:{k}{finish}',f'detail:{model}:{d}')
  details.append({'model':model,'header':header,'mean_row':meanrow});row=meanrow+3
 value(f'B{row}','评测维度分析');header=row+1
 for d,metric in enumerate(metrics):value(f'{col(d+3)}{header}',metric)
 start=row+2;value(f'I{header}','已评任务数')
 for j,detail in enumerate(details):
  r=start+j;value(f'B{r}',detail['model'])
  formula(f'I{r}',f'{col(j+4)}{cover}')
  for d in range(6):formula(f'{col(d+3)}{r}',f'{col(d+3)}{detail["mean_row"]}')
 for d in range(6):
  k=col(d+3);comparison=f'{k}{start}:{k}{start+m-1}';eligibility=f'I{start}:I{start+m-1}';average(f'{k}{start+m}',comparison);highlight(comparison,f'dimension:{d}',eligible_range=eligibility,required_count=n)
  for j,detail in enumerate(details):highlight(f'{k}{detail["mean_row"]}',f'detail-mean:{detail["model"]}:{d}',comparison_range=comparison,eligible_range=eligibility,eligible_cell=f'{col(j+4)}{cover}',required_count=n)
 value(f'B{start+m}','维度平均得分');row=start+m+3
 value(f'B{row}','任务标签分类');labelstart=row+2;labelend=labelstart+n-1
 for d,title in enumerate(['主体复杂度','主体与背景融合度','综合难度']):
  k=col(d+3);formula(f'{k}{row+1}',f'"{title}（"&COUNTIF({k}{labelstart}:{k}{labelend},"低")&"低"&COUNTIF({k}{labelstart}:{k}{labelend},"中")&"中"&COUNTIF({k}{labelstart}:{k}{labelend},"高")&"高）"')
 for i,task in enumerate(tasks):
  r=labelstart+i;value(f'B{r}',task)
  for j,key in enumerate(('complexity','fusion')):
   if key in c:formula(f'{col(j+3)}{r}',f'INDEX({source(c[key])},MATCH($B{r},{tr},0))')
   else:value(f'{col(j+3)}{r}',labels[i][j])
  formula(f'E{r}',f'INDEX({{"低","中","高"}},INT((MATCH(C{r},{{"低","中","高"}},0)+MATCH(D{r},{{"低","中","高"}},0))/2))')
 value(f'B{labelend+2}','综合难度约定：低/中/高编码1/2/3，两个标签平均后向下取整。')
 row=labelend+5
 for category,title in enumerate(['主体复杂度','主体与背景融合度','综合难度']):
  value(f'B{row}','评测集维度分析-'+title);start=row+2;labelcol=col(category+3);lr=f'${labelcol}${labelstart}:${labelcol}${labelend}'
  for d,level in enumerate(['低','中','高']):
   k=col(d+3);value(f'{k}{row+1}',level);formula(f'{k}{start+m+1}',f'COUNTIF({lr},{k}${row+1})')
   value(f'{col(d+6)}{row+1}',level+'已评任务数')
  for j,model in enumerate(models):
   r=start+j;value(f'B{r}',model);taskcol=col(j+4)
   for d in range(3):
    k=col(d+3);scores=f'${taskcol}$3:${taskcol}${n+2}';sample=f'SUMPRODUCT(({lr}={k}${row+1})*ISNUMBER({scores}))'
    formula(f'{k}{r}',f'IF({sample}=0,"",SUMIF({lr},{k}${row+1},{scores})/{sample})')
    formula(f'{col(d+6)}{r}',sample)
  for d in range(3):
   k=col(d+3);samplecol=col(d+6);average(f'{k}{start+m}',f'{k}{start}:{k}{start+m-1}');highlight(f'{k}{start}:{k}{start+m-1}',f'{title}:{d}',eligible_range=f'{samplecol}{start}:{samplecol}{start+m-1}',required_count_cell=f'{k}{start+m+1}')
  value(f'B{start+m}','维度平均得分');value(f'B{start+m+1}','任务样本数');row=start+m+5
 return {'cells':cells,'highlights':groups,'detail_tables':details,'source_mapping':config,'notes':['Apply highlights as live conditional formatting; all-equal and blanks stay white.','Summary ranking uses complete task coverage; repeated runs are averaged within task first.','Dimension/category comparisons require matching coverage; add eligibility guards when inputs are incomplete.','Add vendor groups by referencing the main task matrix.']}

def main():
 p=argparse.ArgumentParser(description=__doc__);p.add_argument('mapping',type=Path);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
 a.out.parent.mkdir(parents=True,exist_ok=True);a.out.write_text(json.dumps(build(json.loads(a.mapping.read_text())),ensure_ascii=False,indent=2))
if __name__=='__main__':main()
