import re,sys,json
from xml.sax.saxutils import escape
def edit(path, groups):
    """groups: list of (start,end,text) 1-based inclusive over <a:t> occurrences"""
    s=open(path,encoding='utf-8').read()
    spans=[m.span() for m in re.finditer(r'<a:t>[^<]*</a:t>|<a:t/>',s)]
    repl={}
    for a,b,text in groups:
        repl[a]=escape(text)
        for k in range(a+1,b+1): repl[k]=''
    out=[];last=0
    for i,(x,y) in enumerate(spans,1):
        out.append(s[last:x])
        out.append(f'<a:t>{repl[i]}</a:t>' if i in repl else s[x:y])
        last=y
    out.append(s[last:]); open(path,'w',encoding='utf-8').write(''.join(out))
if __name__=='__main__':
    edit(sys.argv[1], json.load(open(sys.argv[2],encoding='utf-8')))
