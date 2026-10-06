"""Write cached values (computed by the `formulas` engine) into formula cells so previews show numbers."""
import sys, zipfile, re, json, shutil
from xml.sax.saxutils import escape
import formulas
src = sys.argv[1]
xl = formulas.ExcelModel().loads(src).finish(); sol = xl.calculate()
vals = {}
for k, x in sol.items():
    v = x.value[0, 0] if hasattr(x, 'value') else x
    m = re.match(r"^'?\[[^\]]+\](.+?)'?!([A-Z]+\d+)$", str(k))
    if not m: continue
    vals[(m.group(1).upper(), m.group(2))] = v
z = zipfile.ZipFile(src)
wbxml = z.read('xl/workbook.xml').decode()
rels = z.read('xl/_rels/workbook.xml.rels').decode()
rid2file = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="/?(?:xl/)?([^"]+)"', rels))
rid2file.update({a: b for b, a in re.findall(r'Target="/?(?:xl/)?([^"]+)"[^>]*Id="(rId\d+)"', rels)})
sheets = {}
for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml):
    sheets['xl/' + rid2file[rid]] = name.upper()
out = src + '.tmp'
zo = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
n = 0; errs = 0
for item in z.infolist():
    data = z.read(item.filename)
    if item.filename in sheets:
        sname = sheets[item.filename]
        xml = data.decode()
        def rep(m):
            global n, errs
            attrs, f = m.group(1), m.group(2)
            ref = re.search(r'r="([A-Z]+\d+)"', attrs).group(1)
            v = vals.get((sname, ref))
            if v is None: return m.group(0)
            attrs = re.sub(r'\s+t="[^"]*"', '', attrs)
            try:
                fv = float(v); n += 1
                return f'<c{attrs}>{f}<v>{repr(fv)}</v></c>'
            except (TypeError, ValueError):
                s = str(v)
                if s.startswith('#'): errs += 1
                n += 1
                return f'<c{attrs} t="str">{f}<v>{escape(s)}</v></c>'
        xml = re.sub(r'<c([^>]*)>(<f>.*?</f>)(?:<v>.*?</v>|<v\s*/>)?</c>', rep, xml)
        data = xml.encode()
    zo.writestr(item, data)
zo.close(); shutil.move(out, src)
print(json.dumps({"cached_values_written": n, "errors": errs}))
