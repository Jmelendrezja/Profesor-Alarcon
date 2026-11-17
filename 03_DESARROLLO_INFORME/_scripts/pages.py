import sys, json, re, subprocess
pdf, meta, out = sys.argv[1:4]
m = json.load(open(meta))
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
norm = lambda s: re.sub(r"\s+", " ", s).strip()
pages = [norm(subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), pdf, "-"], capture_output=True, text=True).stdout) for i in range(1, n + 1)]
# primera página del cuerpo: aparición de "I. RESUMEN EJECUTIVO" que no sea el índice
start = None
for i, t in enumerate(pages):
    if i > 1 and "ÍNDICE" not in t[:30] and "I. RESUMEN EJECUTIVO" in t:
        start = i; break
res = {}
def find(key, probe):
    for i in range(start, n):
        if probe in pages[i]:
            return i
    return None
for lvl, text in m["toc"]:
    i = find(text, norm(text)[:45])
    if i is not None and not text.startswith("XII."):
        res[text] = str(i - start + 1)
for t in m["tables"] + m["figs"]:
    label = t.split(".")[0]           # "Tabla 3"
    title = norm(t[len(label) + 2:])[:35]
    for i in range(start, n):
        if re.search(re.escape(label) + r"(?!\d)", pages[i]) and title[:20] in pages[i]:
            res[t] = str(i - start + 1); break
missing = [x for _, x in m["toc"] if x not in res] + [x for x in m["tables"] + m["figs"] if x not in res]
print("cuerpo inicia en pág. PDF", start + 1, "| sin página:", missing)
json.dump(res, open(out, "w"), ensure_ascii=False, indent=1)
