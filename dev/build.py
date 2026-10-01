"""Assemble MapTools.html à partir des sources de dev/src.

    python dev/build.py            (depuis le dossier MapTools)

Produit :
  - MapTools.html                      version locale (bibliothèques dans lib\\)
  - dev/out/MapTools-en-ligne.html     version autonome (bibliothèques sur CDN, exemple chargé si rien n'est mémorisé)
"""
import os, re, subprocess, sys
ICI = os.path.dirname(os.path.abspath(__file__)); RACINE = os.path.dirname(ICI)
lire = lambda p: open(os.path.join(ICI, 'src', p), encoding='utf-8').read()
t = lire('MapTools.src.html').replace('__CSS__', lire('styles.css')).replace('__DEMO__', lire('demo-data.js'))
CDN = {'__PAPA__': 'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js',
       '__VIS__': 'https://cdn.jsdelivr.net/npm/vis-network@9.1.9/standalone/umd/vis-network.min.js'}
local = t.replace('__PAPA__', 'lib/papaparse.min.js').replace('__VIS__', 'lib/vis-network.min.js')
open(os.path.join(RACINE, 'MapTools.html'), 'w', encoding='utf-8').write(local)
ligne = t
for k, v in CDN.items(): ligne = ligne.replace(k, v)
ligne = ligne.replace("\nrestore();\n", "\n// Version en ligne : charge les données fictives si rien n'est mémorisé\nrestore().then(()=>{ if(!files.length) $('demo').click(); });\n")
os.makedirs(os.path.join(ICI, 'out'), exist_ok=True)
open(os.path.join(ICI, 'out', 'MapTools-en-ligne.html'), 'w', encoding='utf-8').write(ligne)
# Contrôle de syntaxe JavaScript si Node.js est disponible
js = re.findall(r'<script>\n(.*?)</script>', local, re.S)[0]
try:
    tmp = os.path.join(ICI, 'out', '_check.js'); open(tmp, 'w', encoding='utf-8').write(js)
    r = subprocess.run(['node', '--check', tmp], capture_output=True, text=True); os.remove(tmp)
    print('Syntaxe JavaScript :', 'OK' if r.returncode == 0 else r.stderr)
except FileNotFoundError:
    print('Node.js absent : contrôle de syntaxe ignoré')
print('MapTools.html et dev/out/MapTools-en-ligne.html générés')
