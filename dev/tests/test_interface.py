"""Tests de l'interface MapTools (Playwright + Chromium).

    pip install playwright && python -m playwright install chromium
    python dev/tests/test_interface.py          (depuis le dossier MapTools)

Ouvre une copie de MapTools.html qui expose le réseau (window.__net), charge l'exemple et vérifie
les fonctions principales. Affiche OK / ÉCHEC par vérification et le code de sortie 1 en cas d'échec.
"""
import os, re, sys, tempfile, shutil
from playwright.sync_api import sync_playwright
RACINE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src = open(os.path.join(RACINE, 'MapTools.html'), encoding='utf-8').read()
src = re.sub(r"(network=new vis\.Network\(\$\('net'\),\{nodes:nodesDS,edges:edgesDS\},opts2\);)", r"\1window.__net=network;", src)
test_html = os.path.join(RACINE, '_test_interface.html'); open(test_html, 'w', encoding='utf-8').write(src)
URL = 'file:///' + test_html.replace('\\', '/').lstrip('/')
res = []
def ok(nom, cond, detail=''): res.append((nom, bool(cond), detail)); print(('OK    ' if cond else 'ÉCHEC ') + nom + (f' — {detail}' if detail and not cond else ''))
profil = tempfile.mkdtemp()
try:
    with sync_playwright() as p:
        opts = {'viewport': {'width': 1500, 'height': 950}, 'accept_downloads': True}
        if os.environ.get('MAPTOOLS_CHROMIUM'): opts['executable_path'] = os.environ['MAPTOOLS_CHROMIUM']   # navigateur imposé (facultatif)
        ctx = p.chromium.launch_persistent_context(profil, **opts)
        pg = ctx.new_page(); errs = []; pg.on('pageerror', lambda e: errs.append(str(e)))
        nb = lambda: pg.evaluate("window.__net?window.__net.body.data.nodes.length:0")
        outils = lambda b: (pg.evaluate("document.getElementById('toolsMenu').open=true"), pg.click(b))
        pg.goto(URL); pg.wait_for_timeout(500)
        pg.click('#demo'); pg.wait_for_timeout(1600)
        ok('Exemple chargé (4 fichiers)', pg.evaluate("document.getElementById('fileCards').textContent").count('.csv') == 4)
        ok('Carte affichée', nb() > 20, nb())
        for v in ['structure', 'todo', 'off', 'roles', 'mfa', 'guests', 'extshare', 'lic', 'svc', 'admins', 'noowner', 'loops', 'iso']:
            pg.select_option('#preset', v); pg.wait_for_timeout(500); ok(f'Vue rapide {v}', nb() > 0, nb())
        pg.click('#fbReset'); pg.wait_for_timeout(600); ok('Tout réinitialiser', not pg.is_visible('#filtBar'))
        pg.fill('#q', 'Thomas Genoud'); pg.press('#q', 'Enter'); pg.wait_for_timeout(700)
        ok('Panneau de détails en onglets', len(pg.query_selector_all('.d-tabs button')) >= 3)
        pg.click('.d-tabs button[data-tab="acc"]'); pg.select_option('#paTarget', 'Rôle · Exchange Administrator'); pg.click('#paGo'); pg.wait_for_timeout(900)
        ok("Chemin d'héritage", 'Exchange Administrator' in pg.inner_text('.pa-list'))
        pg.click('#paClear'); pg.wait_for_timeout(500)
        pg.click('#reportBtn'); pg.wait_for_timeout(900); t = pg.inner_text('#rpBody')
        for part in ['Synthèse pour la direction', "Plan d'action", '1. Comptes utilisateurs', '2. Groupes', '3. Partages et sites', 'Périmètre et limites']:
            ok(f'Rapport : {part}', part in t)
        b = pg.locator('.rp-acc[data-key="breakglass01@contoso.onmicrosoft.com"]').first; b.click(); pg.wait_for_timeout(300)
        pg.fill('#acJust', 'Test'); pg.click('#acOk'); pg.wait_for_timeout(700)
        ok('Risque accepté', 'Risques acceptés' in pg.inner_text('#rpBody'))
        with pg.expect_download() as d: pg.click('#rpCsv')
        ok('Plan d\'action CSV', d.value.suggested_filename.endswith('.csv'))
        pg.click('#rpClose')
        outils('#exportsBtn'); pg.wait_for_timeout(400)
        for pid in ['mfaok', 'members', 'anon', 'sitesanon']:
            pg.click(f'#xPresets button[data-p="{pid}"]'); pg.wait_for_timeout(250); ok(f'Liste {pid}', pg.is_visible('.xcount'))
        pg.click('#xClose')
        outils('#compareBtn'); pg.click('#cmpDemo'); pg.click('#cmpRun'); pg.wait_for_timeout(800)
        tc = pg.inner_text('#cmpResult'); ok('Avant / après', 'Synthèse pour la direction' in tc and '3. Partages et sites' in tc)
        pg.click('#cmpClose')
        pg.reload(); pg.wait_for_timeout(1800); ok('Restauration après rafraîchissement', 'restaurées' in pg.evaluate("document.getElementById('msg').textContent"))
        pg.evaluate("document.getElementById('secData').open=true"); pg.click('#reset'); pg.wait_for_timeout(200); pg.click('#cfOk'); pg.wait_for_timeout(500)
        ok('Tout effacer', pg.is_visible('#empty'))
        ok('Aucune erreur JavaScript', not errs, '; '.join(errs[:3]))
        ctx.close()
finally:
    shutil.rmtree(profil, ignore_errors=True); os.remove(test_html)
echecs = [r for r in res if not r[1]]
print(f'\n{len(res) - len(echecs)} / {len(res)} vérifications réussies')
sys.exit(1 if echecs else 0)
