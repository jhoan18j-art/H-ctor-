"""Prueba local del almacenamiento offline real en Chromium móvil simulado.

Uso: python test_valet11_browser.py /ruta/a/valet_11_corregido.py
Requiere el paquete playwright (usa su Chromium; o pon la variable CHROMIUM con otra ruta). No sustituye pruebas en un teléfono Android real.
"""
import argparse, ast, base64, json, re, threading, http.server, socketserver, os
from pathlib import Path
from playwright.sync_api import sync_playwright

parser = argparse.ArgumentParser()
parser.add_argument('source', nargs='?', default=str(Path(__file__).with_name('valet_11_corregido.py')))
source = Path(parser.parse_args().source).resolve()
code = source.read_text(encoding='utf-8')
compile(code, str(source), 'exec')

assets_dir = Path('/tmp/clubpark_offline_browser_test'); assets_dir.mkdir(exist_ok=True)
for name in ['app.html', 'db.js', 'index.html', 'logica.js', 'sw.js', 'manifest.json']:
    match = re.search(r"^    '" + re.escape(name) + r"': \(\n(.*?)^    \),", code, re.M | re.S)
    assert match, f'Recurso embebido ausente: {name}'
    encoded = ''.join(ast.literal_eval('(' + match.group(1) + ')'))
    (assets_dir / name).write_bytes(base64.b64decode(encoded))

# Ejecuta las operaciones reales de CPDB/CPSuc desde db.js dentro de Chromium.
(assets_dir / 'harness.html').write_text('''<!doctype html><meta charset="utf-8"><script src="db.js"></script>
<div id="done"></div><script>
window.results={};
(async()=>{try{
 await CPDB.clear('kv'); await CPDB.clear('cola'); await CPDB.clear('locales'); await CPDB.clear('adjuntos');
 await CPDB.put('kv',{sucursal:'SIPIRILY COAPA'},'config');
 await CPDB.put('kv',[[1,5000]],'bloques'); await CPDB.put('kv',1,'siguiente_folio');
 const nums=await Promise.all(Array.from({length:100},()=>CPDB.tomarFolio()));
 window.results.atomic={count:nums.length,unique:new Set(nums).size,min:Math.min(...nums),max:Math.max(...nums)};
 await CPDB.encolar({uid:'offline-1',tipo:'ingreso',folio:'CP-SIP-000101',fotos:{frente:'data:image/jpeg;base64,AA=='},registro:{placas:'ABC123'}});
 const loaded=await CPDB.completo('offline-1'); window.results.attachment=!!(loaded&&loaded.fotos&&loaded.fotos.frente);
 window.results.persistedCount=await CPDB.count('cola');
 document.getElementById('done').textContent='ready';
}catch(e){window.results.error=String(e.stack||e);document.getElementById('done').textContent='error';}})();
</script>''', encoding='utf-8')
(assets_dir / 'worker.html').write_text('<!doctype html><meta charset="utf-8"><script src="db.js"></script><p>ready</p>', encoding='utf-8')

class Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args): pass
os.chdir(assets_dir)
server = socketserver.TCPServer(('127.0.0.1', 0), Quiet)
threading.Thread(target=server.serve_forever, daemon=True).start()
url = f'http://127.0.0.1:{server.server_address[1]}'

try:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, executable_path=os.environ.get('CHROMIUM') or None,
                                    args=['--no-sandbox', '--disable-dev-shm-usage'])
        context = browser.new_context(viewport={'width':393,'height':873}, device_scale_factor=2, is_mobile=True, has_touch=True)
        page = context.new_page(); errors=[]
        cdp = context.new_cdp_session(page); cdp.send('Emulation.setCPUThrottlingRate', {'rate':4})
        page.on('pageerror', lambda e: errors.append(str(e)))
        page.goto(url+'/harness.html')
        page.wait_for_function("document.getElementById('done').textContent==='ready' || document.getElementById('done').textContent==='error'", timeout=20000)
        result = page.evaluate('window.results')
        assert not result.get('error'), result.get('error')
        assert result['atomic'] == {'count':100,'unique':100,'min':1,'max':100}, result['atomic']
        # Two tabs allocate at once; browser Web Locks must prevent cross-tab folio collisions.
        worker=context.new_page(); worker.goto(url+'/worker.html')
        page.evaluate("() => { window.crossTabPromise = Promise.all(Array.from({length:50}, () => CPDB.tomarFolio())); return true; }")
        worker.evaluate("() => { window.crossTabPromise = Promise.all(Array.from({length:50}, () => CPDB.tomarFolio())); return true; }")
        nums_a=page.evaluate('window.crossTabPromise')
        nums_b=worker.evaluate('window.crossTabPromise')
        combined=nums_a+nums_b
        assert len(set(combined))==100 and min(combined)==101 and max(combined)==200, {'unique':len(set(combined)),'min':min(combined),'max':max(combined)}
        worker.close()
        assert result['attachment'] is True
        assert result['persistedCount'] == 1

        # Cut the upload request; the IndexedDB queue must still survive page reload.
        async_fetch = "fetch('/cp_api/sync',{method:'POST',body:'{}'}).catch(()=>null)"
        page.route('**/cp_api/sync', lambda route: route.abort('internetdisconnected'))
        page.evaluate(async_fetch)
        page.reload(); page.wait_for_function("document.getElementById('done').textContent==='ready'", timeout=20000)
        retained = page.evaluate("CPDB.count('cola').then(n=>n)")
        assert retained == 1, f'Queue lost across failed upload/reload: {retained}'
        full = page.evaluate("CPDB.completo('offline-1').then(x=>!!(x&&x.fotos&&x.fotos.frente))")
        assert full is True

        # Positive ACK is the only event that removes the queued item.
        page.evaluate("CPDB.quitar('offline-1')")
        after_ack = page.evaluate("CPDB.count('cola').then(n=>n)")
        assert after_ack == 0

        # A branch change is deferred while old-branch work remains queued.
        page.evaluate("CPDB.encolar({uid:'old-branch',tipo:'ingreso',folio:'CP-SIP-000102'})")
        change = page.evaluate("CPSuc.aplicar({sucursal:'RYOSHI CDMX'}).then(x=>x)")
        still_queued = page.evaluate("CPDB.get('cola','old-branch').then(Boolean)")
        assert change['estado'] == 'en_espera' and still_queued

        # 3,000 lightweight local records under 4x CPU throttling, in a mobile viewport.
        page.evaluate('CPDB.clear("cola")')
        load = page.evaluate('''async()=>{
          const db=await new Promise((ok,no)=>{const r=indexedDB.open('clubpark_offline',2);r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error);});
          await new Promise((ok,no)=>{const t=db.transaction('cola','readwrite'),s=t.objectStore('cola');
            for(let i=0;i<3000;i++) s.put({uid:'load-'+i,tipo:'ingreso',folio:'CP-TEST-'+i,creado:new Date(1700000000000+i).toISOString(),registro:{placas:'ABC123',marca:'SEDAN'}});
            t.oncomplete=ok;t.onerror=()=>no(t.error);});
          const t0=performance.now(); const rows=(await CPDB.all('cola')).sort((a,b)=>String(a.creado).localeCompare(String(b.creado)));
          return {count:rows.length,elapsedMs:Math.round(performance.now()-t0)};
        }''')
        assert load['count']==3000, load
        page.evaluate('CPDB.clear("cola")')
        assert not errors, errors
        print(json.dumps({'python_compile':'PASS','mobile_chromium_indexeddb':'PASS',
                          'parallel_folios':result['atomic'],'cross_tab_folios':{'count':len(combined),'unique':len(set(combined))},'attachment_transaction':'PASS',
                          'failed_sync_queue_survives_reload':'PASS','ack_removes_queue':'PASS',
                          'branch_change_waits_for_old_queue':'PASS','3000_record_indexeddb_load':load,'browser_errors':errors},ensure_ascii=False))
        browser.close()
finally:
    server.shutdown()
