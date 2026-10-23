"""Install a local, pinned OCR runtime and verified models. Nothing global."""
from pathlib import Path
import argparse,concurrent.futures,hashlib,importlib,importlib.metadata,json,os,platform,shutil,subprocess,sys,tempfile,urllib.request,uuid

sys.dont_write_bytecode=True
ROOT=Path(__file__).resolve().parent.parent
CORE={'PyMuPDF':'2.0','python-docx':'1.2.0','python-pptx':'1.0.2'}

def sha256(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for block in iter(lambda:f.read(1024*1024),b''):h.update(block)
    return h.hexdigest()

def verify_models(root):
    for model in json.loads((root/'vendor/ocr_models/models.json').read_text(encoding='utf-8')):
        path=root/'vendor/ocr_models'/model['file']
        if not path.is_file() or path.stat().st_size!=model['bytes'] or sha256(path)!=model['sha256']:
            raise RuntimeError('Missing or invalid model: '+model['file'])

def check(root):
    sys.path[:0]=[str(root/'vendor'),str(root/'vendor/ocr')]
    manifest=json.loads((root/'vendor/ocr-runtime.json').read_text(encoding='utf-8'))
    expected={**CORE,**{p['name']:p['version'] for p in manifest['packages']}}
    for name,version in expected.items():
        actual=importlib.metadata.version(name)
        if actual!=version:raise RuntimeError(f'{name}: expected {version}, found {actual}')
    for name in ('pymupdf','docx','pptx','rapidocr','onnxruntime','cv2','PIL','numpy'):importlib.import_module(name)
    verify_models(root)
    print('Runtime imports, pinned versions and model integrity: OK')

def download_model(model,folder):
    dest=folder/model['file']
    if dest.exists() and dest.stat().st_size==model['bytes'] and sha256(dest)==model['sha256']:
        print('Verified existing model:',model['file'],flush=True);return
    partial=folder/(model['file']+'.part-'+uuid.uuid4().hex)
    try:
        request=urllib.request.Request(model['source'],headers={'User-Agent':'pdf-to-docx-pptx-installer'})
        with urllib.request.urlopen(request,timeout=90) as source,partial.open('wb') as target:shutil.copyfileobj(source,target)
        if partial.stat().st_size!=model['bytes'] or sha256(partial)!=model['sha256']:
            raise RuntimeError('Model integrity failure: '+model['file'])
        os.replace(partial,dest)
        print('Downloaded and verified:',model['file'],flush=True)
    finally:
        if partial.exists():partial.unlink()

def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--check-only',action='store_true',help='Check installed versions, imports and model hashes without downloading')
    ap.add_argument('--skill-root',type=Path,default=ROOT,help='Skill folder to install/check; defaults to this repository')
    ap.add_argument('--workers',type=int,choices=range(1,9),default=3,metavar='1-8')
    args=ap.parse_args();root=args.skill_root.resolve()
    if args.check_only:check(root);return
    if sys.version_info[:2]!=(3,12) or sys.platform!='win32' or platform.machine().lower() not in ('amd64','x86_64'):
        ap.error('This pinned runtime targets 64-bit Windows CPython 3.12. Other environments are not validated.')
    folder=root/'vendor/ocr_models';folder.mkdir(parents=True,exist_ok=True)
    manifest=json.loads((root/'vendor/ocr-runtime.json').read_text(encoding='utf-8'))
    requirements=[f'{name}=={version}' for name,version in CORE.items()]+[p['name']+'=='+p['version'] for p in manifest['packages']]
    models=json.loads((folder/'models.json').read_text(encoding='utf-8'))
    with tempfile.TemporaryDirectory(prefix='.install-',dir=root) as scratch:
        work=Path(scratch);req=work/'requirements.txt';req.write_text('\n'.join(requirements)+'\n',encoding='utf-8')
        env=os.environ.copy();env.update(TEMP=scratch,TMP=scratch,PIP_CACHE_DIR=str(work/'pip-cache'),PYTHONDONTWRITEBYTECODE='1')
        # Model tasks run concurrently with dependency installation.
        with concurrent.futures.ThreadPoolExecutor(max_workers=args.workers) as pool:
            downloads=[pool.submit(download_model,m,folder) for m in models]
            subprocess.run([sys.executable,'-m','pip','install','--disable-pip-version-check','--only-binary=:all:','--no-compile','--no-cache-dir','--upgrade','--target',str(root/'vendor/ocr'),'-r',str(req)],env=env,check=True)
            for future in downloads:future.result()
    check(root)

if __name__=='__main__':main()
