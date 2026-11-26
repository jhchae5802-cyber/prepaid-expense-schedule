"""선급비용 명세서 웹앱 (로컬 전용)

실행:  python app.py   →  브라우저에서 http://127.0.0.1:5055
원자료 엑셀을 올리고 기준일을 고르면 명세서를 만들고, 검증 결과를 화면에 보여 주며, 엑셀로 내려받을 수 있다.
"""
import base64, datetime as dt, json, os, re, shutil, time, uuid, webbrowser, threading
from flask import Flask, jsonify, request, send_file, render_template, abort

import schedule_core as core

HERE = os.path.dirname(os.path.abspath(__file__))
# Vercel 서버리스: /tmp 만 쓸 수 있고 요청마다 다른 인스턴스가 처리할 수 있다.
# → 결과 엑셀을 응답에 담아 보내고, '작업 폴더에 저장'은 쓰지 않는다.
CLOUD = bool(os.environ.get('VERCEL'))
WORK = '/tmp/_work' if CLOUD else os.path.join(HERE, '_work')   # 업로드 원자료와 생성 결과 (세션별 폴더)
SAVE_DIR = os.path.dirname(HERE)              # '작업 폴더에 저장' 대상 = 원자료가 있는 프로젝트 폴더
PORT = 5055

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024


def _cleanup(max_age_h=24):
    os.makedirs(WORK, exist_ok=True)
    now = time.time()
    for d in os.listdir(WORK):
        p = os.path.join(WORK, d)
        if os.path.isdir(p) and now - os.path.getmtime(p) > max_age_h * 3600:
            shutil.rmtree(p, ignore_errors=True)


def _job_dir(job):
    if not re.fullmatch(r'[0-9a-f]{32}', job or ''):
        abort(400, '잘못된 작업 ID')
    p = os.path.join(WORK, job)
    if not os.path.isdir(p):
        abort(404, '작업을 찾을 수 없습니다. 원자료를 다시 올려 주세요.')
    return p


def _meta(p, data=None):
    f = os.path.join(p, 'meta.json')
    if data is None:
        with open(f, encoding='utf-8') as fh: return json.load(fh)
    with open(f, 'w', encoding='utf-8') as fh: json.dump(data, fh, ensure_ascii=False)


@app.get('/')
def index():
    return render_template('index.html')


@app.post('/api/build')
def api_build():
    """multipart: file(원자료) 또는 form job(이전 업로드), base_date(선택, YYYY-MM-DD)"""
    base = (request.form.get('base_date') or '').strip() or None
    if base:
        try: dt.date.fromisoformat(base)
        except ValueError: return jsonify(error='기준일 형식이 올바르지 않습니다 (YYYY-MM-DD).'), 400

    up = request.files.get('file')
    if up and up.filename:
        if not up.filename.lower().endswith(('.xlsx', '.xlsm')):
            return jsonify(error='.xlsx 파일만 올릴 수 있습니다.'), 400
        job = uuid.uuid4().hex
        p = os.path.join(WORK, job); os.makedirs(p)
        up.save(os.path.join(p, 'source.xlsx'))
        _meta(p, {'source_name': os.path.basename(up.filename)})
    else:
        job = request.form.get('job', '')
        p = _job_dir(job)

    meta = _meta(p)
    out = os.path.join(p, 'result.xlsx')
    try:
        res = core.build(os.path.join(p, 'source.xlsx'), out, base_date=base)
    except core.ScheduleError as e:
        return jsonify(error=str(e)), 422
    except Exception as e:  # 손상된 파일 등
        return jsonify(error=f'처리 중 오류: {type(e).__name__}: {e}'), 500
    meta['filename'] = res['filename']; _meta(p, meta)
    res.update(job=job, source_name=meta['source_name'], cloud=CLOUD)
    if CLOUD:
        with open(out, 'rb') as fh:
            res['xlsx_b64'] = base64.b64encode(fh.read()).decode('ascii')
        shutil.rmtree(p, ignore_errors=True)
    else:
        res['exists_in_folder'] = os.path.exists(os.path.join(SAVE_DIR, res['filename']))
    return jsonify(res)


@app.get('/api/download/<job>')
def api_download(job):
    p = _job_dir(job); meta = _meta(p)
    return send_file(os.path.join(p, 'result.xlsx'), as_attachment=True, download_name=meta['filename'])


@app.post('/api/save/<job>')
def api_save(job):
    """생성한 명세서를 원자료 폴더(프로젝트 폴더)에 저장. 같은 이름이 있으면 overwrite=1 일 때만 덮어씀"""
    if CLOUD:
        abort(404)
    p = _job_dir(job); meta = _meta(p)
    dest = os.path.join(SAVE_DIR, meta['filename'])
    if os.path.exists(dest) and request.args.get('overwrite') != '1':
        return jsonify(error='exists', path=dest), 409
    try:
        shutil.copyfile(os.path.join(p, 'result.xlsx'), dest)
    except PermissionError:
        return jsonify(error=f'저장 실패: {meta["filename"]} 파일이 Excel에서 열려 있으면 닫고 다시 시도하세요.'), 423
    return jsonify(ok=True, path=dest)


if __name__ == '__main__':
    _cleanup()
    url = f'http://127.0.0.1:{PORT}'
    print(f'선급비용 명세서 웹앱: {url}  (종료: Ctrl+C)')
    if not os.environ.get('NO_BROWSER'):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)
