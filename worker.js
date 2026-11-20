// 분석 계산은 UI 스레드와 분리된 Web Worker의 Pyodide(Python)에서 실행합니다.
const PYODIDE_URL = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let pyodide = null;
let booting = null;
let jobs = Promise.resolve();

async function boot(){
  if (pyodide) return pyodide;
  if (!booting) booting = (async () => {
    postMessage({type:'status',text:'Pyodide Python 실행 환경 준비 중… (최초 1회)'});
    importScripts(PYODIDE_URL + 'pyodide.js');
    const py = await loadPyodide({indexURL: PYODIDE_URL});
    await py.loadPackage(['numpy','micropip']);
    postMessage({type:'status',text:'Excel 읽기 라이브러리 설치 중…'});
    await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
    const source = await fetch('./model.py');
    if (!source.ok) throw Error('model.py 읽기 실패: HTTP ' + source.status);
    await py.runPythonAsync(await source.text());
    pyodide = py;
    return py;
  })().catch(error => {booting = null; throw error;});
  return booting;
}

async function handle({id, type, payload}){
  try{
    const py = await boot();
    let result;
    if (type === 'init'){
      postMessage({type:'status',text:'원본 Excel 12개 시트 읽기 및 행렬계수 검증 중…'});
      py.globals.set('raw_input', new Uint8Array(payload.buffer));
      py.globals.set('n_input', payload.count);
      result = await py.runPythonAsync('initialize(raw_input.to_py(), n_input)');
      py.runPython('del raw_input, n_input');
    } else if (type === 'explore'){
      py.globals.set('sheet_input', payload.sheet);
      py.globals.set('row_input', payload.row || '');
      result = py.runPython('explore(sheet_input, row_input)');
    } else if (type === 'simulate'){
      py.globals.set('scenario_input', JSON.stringify(payload));
      result = py.runPython('simulate(scenario_input)');
    } else {
      throw Error('알 수 없는 작업입니다: '+type);
    }
    postMessage({id, ok:true, result:JSON.parse(result)});
  }catch(error){
    postMessage({id, ok:false, error:String(error?.message || error)});
  }
}

onmessage = event => {
  // 하나의 Python 인터프리터에서의 작업은 순서대로 처리합니다.
  jobs = jobs.then(() => handle(event.data)).catch(error => {
    postMessage({id:event.data.id,ok:false,error:String(error)});
  });
};
