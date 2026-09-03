<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<title>Dashboard Clínico</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/lz-string@1.5.0/libs/lz-string.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:#eef3f7;color:#2c3e50;padding:20px}
header{background:linear-gradient(135deg,#0a6ebd,#14a085);color:#fff;padding:20px 30px;border-radius:10px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
header h1{font-size:24px}
header .actions button{background:#fff;color:#0a6ebd;border:none;padding:10px 18px;margin-left:8px;border-radius:6px;cursor:pointer;font-weight:700;transition:.2s}
header .actions button:hover{background:#f0f0f0}
header .actions .btn-excel{background:#1d6f42;color:#fff}
header .actions .btn-share{background:#e67e22;color:#fff}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
@media(max-width:900px){.grid{grid-template-columns:1fr}}
.card{background:#fff;border-radius:10px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.card h2{font-size:16px;color:#0a6ebd;margin-bottom:15px;border-bottom:2px solid #e0e7ef;padding-bottom:8px;display:flex;justify-content:space-between;align-items:center}
.btn-sm{background:#14a085;color:#fff;border:none;padding:5px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:700}
.btn-sm:hover{background:#0e7a64}
form .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:12px}
form label{display:block;font-size:12px;color:#555;margin-bottom:4px;font-weight:600}
form input,form select,form textarea{width:100%;padding:8px 10px;border:1px solid #cfd8e3;border-radius:6px;font-size:14px;background:#fafbfc}
form input:focus,form select:focus,form textarea:focus{outline:none;border-color:#0a6ebd;background:#fff}
form textarea{resize:vertical;min-height:60px}
.form-actions{display:flex;gap:10px;margin-top:10px}
.btn-primary{background:#0a6ebd;color:#fff;border:none;padding:10px 20px;border-radius:6px;cursor:pointer;font-weight:700}
.btn-secondary{background:#95a5a6;color:#fff;border:none;padding:10px 20px;border-radius:6px;cursor:pointer}
.filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:20px}
.filters select,.filters input{padding:8px;border:1px solid #cfd8e3;border-radius:6px;font-size:13px}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}
.stat{background:#fff;padding:15px;border-radius:10px;text-align:center;box-shadow:0 2px 6px rgba(0,0,0,.05);border-left:4px solid #0a6ebd}
.stat .num{font-size:24px;font-weight:700;color:#0a6ebd}
.stat .lbl{font-size:12px;color:#7f8c8d;margin-top:4px}
.table-wrap{overflow-x:auto;background:#fff;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.06)}
table{width:100%;border-collapse:collapse;font-size:13px;min-width:1600px}
th{background:#0a6ebd;color:#fff;padding:10px 6px;text-align:left;font-weight:600;position:sticky;top:0}
td{padding:4px;border-bottom:1px solid #eaeef3;vertical-align:middle}
tr:nth-child(even){background:#f7fafc}
tr:hover{background:#e8f1fa}
td input,td select{width:100%;border:1px solid transparent;background:transparent;padding:6px 4px;font-size:13px;font-family:inherit;color:#2c3e50;border-radius:4px;transition:.2s}
td input:hover,td select:hover{background:#f0f6fb;border-color:#cfd8e3}
td input:focus,td select:focus{outline:none;border-color:#0a6ebd;background:#fff;box-shadow:0 0 0 2px rgba(10,110,189,.15)}
td select{cursor:pointer;appearance:auto}
.row-num{background:#eef3f7;color:#7f8c8d;font-weight:700;text-align:center;width:40px}
.del-btn{background:#e74c3c;color:#fff;border:none;padding:5px 10px;border-radius:4px;cursor:pointer;font-size:12px}
.chart-box{position:relative;height:280px}
.aviso{background:#fff3cd;border-left:4px solid #f39c12;padding:10px 15px;border-radius:6px;font-size:13px;margin-bottom:15px}
.export-bar{background:#fff;padding:15px 20px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.export-bar strong{color:#0a6ebd;margin-right:10px}
.export-bar button{background:#14a085;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:700;font-size:13px}
.export-bar button:hover{background:#0e7a64}
.export-bar button.secondary{background:#6c757d}
.export-bar button.btn-excel{background:#1d6f42}
.export-bar button.btn-share{background:#e67e22}
/* IMPORT BAR */
.import-bar{background:#d4edda;border:2px solid #28a745;padding:12px 20px;border-radius:10px;margin-bottom:20px;display:none;align-items:center;gap:10px;flex-wrap:wrap}
.import-bar.active{display:flex}
.import-bar label{font-weight:700;color:#155724;font-size:14px}
.import-bar textarea{flex:1;min-width:200px;padding:8px;border:1px solid #28a745;border-radius:6px;font-family:monospace;font-size:12px;resize:none;height:40px}
.import-bar button{background:#28a745;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:700}
/* MODAL */
.modal-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.6);z-index:1000;justify-content:center;align-items:center;padding:20px}
.modal-overlay.active{display:flex}
.modal{background:#fff;border-radius:12px;padding:30px;max-width:700px;width:100%;max-height:90vh;overflow-y:auto;box-shadow:0 10px 40px rgba(0,0,0,.3)}
.modal h2{color:#0a6ebd;margin-bottom:10px;border-bottom:2px solid #e0e7ef;padding-bottom:10px}
.modal .close-btn{float:right;background:#e74c3c;color:#fff;border:none;width:30px;height:30px;border-radius:50%;cursor:pointer;font-size:16px;font-weight:700}
.m-card{margin-bottom:20px;padding:20px;background:#f7fafc;border-radius:10px;border-left:5px solid #0a6ebd}
.m-card h3{color:#0a6ebd;font-size:15px;margin-bottom:8px}
.m-card p{font-size:13px;color:#555;margin-bottom:12px}
.m-card .big-btn{display:block;width:100%;padding:14px;border:none;border-radius:8px;font-size:16px;font-weight:700;cursor:pointer;color:#fff;transition:.2s}
.m-card .big-btn:hover{opacity:.9}
.m-card .big-btn.green{background:#28a745}
.m-card .big-btn.blue{background:#0a6ebd}
.m-card .big-btn.orange{background:#e67e22}
.code-area{background:#fff;border:2px dashed #0a6ebd;padding:12px;border-radius:8px;font-family:monospace;font-size:11px;word-break:break-all;max-height:120px;overflow-y:auto;margin:10px 0;user-select:all}
.copy-row{display:flex;gap:8px;margin-top:8px}
.copy-row button{background:#0a6ebd;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:700;font-size:13px}
.copy-row button:hover{background:#085a9c}
.tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;margin-left:6px}
.tag-ok{background:#d4edda;color:#155724}
.tag-warn{background:#fff3cd;color:#856404}
.info-box{background:#d1ecf1;border-left:4px solid #17a2b8;padding:12px;border-radius:6px;font-size:12px;color:#0c5460;margin-top:15px}
.toast{position:fixed;bottom:30px;right:30px;background:#27ae60;color:#fff;padding:12px 24px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,.2);font-weight:700;z-index:2000;opacity:0;transform:translateY(20px);transition:.3s}
.toast.show{opacity:1;transform:translateY(0)}
</style>
</head>
<body>

<header>
  <div>
    <h1>🏥 Dashboard Clínico</h1>
    <small>Ficha de Registo Clínico</small>
  </div>
  <div class="actions">
    <button class="btn-share" onclick="openShareModal()">🔗 Partilhar</button>
    <button class="btn-excel" onclick="exportExcel()">📊 Excel</button>
    <button onclick="exportWord()">📄 Word</button>
    <button onclick="exportJSON()">💾 Guardar</button>
    <button onclick="clearAll()">🗑️ Limpar</button>
  </div>
</header>

<!-- BARRA DE IMPORTAÇÃO -->
<div class="import-bar" id="importBar">
  <label>📥 Colar código de partilha:</label>
  <textarea id="importCode" placeholder="Cole aqui o código recebido..."></textarea>
  <button onclick="importFromCode()">✅ Importar</button>
  <button onclick="document.getElementById('importBar').classList.remove('active')" style="background:#e74c3c">✕</button>
</div>

<!-- FORMULÁRIO -->
<div class="card" style="margin-bottom:20px">
  <h2>📝 Nova Ficha de Registo Clínico</h2>
  <form id="formRegisto">
    <div class="row">
      <div><label>ID (automático)</label><input type="text" id="idDoente" readonly></div>
      <div><label>Data/Hora Entrada</label><input type="datetime-local" id="dataEntrada" required></div>
      <div><label>Sexo</label><select id="sexo" required><option value="">--</option><option>Masculino</option><option>Feminino</option></select></div>
      <div><label>Idade</label><input type="number" id="idade" min="0" required></div>
      <div><label>Unidade</label><select id="unidade"><option value="anos">Anos</option><option value="meses">Meses</option><option value="dias">Dias</option></select></div>
      <div><label>Faixa Etária (auto)</label><input type="text" id="faixaEtaria" readonly></div>
    </div>
    <div class="row">
      <div><label>Proveniência</label><select id="proveniencia" required><option value="">--</option><option>Urgência</option><option>Consulta Externa</option><option>Transferência de outro hospital</option><option>Referenciação</option><option>Domicílio</option><option>Outro</option></select></div>
      <div><label>Diagnóstico (livre)</label><input type="text" id="diagnostico" placeholder="Escreva o diagnóstico..." required></div>
      <div><label>Destino</label><select id="destino" required><option value="">--</option><option>Internado</option><option>Alta</option><option>Consulta Externa</option><option>Transferência</option></select></div>
      <div><label>Data Saída</label><input type="datetime-local" id="dataSaida"></div>
      <div><label>Médico</label><input type="text" id="medico" required></div>
    </div>
    <div><label>Observações</label><textarea id="observacoes" placeholder="Notas clínicas..."></textarea></div>
    <div class="form-actions">
      <button type="submit" class="btn-primary">➕ Adicionar</button>
      <button type="button" class="btn-secondary" onclick="document.getElementById('formRegisto').reset();gerarID()">Limpar</button>
    </div>
  </form>
</div>

<div class="card" style="margin-bottom:20px">
  <h2>🔍 Filtros</h2>
  <div class="filters">
    <select id="fSexo" onchange="render()"><option value="">Todos os sexos</option><option>Masculino</option><option>Feminino</option></select>
    <select id="fDestino" onchange="render()"><option value="">Todos os destinos</option><option>Internado</option><option>Alta</option><option>Consulta Externa</option><option>Transferência</option></select>
    <select id="fProveniencia" onchange="render()"><option value="">Todas proveniências</option><option>Urgência</option><option>Consulta Externa</option><option>Transferência de outro hospital</option><option>Referenciação</option><option>Domicílio</option><option>Outro</option></select>
    <select id="fFaixa" onchange="render()"><option value="">Todas faixas</option></select>
    <input type="text" id="fDiag" placeholder="Filtrar diagnóstico..." oninput="render()">
    <input type="text" id="fMedico" placeholder="Filtrar médico..." oninput="render()">
  </div>
</div>

<div class="stats" id="statsBox"></div>

<div class="export-bar">
  <strong>📊 Exportar / Partilhar:</strong>
  <button class="btn-share" onclick="openShareModal()">🔗 Partilhar</button>
  <button onclick="document.getElementById('importBar').classList.add('active')">📥 Importar Código</button>
  <button class="btn-excel" onclick="exportExcel()">📗 Excel</button>
  <button onclick="exportAllCharts()">📥 Gráficos PNG</button>
  <button onclick="exportWord()">📄 Word</button>
  <button onclick="exportStandaloneHTML()">📦 HTML com Dados</button>
  <button class="secondary" onclick="window.print()">🖨️ PDF</button>
</div>

<div class="grid">
  <div class="card"><h2>👥 Sexo <button class="btn-sm" onclick="exportChart('chartSexo','Sexo')">📷</button></h2><div class="chart-box"><canvas id="chartSexo"></canvas></div></div>
  <div class="card"><h2>🎯 Destino <button class="btn-sm" onclick="exportChart('chartDestino','Destino')">📷</button></h2><div class="chart-box"><canvas id="chartDestino"></canvas></div></div>
  <div class="card"><h2>🏥 Proveniência <button class="btn-sm" onclick="exportChart('chartProv','Proveniencia')">📷</button></h2><div class="chart-box"><canvas id="chartProv"></canvas></div></div>
  <div class="card"><h2>📊 Faixa Etária <button class="btn-sm" onclick="exportChart('chartFaixa','Faixa')">📷</button></h2><div class="chart-box"><canvas id="chartFaixa"></canvas></div></div>
  <div class="card" style="grid-column:span 2"><h2>🩺 Diagnóstico <button class="btn-sm" onclick="exportChart('chartDiag','Diagnostico')">📷</button></h2><div class="chart-box"><canvas id="chartDiag"></canvas></div></div>
</div>

<div class="card" style="margin-top:20px">
  <h2>📋 Registos (200 linhas)</h2>
  <div class="aviso">💡 Todos os campos são editáveis inline.</div>
  <div class="table-wrap">
    <table><thead><tr><th>#</th><th>ID</th><th>Entrada</th><th>Sexo</th><th>Idade</th><th>Un.</th><th>Faixa</th><th>Proveniência</th><th>Diagnóstico</th><th>Destino</th><th>Saída</th><th>Médico</th><th>Obs.</th><th>×</th></tr></thead><tbody id="tbody"></tbody></table>
  </div>
</div>

<!-- MODAL DE PARTILHA -->
<div class="modal-overlay" id="shareModal">
  <div class="modal">
    <button class="close-btn" onclick="closeShareModal()">×</button>
    <h2>🔗 Partilhar entre Dispositivos</h2>
    <p style="color:#7f8c8d;font-size:13px;margin-bottom:20px">3 métodos disponíveis. O Método 1 funciona <b>sempre</b>, sem internet.</p>

    <!-- MÉTODO 1 -->
    <div class="m-card">
      <h3>📦 Método 1: Ficheiro HTML com Dados <span class="tag tag-ok">FUNCIONA SEMPRE</span></h3>
      <p>Gera um ficheiro <b>.html</b> completo com todos os dados embutidos. Envie por <b>email, WhatsApp, USB, Bluetooth</b>. O destinatário abre no browser e tem tudo — dashboard, gráficos, tabela editável. <b>Não precisa de internet.</b></p>
      <button class="big-btn green" onclick="exportStandaloneHTML();showToast('✅ HTML descarregado! Envie este ficheiro para o outro dispositivo.')">📥 Descarregar HTML com Dados</button>
    </div>

    <!-- MÉTODO 2 -->
    <div class="m-card">
      <h3>📋 Método 2: Código de Partilha <span class="tag tag-ok">FUNCIONA SEMPRE</span></h3>
      <p>Gera um código de texto comprimido. Copie e envie por <b>chat, SMS, email</b>. No outro dispositivo, abra o dashboard e cole o código na barra "📥 Importar Código".</p>
      <button class="big-btn blue" onclick="generateShareCode()">🔑 Gerar Código de Partilha</button>
      <div id="codeResult" style="display:none">
        <p style="margin-top:12px"><b>Código para partilhar:</b> (copie tudo)</p>
        <div class="code-area" id="shareCodeBox"></div>
        <div class="copy-row">
          <button onclick="copyCode()">📋 Copiar Código Completo</button>
        </div>
        <p style="font-size:11px;color:#7f8c8d;margin-top:8px">
          Tamanho: <span id="codeSize">-</span> | Registos: <span id="codeRegs">-</span><br>
          <b>No outro dispositivo:</b> Abra o dashboard → clique em "📥 Importar Código" → cole → clique "✅ Importar"
        </p>
      </div>
    </div>

    <!-- MÉTODO 3 -->
    <div class="m-card">
      <h3>🌐 Método 3: Link Online <span class="tag tag-warn">REQUER HOSPEDAGEM</span></h3>
      <p>Para ter um link clicável que abre noutro dispositivo, precisa hospedar o HTML online (gratuito). Siga estes passos:</p>
      <ol style="font-size:13px;color:#555;padding-left:20px;line-height:1.8">
        <li>Vá a <a href="https://pages.github.com" target="_blank" style="color:#0a6ebd">pages.github.com</a> e crie uma conta GitHub (grátis)</li>
        <li>Crie um repositório chamado <code>dashboard</code></li>
        <li>Faça upload do ficheiro HTML (use o Método 1 para gerar)</li>
        <li>Em Settings → Pages → ative o GitHub Pages</li>
        <li>O seu dashboard ficará em: <code>https://SEU_USER.github.io/dashboard/ficheiro.html</code></li>
      </ol>
      <p style="margin-top:10px">Depois de hospedado, use este botão para gerar um link com os dados:</p>
      <div class="copy-row">
        <input type="text" id="hostedUrl" placeholder="Cole o URL do seu dashboard hospedado..." style="flex:1;padding:8px;border:1px solid #cfd8e3;border-radius:6px;font-size:12px">
        <button onclick="generateHostedLink()">🔗 Gerar Link</button>
      </div>
      <div id="hostedResult" style="display:none;margin-top:10px">
        <div class="copy-row">
          <input type="text" id="finalLink" readonly style="flex:1;padding:8px;border:1px solid #28a745;border-radius:6px;font-size:12px;background:#d4edda">
          <button onclick="copyFinalLink()">📋 Copiar</button>
        </div>
      </div>
    </div>

    <div class="info-box">
      <strong>💡 Recomendação:</strong> Use o <b>Método 1</b> para partilha rápida (funciona sempre). Use o <b>Método 3</b> se quiser um link permanente para uso diário numa equipa.
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
let registos=JSON.parse(localStorage.getItem('registosClinicos')||'[]');
let charts={};

window.addEventListener('load',()=>{
  verificarImportacaoURL();
  gerarID();
  document.getElementById('dataEntrada').value=new Date().toISOString().slice(0,16);
  calcularFaixa();
  document.getElementById('idade').addEventListener('input',calcularFaixa);
  document.getElementById('unidade').addEventListener('change',calcularFaixa);
  render();
});

function verificarImportacaoURL(){
  const h=window.location.hash;
  if(h.startsWith('#d=')){
    try{
      const json=LZString.decompressFromEncodedURIComponent(h.substring(3));
      if(json){
        const d=JSON.parse(json);
        if(Array.isArray(d)&&d.length>0){
          if(confirm(d.length+' registos encontrados no link. Importar?')){
            registos=d;salvar();render();
            showToast('✅ '+d.length+' registos importados!');
          }
        }
      }
    }catch(e){console.error(e)}
    history.replaceState(null,'',window.location.pathname);
  }
}

function gerarID(){
  document.getElementById('idDoente').value='PAC-'+new Date().getFullYear()+'-'+String(registos.length+1).padStart(4,'0');
}
function calcularFaixa(){
  document.getElementById('faixaEtaria').value=calcFaixa(document.getElementById('idade').value,document.getElementById('unidade').value);
}
function calcFaixa(id,u){
  let a=parseFloat(id)||0;
  if(u==='meses')a/=12;if(u==='dias')a/=365;
  if(a<1)return'< 1 ano';
  const b=Math.floor(a/5)*5;return b+'-'+(b+4)+' anos';
}

document.getElementById('formRegisto').addEventListener('submit',e=>{
  e.preventDefault();
  registos.push({
    id:document.getElementById('idDoente').value,
    dataEntrada:document.getElementById('dataEntrada').value,
    sexo:document.getElementById('sexo').value,
    idade:document.getElementById('idade').value,
    unidade:document.getElementById('unidade').value,
    faixaEtaria:document.getElementById('faixaEtaria').value,
    proveniencia:document.getElementById('proveniencia').value,
    diagnostico:document.getElementById('diagnostico').value,
    destino:document.getElementById('destino').value,
    dataSaida:document.getElementById('dataSaida').value,
    medico:document.getElementById('medico').value,
    observacoes:document.getElementById('observacoes').value
  });
  salvar();e.target.reset();
  document.getElementById('dataEntrada').value=new Date().toISOString().slice(0,16);
  gerarID();calcularFaixa();render();
});

function salvar(){localStorage.setItem('registosClinicos',JSON.stringify(registos))}
function apagar(i){if(confirm('Apagar?')){registos.splice(i,1);salvar();gerarID();render()}}
function clearAll(){if(confirm('Apagar TUDO?')){registos=[];salvar();gerarID();render()}}
function atualizarCampo(i,c,v){
  if(!registos[i])return;registos[i][c]=v;
  if(c==='idade'||c==='unidade')registos[i].faixaEtaria=calcFaixa(registos[i].idade,registos[i].unidade);
  salvar();render();
}

function getFiltrados(){
  const s=document.getElementById('fSexo').value,d=document.getElementById('fDestino').value,
  p=document.getElementById('fProveniencia').value,f=document.getElementById('fFaixa').value,
  dg=document.getElementById('fDiag').value.toLowerCase(),m=document.getElementById('fMedico').value.toLowerCase();
  return registos.map((r,i)=>({...r,_i:i})).filter(r=>
    (!s||r.sexo===s)&&(!d||r.destino===d)&&(!p||r.proveniencia===p)&&(!f||r.faixaEtaria===f)&&
    (!dg||r.diagnostico.toLowerCase().includes(dg))&&(!m||r.medico.toLowerCase().includes(m)));
}

function render(){
  const dados=getFiltrados(),t=dados.length;
  const faixas=[...new Set(registos.map(r=>r.faixaEtaria))].sort();
  const sf=document.getElementById('fFaixa'),at=sf.value;
  sf.innerHTML='<option value="">Todas faixas</option>'+faixas.map(f=>'<option'+(f===at?' selected':'')+'>'+f+'</option>').join('');
  const sexo=contar(dados,'sexo'),dest=contar(dados,'destino'),prov=contar(dados,'proveniencia'),
  faixa=contar(dados,'faixaEtaria'),diag=contar(dados,'diagnostico');
  document.getElementById('statsBox').innerHTML=
    '<div class="stat"><div class="num">'+t+'</div><div class="lbl">Total</div></div>'+
    '<div class="stat"><div class="num">'+(sexo.Masculino||0)+'</div><div class="lbl">Masc.</div></div>'+
    '<div class="stat"><div class="num">'+(sexo.Feminino||0)+'</div><div class="lbl">Fem.</div></div>'+
    '<div class="stat"><div class="num">'+(dest.Internado||0)+'</div><div class="lbl">Internados</div></div>'+
    '<div class="stat"><div class="num">'+(dest.Alta||0)+'</div><div class="lbl">Altas</div></div>'+
    '<div class="stat"><div class="num">'+Object.keys(diag).length+'</div><div class="lbl">Diagnósticos</div></div>';
  upChart('chartSexo','pie',Object.keys(sexo),Object.values(sexo),['#3498db','#e91e63']);
  upChart('chartDestino','doughnut',Object.keys(dest),Object.values(dest),['#f39c12','#27ae60','#3498db','#9b59b6']);
  upChart('chartProv','bar',Object.keys(prov),Object.values(prov),['#0a6ebd'],true);
  upChart('chartFaixa','bar',Object.keys(faixa),Object.values(faixa),['#14a085'],true);
  upDiag(Object.keys(diag),Object.values(diag));
  const tb=document.getElementById('tbody');let h='';
  dados.forEach((r,i)=>{
    const x=r._i;
    h+='<tr><td class="row-num">'+(i+1)+'</td>'+
    '<td><input value="'+esc(r.id)+'" onchange="atualizarCampo('+x+',\'id\',this.value)"></td>'+
    '<td><input type="datetime-local" value="'+(r.dataEntrada||'')+'" onchange="atualizarCampo('+x+',\'dataEntrada\',this.value)"></td>'+
    '<td><select onchange="atualizarCampo('+x+',\'sexo\',this.value)"><option'+(r.sexo==='Masculino'?' selected':'')+'>Masculino</option><option'+(r.sexo==='Feminino'?' selected':'')+'>Feminino</option></select></td>'+
    '<td><input type="number" min="0" value="'+(r.idade||'')+'" onchange="atualizarCampo('+x+',\'idade\',this.value)"></td>'+
    '<td><select onchange="atualizarCampo('+x+',\'unidade\',this.value)"><option value="anos"'+(r.unidade==='anos'?' selected':'')+'>anos</option><option value="meses"'+(r.unidade==='meses'?' selected':'')+'>meses</option><option value="dias"'+(r.unidade==='dias'?' selected':'')+'>dias</option></select></td>'+
    '<td><input value="'+esc(r.faixaEtaria||'')+'" readonly style="background:#eef3f7;color:#7f8c8d"></td>'+
    '<td><select onchange="atualizarCampo('+x+',\'proveniencia\',this.value)">'+['Urgência','Consulta Externa','Transferência de outro hospital','Referenciação','Domicílio','Outro'].map(p=>'<option'+(r.proveniencia===p?' selected':'')+'>'+p+'</option>').join('')+'</select></td>'+
    '<td><input value="'+esc(r.diagnostico||'')+'" onchange="atualizarCampo('+x+',\'diagnostico\',this.value)"></td>'+
    '<td><select onchange="atualizarCampo('+x+',\'destino\',this.value)">'+['Internado','Alta','Consulta Externa','Transferência'].map(d=>'<option'+(r.destino===d?' selected':'')+'>'+d+'</option>').join('')+'</select></td>'+
    '<td><input type="datetime-local" value="'+(r.dataSaida||'')+'" onchange="atualizarCampo('+x+',\'dataSaida\',this.value)"></td>'+
    '<td><input value="'+esc(r.medico||'')+'" onchange="atualizarCampo('+x+',\'medico\',this.value)"></td>'+
    '<td><input value="'+esc(r.observacoes||'')+'" onchange="atualizarCampo('+x+',\'observacoes\',this.value)"></td>'+
    '<td><button class="del-btn" onclick="apagar('+x+')">×</button></td></tr>';
  });
  for(let i=dados.length;i<200;i++)h+='<tr><td class="row-num">'+(i+1)+'</td>'+'<td></td>'.repeat(13)+'</tr>';
  tb.innerHTML=h;
}

function esc(s){return s==null?'':String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function contar(a,c){const o={};a.forEach(r=>{const v=r[c]||'(vazio)';o[v]=(o[v]||0)+1});return o}

function upChart(id,tipo,labels,data,cores,hz=false){
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart(document.getElementById(id).getContext('2d'),{
    type:tipo,data:{labels,datasets:[{data,backgroundColor:cores.length===data.length?cores:data.map(()=>cores[0]),borderWidth:1}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:hz?'y':'x',
      plugins:{legend:{position:tipo==='pie'||tipo==='doughnut'?'right':'top'},
        tooltip:{callbacks:{label:c=>{const t=c.dataset.data.reduce((a,b)=>a+b,0);return c.label+': '+c.raw+' ('+((c.raw/t)*100).toFixed(1)+'%)'}}}},
      scales:tipo==='bar'?{x:{beginAtZero:true,ticks:{stepSize:1}},y:{beginAtZero:true,ticks:{stepSize:1}}}:{}}
  });
}
function upDiag(l,v){
  if(charts.chartDiag)charts.chartDiag.destroy();
  charts.chartDiag=new Chart(document.getElementById('chartDiag').getContext('2d'),{
    type:'bar',data:{labels:l,datasets:[{label:'Freq. Absoluta',data:v,backgroundColor:'#0a6ebd'}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',
      plugins:{tooltip:{callbacks:{label:c=>{const t=c.dataset.data.reduce((a,b)=>a+b,0);return c.raw+' ('+((c.raw/t)*100).toFixed(1)+'%)'}}}},
      scales:{x:{beginAtZero:true,ticks:{stepSize:1}}}}
  });
}

function exportChart(id,fn){
  const c=charts[id];if(!c){alert('Sem gráfico');return}
  const cv=c.canvas,tc=document.createElement('canvas');tc.width=cv.width;tc.height=cv.height;
  const x=tc.getContext('2d');x.fillStyle='#FFF';x.fillRect(0,0,tc.width,tc.height);x.drawImage(cv,0,0);
  const a=document.createElement('a');a.href=tc.toDataURL('image/png',1);a.download=fn+'_'+new Date().toISOString().slice(0,10)+'.png';a.click();
}
function exportAllCharts(){
  [['chartSexo','Sexo'],['chartDestino','Destino'],['chartProv','Proveniencia'],['chartFaixa','Faixa'],['chartDiag','Diagnostico']].forEach(([id,n],i)=>{
    if(charts[id])setTimeout(()=>exportChart(id,n),i*300);
  });
}

/* ===== PARTILHA ===== */
function openShareModal(){
  if(!registos.length){alert('Adicione registos primeiro.');return}
  document.getElementById('codeResult').style.display='none';
  document.getElementById('hostedResult').style.display='none';
  document.getElementById('shareModal').classList.add('active');
}
function closeShareModal(){document.getElementById('shareModal').classList.remove('active')}
document.getElementById('shareModal').addEventListener('click',e=>{if(e.target.id==='shareModal')closeShareModal()});

// MÉTODO 2: Código de partilha
function generateShareCode(){
  const json=JSON.stringify(registos);
  const code=LZString.compressToEncodedURIComponent(json);
  document.getElementById('shareCodeBox').textContent=code;
  document.getElementById('codeSize').textContent=(code.length/1024).toFixed(1)+' KB';
  document.getElementById('codeRegs').textContent=registos.length;
  document.getElementById('codeResult').style.display='block';
}
function copyCode(){
  const t=document.getElementById('shareCodeBox').textContent;
  navigator.clipboard.writeText(t).then(()=>showToast('✅ Código copiado!')).catch(()=>{
    const ta=document.createElement('textarea');ta.value=t;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);
    showToast('✅ Código copiado!');
  });
}

// IMPORTAR código
function importFromCode(){
  const code=document.getElementById('importCode').value.trim();
  if(!code){alert('Cole um código primeiro.');return}
  try{
    const json=LZString.decompressFromEncodedURIComponent(code);
    if(!json)throw new Error('Código inválido');
    const d=JSON.parse(json);
    if(!Array.isArray(d))throw new Error('Formato inválido');
    if(confirm('Importar '+d.length+' registos? (substitui os atuais)')){
      registos=d;salvar();gerarID();render();
      document.getElementById('importBar').classList.remove('active');
      document.getElementById('importCode').value='';
      showToast('✅ '+d.length+' registos importados!');
    }
  }catch(e){alert('❌ Erro: código inválido ou corrompido.\n\n'+e.message)}
}

// MÉTODO 3: Link com hospedagem
function generateHostedLink(){
  const url=document.getElementById('hostedUrl').value.trim();
  if(!url){alert('Cole o URL do dashboard hospedado.');return}
  const json=JSON.stringify(registos);
  const compressed=LZString.compressToEncodedURIComponent(json);
  const finalUrl=url+(url.includes('#')?'&':'#')+'d='+compressed;
  document.getElementById('finalLink').value=finalUrl;
  document.getElementById('hostedResult').style.display='block';
}
function copyFinalLink(){
  const v=document.getElementById('finalLink').value;
  navigator.clipboard.writeText(v).then(()=>showToast('✅ Link copiado!')).catch(()=>{
    const ta=document.createElement('textarea');ta.value=v;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);
    showToast('✅ Link copiado!');
  });
}

// MÉTODO 1: HTML autónomo
function exportStandaloneHTML(){
  if(!registos.length){alert('Sem registos.');return}
  const dadosJSON=JSON.stringify(registos).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
  const htmlContent=`<!DOCTYPE html>
<html lang="pt"><head><meta charset="UTF-8"><title>Dashboard Clínico (Partilhado)</title>
<script>
// DADOS IMPORTADOS AUTOMATICAMENTE - ${new Date().toISOString()}
window.__IMPORTED_DATA__ = JSON.parse('${dadosJSON}');
<\/script>
<script src="https://cdn.jsdelivr.net/npm/chart.js"><\/script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/lz-string@1.5.0/libs/lz-string.min.js"><\/script>
</head><body>
<p style="text-align:center;padding:20px;font-family:sans-serif;font-size:18px;color:#0a6ebd">
⏳ A carregar dashboard com ${registos.length} registos partilhados...</p>
<script>
// Este script redireciona para o dashboard original com os dados no hash
const compressed = LZString.compressToEncodedURIComponent(JSON.stringify(window.__IMPORTED_DATA__));
// Tenta carregar o dashboard do mesmo diretório
const basePath = window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/') + 1);
// Se este ficheiro for o próprio dashboard, usa os dados diretamente
if(typeof render === 'function') {
  registos = window.__IMPORTED_DATA__;
  localStorage.setItem('registosClinicos', JSON.stringify(registos));
  render();
} else {
  // Senão, injecta os dados no localStorage e recarrega
  localStorage.setItem('registosClinicos', JSON.stringify(window.__IMPORTED_DATA__));
  document.body.innerHTML = '<p style="text-align:center;padding:40px;font-family:sans-serif;font-size:16px;color:#27ae60">✅ ${registos.length} registos importados! O dashboard está pronto.<br><br>Recarregue a página ou abra o ficheiro dashboard original.</p>';
}
<\/script></body></html>`;
  
  // Abordagem mais simples: copiar o HTML atual e injetar dados
  const currentHTML = document.documentElement.outerHTML;
  const injected = currentHTML.replace(
    "let registos=JSON.parse(localStorage.getItem('registosClinicos')||'[]');",
    "let registos=(function(){try{const d=JSON.parse('"+dadosJSON+"');if(d.length>0&&!localStorage.getItem('registosClinicos')){localStorage.setItem('registosClinicos',JSON.stringify(d));return d}return JSON.parse(localStorage.getItem('registosClinicos')||'[]')}catch(e){return JSON.parse(localStorage.getItem('registosClinicos')||'[]')}})();"
  );
  
  const blob=new Blob([injected],{type:'text/html;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='Dashboard_Clinico_'+new Date().toISOString().slice(0,10)+'.html';
  a.click();URL.revokeObjectURL(a.href);
  showToast('✅ HTML descarregado!');
}

function showToast(msg){
  const t=document.getElementById('toast');t.textContent=msg;
  t.classList.add('show');setTimeout(()=>t.classList.remove('show'),3000);
}

/* ===== EXCEL ===== */
function exportExcel(){
  if(typeof XLSX==='undefined'){alert('Excel não carregado');return}
  if(!registos.length){alert('Sem registos');return}
  const dados=getFiltrados(),t=dados.length,wb=XLSX.utils.book_new();
  const fd=[['FICHAS DE REGISTO CLÍNICO'],['Exportado:',new Date().toLocaleString('pt-PT')],['Total:',t],[],
    ['#','ID','Entrada','Sexo','Idade','Un.','Faixa','Proveniência','Diagnóstico','Destino','Saída','Médico','Obs.']];
  dados.forEach((r,i)=>fd.push([i+1,r.id,fmtD(r.dataEntrada),r.sexo,r.idade,r.unidade,r.faixaEtaria,r.proveniencia,r.diagnostico,r.destino,fmtD(r.dataSaida),r.medico,r.observacoes||'']));
  const ws=XLSX.utils.aoa_to_sheet(fd);
  ws['!cols']=[{wch:5},{wch:15},{wch:18},{wch:12},{wch:8},{wch:8},{wch:14},{wch:28},{wch:30},{wch:18},{wch:18},{wch:20},{wch:30}];
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:12}}];
  XLSX.utils.book_append_sheet(wb,ws,'Registos');
  const pct=v=>t?((v/t)*100).toFixed(1)+'%':'0%';
  const sd=[['ESTATÍSTICAS'],['Data:',new Date().toLocaleString('pt-PT')],['Total:',t],[]];
  [['SEXO','sexo'],['DESTINO','destino'],['PROVENIÊNCIA','proveniencia'],['FAIXA ETÁRIA','faixaEtaria'],['DIAGNÓSTICO','diagnostico']].forEach(([titulo,campo])=>{
    const c=contar(dados,campo);
    sd.push([titulo],[campo,'Freq. Abs.','Freq. %']);
    Object.keys(c).forEach(k=>sd.push([k,c[k],pct(c[k])]));
    sd.push([]);
  });
  const ws2=XLSX.utils.aoa_to_sheet(sd);ws2['!cols']=[{wch:30},{wch:22},{wch:18}];
  XLSX.utils.book_append_sheet(wb,ws2,'Estatísticas');
  XLSX.writeFile(wb,'Relatorio_'+new Date().toISOString().slice(0,10)+'.xlsx');
}

/* ===== WORD ===== */
function exportWord(){
  const dados=getFiltrados(),t=dados.length;
  const sexo=contar(dados,'sexo'),dest=contar(dados,'destino'),prov=contar(dados,'proveniencia'),
  faixa=contar(dados,'faixaEtaria'),diag=contar(dados,'diagnostico');
  const pct=v=>t?((v/t)*100).toFixed(1):0;
  function tDist(o,tit){
    let h='<h3 style="color:#0a6ebd">'+tit+'</h3><table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse;width:60%"><tr style="background:#0a6ebd;color:#fff"><th>Categoria</th><th>Freq. Abs.</th><th>Freq. %</th></tr>';
    Object.keys(o).forEach(k=>{h+='<tr><td>'+k+'</td><td style="text-align:center">'+o[k]+'</td><td style="text-align:center">'+pct(o[k])+'%</td></tr>'});
    return h+'</table><br>';
  }
  function imgB(id,tit){
    const c=charts[id];if(!c)return'<p><i>Gráfico não disponível</i></p>';
    const cv=c.canvas,tc=document.createElement('canvas');tc.width=cv.width;tc.height=cv.height;
    const x=tc.getContext('2d');x.fillStyle='#FFF';x.fillRect(0,0,tc.width,tc.height);x.drawImage(cv,0,0);
    return'<p><img src="'+tc.toDataURL('image/png')+'" style="max-width:500px;width:100%"/></p><br>';
  }
  let tab='<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:11px"><tr style="background:#0a6ebd;color:#fff"><th>#</th><th>ID</th><th>Entrada</th><th>Sexo</th><th>Idade</th><th>Faixa</th><th>Prov.</th><th>Diag.</th><th>Destino</th><th>Saída</th><th>Médico</th><th>Obs.</th></tr>';
  dados.forEach((r,i)=>{tab+='<tr><td>'+(i+1)+'</td><td>'+r.id+'</td><td>'+fmtD(r.dataEntrada)+'</td><td>'+r.sexo+'</td><td>'+r.idade+' '+r.unidade+'</td><td>'+r.faixaEtaria+'</td><td>'+r.proveniencia+'</td><td>'+r.diagnostico+'</td><td>'+r.destino+'</td><td>'+fmtD(r.dataSaida)+'</td><td>'+r.medico+'</td><td>'+(r.observacoes||'')+'</td></tr>'});
  tab+='</table>';
  const html='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"></head><body style="font-family:Calibri"><h1 style="color:#0a6ebd;text-align:center">RELATÓRIO CLÍNICO</h1><p style="text-align:center">'+new Date().toLocaleString('pt-PT')+'</p><h2>Resumo: '+t+' doentes</h2>'+tDist(sexo,'Sexo')+imgB('chartSexo','Sexo')+tDist(dest,'Destino')+imgB('chartDestino','Destino')+tDist(prov,'Proveniência')+imgB('chartProv','Prov')+tDist(faixa,'Faixa Etária')+imgB('chartFaixa','Faixa')+tDist(diag,'Diagnóstico')+imgB('chartDiag','Diag')+'<h2>Registos</h2>'+tab+'</body></html>';
  const b=new Blob(['\ufeff'+html],{type:'application/msword'});
  const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='Relatorio_'+new Date().toISOString().slice(0,10)+'.doc';a.click();
}

function exportJSON(){
  const b=new Blob([JSON.stringify(registos,null,2)],{type:'application/json'});
  const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='registos.json';a.click();
}
function fmtD(d){return d?d.replace('T',' ').slice(0,16):''}
</script>
</body>
</html>
