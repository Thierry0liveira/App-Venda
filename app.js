// GESTÃO COMERCIAL (PWA) - como eu projetei este aplicativo
// - É um app-base que eu entrego a cada comerciante. Tudo roda no aparelho dele: sem servidor, sem login, sem mensalidade.
// - Os dados ficam no IndexedDB. Eu NÃO uso o navegador como "banco seguro": por isso existe o backup.
// - Fluxo: boas-vindas -> assistente de configuração (4 etapas) -> painel. Depois, Configurações edita tudo.
// - Dia a dia: Nova comanda -> Histórico. Para organização eu escolhi uma PLANILHA (.xlsx) de consulta,
//   em vez de dashboards: o comerciante já sabe filtrar e somar no Excel.
// - Planilha e backup são coisas diferentes de propósito: a planilha é para LER (o dono pode mexer nela),
//   o backup (.json) é para RESTAURAR (formato fixo, guarda também logo, cores e catálogo).
// - Salvar no Google Drive: no celular uso a folha de compartilhar do aparelho; no computador o arquivo baixa
//   e o dono guarda na pasta do Drive. Não integro API do Google para não exigir login nem cadastro.
'use strict';
// Atalho para document.querySelector.
const $=s=>document.querySelector(s);
// Escapa HTML: impede que textos digitados (nome, categoria...) virem código na tela (XSS).
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
// Formata centavos como moeda brasileira (2550 -> R$ 25,50).
const brl=c=>(c/100).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
// Tipos sugeridos (apenas referência: 'Outro' aceita texto livre) e títulos das etapas do assistente.
// Formas de pagamento da comanda; total da comanda aberta (cart = {idDoItem: quantidade}); data legível.
const PAGS=['Dinheiro','Pix','Cartão de débito','Cartão de crédito'];
const totalCart=()=>S.items.reduce((t,i)=>t+(S.cart[i.id]||0)*i.preco,0);
const dataBR=iso=>new Date(iso).toLocaleString('pt-BR',{dateStyle:'short',timeStyle:'short'});
// ---- Backup e planilha: funções de apoio ----
// Aviso de backup: calculo há quantos dias foi o último; fica em destaque a partir de 7 dias (só se já houver vendas).
const bkDias=()=>S.bk?Math.floor((Date.now()-new Date(S.bk))/864e5):null;
const bkTxt=()=>{const d=bkDias();return d===null?'Nenhum backup feito ainda':d===0?'Último backup: hoje':`Último backup: há ${d} ${d===1?'dia':'dias'}`};
const bkOld=()=>S.vendas.length>0&&(bkDias()===null||bkDias()>=7);
// Confirmo que o arquivo é mesmo um backup deste app antes de deixar ele substituir qualquer coisa.
const okBackup=d=>!!d&&d.app==='gestao-comercial'&&!!d.config&&typeof d.config.nome==='string'&&Array.isArray(d.catalogo)&&Array.isArray(d.vendas);
// Filtro do período da planilha (usa a data local do aparelho): hoje, este mês ou tudo.
const ymd=d=>d.getFullYear()*10000+(d.getMonth()+1)*100+d.getDate();
const filtrarPeriodo=p=>{const h=new Date();return S.vendas.filter(v=>{const d=new Date(v.data);return p==='tudo'||(p==='hoje'?ymd(d)===ymd(h):d.getFullYear()===h.getFullYear()&&d.getMonth()===h.getMonth())})};
// Nome de arquivo seguro: sem acentos, espaços ou símbolos.
const slug=t=>t.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/gi,'-').replace(/^-|-$/g,'').toLowerCase()||'comercio';
const lerTexto=f=>new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=()=>no(r.error);r.readAsText(f)});
// Entrega do arquivo: no celular abro a folha de compartilhar (o dono escolhe "Salvar no Drive");
// no computador baixo direto. Devolve false se o usuário cancelou o compartilhamento.
async function entregar(blob,nome){const f=new File([blob],nome,{type:blob.type});
 if(window.matchMedia&&matchMedia('(pointer:coarse)').matches&&navigator.canShare?.({files:[f]})){
  try{await navigator.share({files:[f]});return true}catch(e){if(e.name==='AbortError')return false}} // outro erro: cai no download
 const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=nome;document.body.appendChild(a);a.click();a.remove();
 setTimeout(()=>URL.revokeObjectURL(a.href),4e3);return true}
// Texto do período no Resumo: "Novembro de 2026" (mês), a data de hoje, ou "Todo o histórico".
const nomePeriodo=p=>{const h=new Date();if(p==='hoje')return h.toLocaleDateString('pt-BR');if(p==='tudo')return'Todo o histórico';
 const m=h.toLocaleDateString('pt-BR',{month:'long',year:'numeric'});return m[0].toUpperCase()+m.slice(1)};
// Monta o .xlsx. A data vai como número serial do Excel com formato dd/mm/aaaa (e não como data nativa),
// porque o Google Planilhas não lê o tipo de data nativo e deixava a coluna em branco. Assim funciona em Excel, Planilhas e LibreOffice.
// Aba "Vendas": uma linha por item vendido (fácil de filtrar/somar). Aba "Resumo": totais com FÓRMULAS
// ligadas à aba Vendas, então se o dono corrigir um valor na planilha, o resumo recalcula sozinho.
function criarPlanilha(vs,per){const X=window.XLSX;if(!X)throw new Error('Biblioteca xlsx ausente');
 const R=c=>c/100,L=[['Comanda','Data','Hora','Item','Quantidade','Preço unitário','Subtotal','Pagamento']];
 vs.slice().reverse().forEach(v=>{const d=new Date(v.data);v.itens.forEach(i=>{const r=L.length+1;
  L.push([v.numero,{t:'n',v:(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())-Date.UTC(1899,11,30))/864e5,z:'dd/mm/yyyy'},d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}),i.nome,i.qtd,R(i.preco),{t:'n',f:`E${r}*F${r}`,v:R(i.qtd*i.preco)},v.pagamento])})});
 const RS='"R$" #,##0.00',ws=X.utils.aoa_to_sheet(L);
 for(let r=2;r<=L.length;r++){ws['B'+r].z='dd/mm/yyyy';ws['F'+r].z=RS;ws['G'+r].z=RS}
 ws['!cols']=[9,12,7,30,11,14,12,18].map(wch=>({wch}));
 const fim=Math.max(L.length,2),rg=c=>`Vendas!${c}2:${c}${fim}`,tot=vs.reduce((t,v)=>t+v.total,0);
 const res=[['Estabelecimento',S.cfg.nome],['Período',nomePeriodo(per)],['Gerado em',new Date().toLocaleString('pt-BR')],[],
  ['Total vendido',{t:'n',f:`SUM(${rg('G')})`,v:R(tot)}],['Comandas',vs.length],['Ticket médio',{t:'n',f:'IF(B6=0,0,B5/B6)',v:R(tot)/vs.length}],[],['Por forma de pagamento'],
  ...PAGS.map(p=>[p,{t:'n',f:`SUMIF(${rg('H')},"${p}",${rg('G')})`,v:R(vs.filter(v=>v.pagamento===p).reduce((t,v)=>t+v.total,0))}])];
 const wr=X.utils.aoa_to_sheet(res);['B5','B7',...PAGS.map((_,i)=>'B'+(10+i))].forEach(c=>wr[c].z=RS);wr['!cols']=[{wch:26},{wch:26}];
 const wb=X.utils.book_new();X.utils.book_append_sheet(wb,wr,'Resumo');X.utils.book_append_sheet(wb,ws,'Vendas');
 return X.write(wb,{bookType:'xlsx',type:'array'})}
// Restaurar: valida, pede confirmação (substitui TUDO) e só então grava. Depois recarrego o estado e reaplico o visual.
async function restaurar(f){let d;try{d=JSON.parse(await lerTexto(f))}catch{d=null}
 if(!okBackup(d))return toast('Este arquivo não é um backup válido.');
 if(!confirm(`Restaurar o backup de ${dataBR(d.exportadoEm)}? Isso SUBSTITUI todos os dados atuais deste aparelho.`))return;
 try{d.config.concluido=true;await DB.restore(d);S.cfg={...DEF,...d.config};S.d={...S.cfg};theme(S.cfg);pwa(S.cfg);await loadItems();await loadVendas();S.view='dash';toast('Backup restaurado')}
 catch{toast('Não foi possível restaurar. Nada foi alterado.')}render()}
const TIPOS=['Chaveiro','Barbearia','Restaurante','Outro'];
const STEPS=['Negócio','Identidade','Catálogo','Conclusão'];
// Configuração padrão: só as cores iniciais. Nenhum dado de comércio fica fixo no código.
const DEF={nome:'',tipo:'',tipoOutro:'',dono:'',info:'',logo:null,c1:'#1f6f5c',c2:'#e8a33d',concluido:false};
// Estado da aplicação:
//  cfg = configuração salva no banco; d = cópia em edição (rascunho);
//  items = catálogo; view = tela atual; step/tab = posição no assistente/configurações;
//  form = formulário de item aberto; errs = mensagens de validação.
const S={cfg:{...DEF},d:{...DEF},items:[],vendas:[],cart:{},pag:'',per:'mes',bk:null,view:'boot',step:0,tab:'dados',form:null,errs:{}};

/* ---------- IndexedDB ---------- */
// Camada mínima sobre o IndexedDB, com dois 'stores':
//  config (chave -> valor) e catalogo (itens com id automático). Tudo retorna Promise.
const DB={db:null,
 open(){return new Promise((ok,no)=>{const r=indexedDB.open('gestao-comercial',2);
  r.onupgradeneeded=()=>{const d=r.result; // versão 2 acrescenta 'vendas'; bancos antigos só ganham o que falta
  if(!d.objectStoreNames.contains('config'))d.createObjectStore('config');
  if(!d.objectStoreNames.contains('catalogo'))d.createObjectStore('catalogo',{keyPath:'id',autoIncrement:true});
  if(!d.objectStoreNames.contains('vendas'))d.createObjectStore('vendas',{keyPath:'numero',autoIncrement:true})}; // 'numero' = nº automático da comanda
  r.onsuccess=()=>{DB.db=r.result;ok()};r.onerror=()=>no(r.error)})},
 tx(st,mode,fn){return new Promise((ok,no)=>{const t=DB.db.transaction(st,mode),q=fn(t.objectStore(st));
  t.oncomplete=()=>ok(q&&q.result);t.onerror=t.onabort=()=>no(t.error)})},
 get:k=>DB.tx('config','readonly',s=>s.get(k)),
 set:(k,v)=>DB.tx('config','readwrite',s=>s.put(v,k)),
 all:()=>DB.tx('catalogo','readonly',s=>s.getAll()),
 put:i=>DB.tx('catalogo','readwrite',s=>s.put(i)),
 del:id=>DB.tx('catalogo','readwrite',s=>s.delete(id)),
 vall:()=>DB.tx('vendas','readonly',s=>s.getAll()),
 vadd:v=>DB.tx('vendas','readwrite',s=>s.add(v)),
 // Restaurar backup: limpo e regravo os três stores na MESMA transação. Se algo falhar, nada muda (tudo ou nada).
 restore(d){return new Promise((ok,no)=>{const t=DB.db.transaction(['config','catalogo','vendas'],'readwrite');
  t.objectStore('config').put(d.config,'estabelecimento');
  const c=t.objectStore('catalogo'),v=t.objectStore('vendas');c.clear();v.clear();
  d.catalogo.forEach(i=>c.put(i));d.vendas.forEach(x=>v.put(x)); // mantenho ids e números originais das comandas
  t.oncomplete=ok;t.onerror=t.onabort=()=>no(t.error)})}}; // vadd devolve o número gerado
// Recarrega o catálogo do banco em ordem alfabética.
// Carrega as comandas, da mais recente para a mais antiga (os números crescem em ordem).
const loadVendas=async()=>{S.vendas=(await DB.vall()).reverse()};
const loadItems=async()=>{S.items=(await DB.all()).sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'))};

/* ---------- Tema, ícones e manifest dinâmico ---------- */
// Escolhe texto branco ou escuro conforme a luminosidade da cor de fundo (garante legibilidade).
const on=h=>{const [r,g,b]=[1,3,5].map(i=>parseInt(h.substr(i,2),16));return(r*299+g*587+b*114)/1000>150?'#16161a':'#fff'};
// Aplica cores e título da aba por variáveis CSS (todo o visual usa --c1 e --c2).
function theme(c){const r=document.documentElement.style;r.setProperty('--c1',c.c1);r.setProperty('--c2',c.c2);
 r.setProperty('--on1',on(c.c1));r.setProperty('--on2',on(c.c2));$('meta[name=theme-color]').content=c.c1;document.title=c.nome||'Gestão Comercial'}
// Identidade do PWA: gera ícones PNG (192 e 512) a partir da logo ou da inicial do nome,
//  atualiza o favicon e cria um manifest dinâmico (Blob). Se falhar, o app continua funcionando.
async function pwa(c){try{
 const icon=size=>new Promise(ok=>{const k=document.createElement('canvas');k.width=k.height=size;const x=k.getContext('2d');
  const letra=()=>{x.fillStyle=c.c1;x.fillRect(0,0,size,size);x.fillStyle=on(c.c1);x.font=`800 ${size*.5}px system-ui,sans-serif`;x.textAlign='center';x.textBaseline='middle';
   x.fillText((c.nome||'G')[0].toUpperCase(),size/2,size*.53);ok(k.toDataURL('image/png'))};
  if(!c.logo)return letra();
  const im=new Image();im.onload=()=>{x.fillStyle='#fff';x.fillRect(0,0,size,size);const w=size*.64,r=Math.min(w/im.width,w/im.height);
   x.drawImage(im,(size-im.width*r)/2,(size-im.height*r)/2,im.width*r,im.height*r);ok(k.toDataURL('image/png'))};im.onerror=letra;im.src=c.logo});
 const [a,b]=await Promise.all([icon(192),icon(512)]);
 for(const id of['#fav','#touch']){$(id).type='image/png';$(id).href=a}
 const base=new URL('./',location.href).href;
 const m={name:c.nome||'Gestão Comercial',short_name:(c.nome||'Gestão').slice(0,12),start_url:base,scope:base,display:'standalone',background_color:'#ffffff',theme_color:c.c1,
  icons:[{src:a,sizes:'192x192',type:'image/png',purpose:'any maskable'},{src:b,sizes:'512x512',type:'image/png',purpose:'any maskable'}]};
 $('#mf').href=URL.createObjectURL(new Blob([JSON.stringify(m)],{type:'application/manifest+json'}));
}catch(e){console.warn('Ícones PWA não atualizados',e)}}

/* ---------- Utilitários ---------- */
// Mensagem temporária (toast) de feedback ao usuário.
let tt;const toast=m=>{const t=$('#toast');t.textContent=m;t.classList.add('on');clearTimeout(tt);tt=setTimeout(()=>t.classList.remove('on'),3200)};
// Tipo exibido: usa o texto livre quando o tipo escolhido é 'Outro'.
const tipoLabel=c=>c.tipo==='Outro'?(c.tipoOutro.trim()||'Outro'):c.tipo;
// Converte '1.234,50' ou '25.9' em centavos (inteiro). Devolve null se for inválido ou negativo.
function parsePreco(s){s=String(s).replace(/[R$\s]/g,'');if(s.includes(','))s=s.replace(/\./g,'').replace(',','.');const n=Number(s);return s&&isFinite(n)&&n>=0&&n<1e7?Math.round(n*100):null}
// Lê a imagem escolhida, valida tipo e tamanho e reduz para no máx. 512 px (PNG em base64) antes de guardar.
function readLogo(f){return new Promise((ok,no)=>{
 if(!f.type.startsWith('image/'))return no(new Error('Escolha um arquivo de imagem.'));
 if(f.size>8e6)return no(new Error('Imagem muito grande. Use um arquivo de até 8 MB.'));
 const u=URL.createObjectURL(f),im=new Image();
 im.onload=()=>{const r=Math.min(1,512/Math.max(im.width,im.height)),c=document.createElement('canvas');c.width=Math.round(im.width*r)||1;c.height=Math.round(im.height*r)||1;
  c.getContext('2d').drawImage(im,0,0,c.width,c.height);URL.revokeObjectURL(u);ok(c.toDataURL('image/png'))};
 im.onerror=()=>{URL.revokeObjectURL(u);no(new Error('Não foi possível ler esta imagem. Tente outro arquivo.'))};im.src=u})}
// Grava a configuração no IndexedDB e só então atualiza estado, tema e ícones.
//  done=true marca a configuração inicial como concluída (é o que evita repeti-la).
async function saveCfg(done){const c={...S.d,nome:S.d.nome.trim(),concluido:done||S.cfg.concluido};
 try{await DB.set('estabelecimento',c);S.cfg=c;S.d={...c};theme(c);pwa(c);return true}
 catch(e){toast('Não foi possível salvar. Verifique o espaço e as permissões do navegador.');return false}}
// Valida os campos obrigatórios da etapa 1: nome e tipo.
function validar(){S.errs={};if(!S.d.nome.trim())S.errs.nome='Informe o nome do estabelecimento.';if(!S.d.tipo)S.errs.tipo='Escolha um tipo (ou "Outro").';return!Object.keys(S.errs).length}

/* ---------- Componentes ---------- */
// Componentes de interface (cada um devolve HTML):
//  err = mensagem de erro; brand = logo ou inicial; head = cabeçalho; chips = botões de opção.
const err=k=>S.errs[k]?`<small class=err>${S.errs[k]}</small>`:'';
const brand=c=>c.logo?`<img class=logo src="${c.logo}" alt="Logotipo">`:`<span class=logo>${esc((c.nome||'G')[0].toUpperCase())}</span>`;
const head=r=>`<header class=top><div class=bd>${brand(S.d)}<b>${esc(S.d.nome||'Gestão Comercial')}</b></div>${r||''}</header>`;
const chips=(items,cur,act)=>items.map(([v,l])=>`<button type=button class="chip ${cur===v?'on':''}" data-a=${act} data-v="${esc(v)}">${l}</button>`).join('');

// Formulário 'Negócio'. Reaproveitado no assistente e nas Configurações.
const fDados=()=>{const d=S.d;return`<div class=card>
<label class=fld><span>Nome do estabelecimento *</span><input data-f=nome value="${esc(d.nome)}" maxlength=60 autocomplete=off>${err('nome')}</label>
<div class=fld><span>Tipo de estabelecimento *</span><div class=chips>${chips(TIPOS.map(t=>[t,t]),d.tipo,'tipo')}</div>${err('tipo')}</div>
${d.tipo==='Outro'?`<label class=fld><span>Qual é o seu tipo de negócio?</span><input data-f=tipoOutro value="${esc(d.tipoOutro)}" maxlength=40></label>`:''}
<label class=fld><span>Nome do proprietário <em>(opcional)</em></span><input data-f=dono value="${esc(d.dono)}" maxlength=60></label>
<label class=fld><span>Informações adicionais <em>(opcional)</em></span><textarea data-f=info rows=3 maxlength=300 placeholder="Endereço, telefone, horário de funcionamento">${esc(d.info)}</textarea></label></div>`};

// Formulário 'Identidade': logotipo, cores e prévia ao vivo.
const fVisual=()=>{const d=S.d;return`<div class=card>
<div class=fld><span>Logotipo</span><div class=logoPick>${brand(d)}<div>
<label class="btn ghost sm" style="color:var(--ink)">${d.logo?'Trocar imagem':'Escolher imagem'}<input type=file accept="image/*" data-a=file hidden></label>
${d.logo?'<button class="btn link" data-a=rmlogo>Remover</button>':''}<small>A imagem é ajustada automaticamente.</small>${err('logo')}</div></div></div>
<div class=cols><label class=fld><span>Cor principal</span><input type=color data-f=c1 value="${d.c1}"></label>
<label class=fld><span>Cor secundária</span><input type=color data-f=c2 value="${d.c2}"></label></div>
<div class=prev><div class=pv-h>${brand(d)}<b class=pvn>${esc(d.nome||'Seu estabelecimento')}</b></div><div class=pv-b><span class=btn>Botão principal</span><span class="btn alt">Destaque</span></div></div></div>`};

// Catálogo: categorias já usadas (sugeridas ao digitar) e formulário + lista de itens (criar, editar, excluir).
const cats=()=>[...new Set(S.items.map(i=>i.categoria).filter(Boolean))];
const fCat=()=>{const f=S.form||(S.form={nome:'',descricao:'',categoria:'',preco:'',tipo:'produto',ativo:true});
return`<div class=card><h3>${f.id?'Editar item':'Novo item'}</h3>
<label class=fld><span>Nome *</span><input data-ff=nome value="${esc(f.nome)}" maxlength=80>${err('nome')}</label>
<label class=fld><span>Descrição <em>(opcional)</em></span><input data-ff=descricao value="${esc(f.descricao)}" maxlength=160></label>
<div class=cols><label class=fld><span>Categoria</span><input data-ff=categoria list=cl value="${esc(f.categoria)}" maxlength=40><datalist id=cl>${cats().map(c=>`<option value="${esc(c)}">`).join('')}</datalist></label>
<label class=fld><span>Preço (R$) *</span><input data-ff=preco inputmode=decimal value="${esc(f.preco)}" placeholder="0,00">${err('preco')}</label></div>
<div class=cols><div class=fld><span>Tipo</span><div class=chips>${chips([['produto','Produto'],['serviço','Serviço']],f.tipo,'ftipo')}</div></div>
<label class=fld><span>Ativo</span><input type=checkbox data-ffc=ativo ${f.ativo?'checked':''}></label></div>
<div class=row><button class=btn data-a=saveitem>${f.id?'Salvar item':'Adicionar item'}</button>${f.id?'<button class="btn ghost" style="color:var(--ink)" data-a=cancelitem>Cancelar</button>':''}</div></div>
<div class=card><h3>Catálogo <em>(${S.items.length})</em></h3>${S.items.length?`<ul class=list>${S.items.map(i=>`<li class="${i.ativo?'':'off'}"><div><b>${esc(i.nome)}</b><small>${i.tipo==='serviço'?'Serviço':'Produto'}${i.categoria?', '+esc(i.categoria):''}${i.ativo?'':', inativo'}</small></div><strong>${brl(i.preco)}</strong><button class="btn link" data-a=edititem data-id=${i.id}>Editar</button><button class="btn link dng" data-a=delitem data-id=${i.id}>Excluir</button></li>`).join('')}</ul>`:'<p class=empty>Nenhum item ainda. Adicione o primeiro acima.</p>'}</div>`};

// Última etapa do assistente: nome, logotipo, tipo e quantidade de produtos/serviços.
const resumo=()=>{const p=S.items.filter(i=>i.tipo!=='serviço').length;return`<div class="card center"><h2>Tudo pronto</h2>${brand(S.d)}<h3>${esc(S.d.nome)}</h3><p class=sub>${esc(tipoLabel(S.d))}</p>
<div class=stats><div><b>${S.items.length}</b>no catálogo</div><div><b>${p}</b>produtos</div><div><b>${S.items.length-p}</b>serviços</div></div></div>`};

/* ---------- Telas ---------- */
// Telas: cada uma é uma função que devolve o HTML. render() escolhe a tela conforme S.view.
const V={
 boot:()=>'',
 fail:()=>`<main class=hero><div class=hero-in><h1>Armazenamento indisponível</h1><p>Não foi possível abrir o armazenamento local do navegador. Saia do modo privado, libere espaço e recarregue a página.</p></div></main>`,
 welcome:()=>`<main class="hero view"><div class=hero-in><h1>O seu comércio, do seu jeito.</h1><p>Configure o aplicativo com o nome, a marca e o catálogo do seu estabelecimento. Leva poucos minutos e não precisa de conta.</p>
<button class="btn lg" data-a=start>Configurar meu estabelecimento</button><ul class=pts><li>Funciona offline</li><li>Dados ficam no seu aparelho</li><li>Sem login</li></ul></div></main>`,
 setup:()=>`${head()}<main class="wrap view"><ol class=steps>${STEPS.map((s,i)=>`<li class="${i===S.step?'on':i<S.step?'done':''}">${s}</li>`).join('')}</ol>
${[()=>'<h2>Sobre o seu negócio</h2>'+fDados(),()=>'<h2>Identidade visual</h2>'+fVisual(),()=>'<h2>Produtos e serviços</h2><p class=sub>Cadastre o que você vende. Dá para alterar depois em Configurações.</p>'+fCat(),resumo][S.step]()}
<nav class=nav>${S.step?'<button class="btn ghost" data-a=back>Voltar</button>':'<span></span>'}${S.step<3?'<button class=btn data-a=next>Continuar</button>':'<button class=btn data-a=finish>Começar a utilizar</button>'}</nav></main>`,
 dash:()=>`${head('<button class="btn ghost sm" data-a=settings>Configurações</button>')}<main class="wrap view"><div class="card center"><h2>${S.d.dono?'Olá, '+esc(S.d.dono):'Tudo pronto'}</h2>
<p class=sub>${S.vendas.length} ${S.vendas.length===1?'comanda registrada':'comandas registradas'}.</p>
<div class=row style="justify-content:center"><button class=btn data-a=newcmd>Nova comanda</button><button class="btn ghost" style="color:var(--ink)" data-a=hist>Histórico</button><button class="btn ghost" style="color:var(--ink)" data-a=dados>Planilha e backup</button></div></div>
<p class="note ${bkOld()?'warn':''}">${bkTxt()}${bkOld()?'. Faça um backup para não perder suas vendas.':''}</p></main>`,
 // Nova comanda: só itens ativos; +/- ajusta a quantidade; total e pagamento no cartão abaixo.
 comanda:()=>{const at=S.items.filter(i=>i.ativo);return`${head('<button class="btn ghost sm" data-a=exit>Cancelar</button>')}<main class="wrap view"><h2>Nova comanda</h2>${at.length?`<div class=card><ul class=list>${at.map(i=>`<li class=prd><div><b>${esc(i.nome)}</b><small>${i.categoria?esc(i.categoria)+', ':''}${brl(i.preco)}</small></div><div class=qty><button class=chip data-a=dec data-id=${i.id} aria-label="Diminuir">−</button><b>${S.cart[i.id]||0}</b><button class=chip data-a=inc data-id=${i.id} aria-label="Aumentar">+</button></div></li>`).join('')}</ul></div>
<div class=card><div class=fld><span>Forma de pagamento</span><div class=chips>${chips(PAGS.map(p=>[p,p]),S.pag,'pag')}</div></div><div class=total><span>Total</span><strong>${brl(totalCart())}</strong></div><button class=btn style="width:100%" data-a=savecmd>Finalizar comanda</button></div>`:'<div class="card center"><p class=empty>Não há itens ativos no catálogo.</p><button class=btn data-a=settings>Cadastrar itens</button></div>'}</main>`},
 // Histórico: cada comanda abre (details) mostrando os itens como foram vendidos.
 hist:()=>`${head('<button class="btn ghost sm" data-a=exit>Fechar</button>')}<main class="wrap view"><h2>Histórico de vendas</h2>${S.vendas.length?S.vendas.map(v=>`<details class=card><summary class=total><span><b>Comanda ${v.numero}</b><small class=blk>${dataBR(v.data)}, ${esc(v.pagamento)}</small></span><strong>${brl(v.total)}</strong></summary><ul class=list>${v.itens.map(i=>`<li class=prd><div>${i.qtd} x ${esc(i.nome)}</div><span>${brl(i.qtd*i.preco)}</span></li>`).join('')}</ul></details>`).join(''):'<p class=empty>Nenhuma comanda registrada ainda.</p>'}</main>`,
 // Planilha e backup: tudo que o dono precisa para organizar e proteger os dados numa tela só, sem relatórios complicados.
 dados:()=>`${head('<button class="btn ghost sm" data-a=exit>Fechar</button>')}<main class="wrap view"><h2>Planilha e backup</h2>
<div class=card><h3>Planilha para consultar</h3><p class=sub>Gera um arquivo Excel com as vendas do período. Abra, filtre e some como quiser.</p><div class=chips>${chips([['hoje','Hoje'],['mes','Este mês'],['tudo','Tudo']],S.per,'per')}</div><button class=btn data-a=xlsx>Gerar planilha (.xlsx)</button></div>
<div class=card><h3>Backup</h3><p class=sub>Um arquivo com configurações, logo, catálogo e vendas. Guarde no Google Drive da empresa. Se os dados do navegador forem apagados, é ele que recupera tudo.</p>
<p class="note ${bkOld()?'warn':''}">${bkTxt()}</p><div class=row style="justify-content:center"><button class=btn data-a=savebk>Salvar backup</button><label class="btn ghost" style="color:var(--ink)">Restaurar backup<input type=file accept=".json,application/json" data-a=restore hidden></label></div></div></main>`,
 settings:()=>`${head('<button class="btn ghost sm" data-a=exit>Fechar</button>')}<main class="wrap view"><h2>Configurações</h2>
<div class=chips>${chips([['dados','Estabelecimento'],['visual','Visual'],['cat','Catálogo']],S.tab,'tab')}</div>
${S.tab==='dados'?fDados():S.tab==='visual'?fVisual():fCat()}${S.tab==='cat'?'':'<nav class=nav><span></span><button class=btn data-a=savecfg>Salvar alterações</button></nav>'}</main>`};
const render=()=>{$('#app').innerHTML=V[S.view]()};

/* ---------- Eventos ---------- */
// Eventos (delegação no document).
//  'input': guarda o que é digitado sem redesenhar a tela (para não perder o foco).
//  data-f = campo da configuração; data-ff = campo do item do catálogo.
document.addEventListener('input',e=>{const t=e.target,f=t.dataset;
 if(f.f){S.d[f.f]=t.value;if(f.f==='c1'||f.f==='c2')theme(S.d);if(f.f==='nome'){const p=$('.pvn');if(p)p.textContent=t.value||'Seu estabelecimento';document.title=t.value||'Gestão Comercial'}}
 if(f.ff)S.form[f.ff]=t.value});
// 'change': checkbox 'Ativo' do item e escolha do arquivo da logo.
document.addEventListener('change',async e=>{const t=e.target;
 if(t.dataset.a==='restore'&&t.files[0]){await restaurar(t.files[0]);t.value='';return}
 if(t.dataset.ffc)S.form[t.dataset.ffc]=t.checked;
 if(t.dataset.a==='file'&&t.files[0]){S.errs={};try{S.d.logo=await readLogo(t.files[0])}catch(x){S.errs.logo=x.message}render()}});
// 'click': cada botão tem data-a (a ação). Depois de tratar a ação, a tela é redesenhada.
document.addEventListener('click',async e=>{const b=e.target.closest('button[data-a]');if(!b)return;
 const a=b.dataset.a,v=b.dataset.v,id=+b.dataset.id;S.errs={};
 if(a==='start'){S.d={...DEF};theme(S.d);S.view='setup';S.step=0}
 else if(a==='tipo')S.d.tipo=v;
 else if(a==='rmlogo')S.d.logo=null;
 else if(a==='back')S.step--;
 else if(a==='next'){if(S.step>0||validar())S.step++}
 else if(a==='finish'){if(!await saveCfg(true))return;S.view='dash'}
 else if(a==='settings'){S.d={...S.cfg};S.tab='dados';S.form=null;S.view='settings'}
 else if(a==='exit'){S.d={...S.cfg};theme(S.cfg);S.form=null;S.cart={};S.pag='';S.view='dash'}
 else if(a==='tab'){S.tab=v;S.form=null}
 else if(a==='savecfg'){if(!validar())return render();if(await saveCfg(false))toast('Alterações salvas')}
 else if(a==='newcmd'){S.cart={};S.pag='';S.view='comanda'}
 else if(a==='hist')S.view='hist';
 else if(a==='inc')S.cart[id]=(S.cart[id]||0)+1;
 else if(a==='dec')S.cart[id]=Math.max(0,(S.cart[id]||0)-1);
 else if(a==='pag')S.pag=v;
 else if(a==='savecmd'){ // grava a comanda guardando nome e preço de cada item (o histórico não muda se o catálogo mudar)
  const itens=S.items.filter(i=>S.cart[i.id]>0).map(i=>({id:i.id,nome:i.nome,preco:i.preco,qtd:S.cart[i.id]}));
  if(!itens.length)return toast('Adicione pelo menos um item.');if(!S.pag)return toast('Escolha a forma de pagamento.');
  try{const n=await DB.vadd({data:new Date().toISOString(),itens,total:totalCart(),pagamento:S.pag});await loadVendas();S.cart={};S.pag='';S.view='dash';toast('Comanda '+n+' registrada')}catch{return toast('Não foi possível registrar a comanda.')}}
 // Ações da planilha e do backup (a restauração fica no evento 'change', porque depende do arquivo escolhido).
 else if(a==='dados')S.view='dados';
 else if(a==='per')S.per=v;
 else if(a==='xlsx'){const vs=filtrarPeriodo(S.per);if(!vs.length)return toast('Nenhuma venda neste período.');
  try{await entregar(new Blob([criarPlanilha(vs,S.per)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`vendas-${slug(S.cfg.nome)}-${S.per}-${new Date().toISOString().slice(0,10)}.xlsx`)}
  catch(e){console.error(e);return toast('Não foi possível gerar a planilha.')}}
 else if(a==='savebk'){ // só marco "último backup" se o arquivo realmente foi entregue
  try{const d={app:'gestao-comercial',versao:2,exportadoEm:new Date().toISOString(),config:S.cfg,catalogo:S.items,vendas:S.vendas.slice().reverse()};
   if(await entregar(new Blob([JSON.stringify(d)],{type:'application/json'}),`backup-${slug(S.cfg.nome)}-${d.exportadoEm.slice(0,10)}.json`)){S.bk=d.exportadoEm;await DB.set('ultimoBackup',S.bk);toast('Backup salvo')}}
  catch(e){console.error(e);return toast('Não foi possível salvar o backup.')}}
 else if(a==='ftipo')S.form.tipo=v;
 else if(a==='cancelitem')S.form=null;
 else if(a==='edititem'){const i=S.items.find(x=>x.id===id);S.form={...i,preco:(i.preco/100).toFixed(2).replace('.',',')};scrollTo(0,0)}
 else if(a==='delitem'){if(!confirm('Excluir este item do catálogo?'))return;try{await DB.del(id);await loadItems()}catch{toast('Não foi possível excluir.')}}
 else if(a==='saveitem'){const f=S.form,n=f.nome.trim(),p=parsePreco(f.preco);
  if(!n)S.errs.nome='Informe o nome.';if(p===null)S.errs.preco='Informe um preço válido, como 25,90.';
  if(!Object.keys(S.errs).length)try{await DB.put({...(f.id?{id:f.id}:{}),nome:n,descricao:f.descricao.trim(),categoria:f.categoria.trim(),preco:p,tipo:f.tipo,ativo:f.ativo});
   await loadItems();S.form=null;toast(f.id?'Item atualizado':'Item adicionado')}catch{toast('Não foi possível salvar o item.')}}
 render()});

/* ---------- Inicialização ---------- */
// Inicialização: abre o banco, pede armazenamento persistente, carrega configuração e catálogo,
//  escolhe a tela (painel se já configurado; boas-vindas se não) e registra o service worker.
(async()=>{try{await DB.open();navigator.storage?.persist?.();
 const c=await DB.get('estabelecimento');await loadItems();await loadVendas();S.bk=(await DB.get('ultimoBackup'))||null;
 if(c){S.cfg={...DEF,...c};S.d={...S.cfg};theme(S.cfg);pwa(S.cfg)}
 S.view=c&&c.concluido?'dash':'welcome'}catch(e){console.error(e);S.view='fail'}
 render();
 if('serviceWorker'in navigator&&(location.protocol==='https:'||location.hostname==='localhost'))navigator.serviceWorker.register('sw.js').catch(()=>{})})();
