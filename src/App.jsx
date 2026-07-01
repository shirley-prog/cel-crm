import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE AUTH CLIENT ─────────────────────────────────────────────────────
const SUPA_URL = "https://ctaxhvyepcbnjffmbudw.supabase.co";
const SUPA_KEY = "sb_publishable_TMnCzqVpMzXhmWcE64zMUQ_36Cevv8t";

// Chamada REST autenticada com token do usuário logado
const apiFetch = async (path, options = {}, token = null) => {
  const headers = {
    "apikey": SUPA_KEY,
    "Content-Type": "application/json",
    "Prefer": options.prefer || "return=representation",
    ...(token ? { "Authorization": `Bearer ${token}` } : { "Authorization": `Bearer ${SUPA_KEY}` }),
    ...options.headers,
  };
  const res = await fetch(`${SUPA_URL}/rest/v1/${path}`, { ...options, headers });
  if (!res.ok) { const t = await res.text(); throw new Error(t); }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
};

// Auth via Supabase Auth API
const authFetch = async (path, body) => {
  const res = await fetch(`${SUPA_URL}/auth/v1/${path}`, {
    method: "POST",
    headers: { "apikey": SUPA_KEY, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error_description || data.msg || "Erro de autenticação");
  return data;
};

const supaAuth = {
  signUp: (email, senha, meta) => authFetch("signup", { email, password: senha, data: meta }),
  signIn: (email, senha) => authFetch("token?grant_type=password", { email, password: senha }),
  signOut: (token) => fetch(`${SUPA_URL}/auth/v1/logout`, {
    method: "POST",
    headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${token}` },
  }),
  updateUser: async (token, campos) => {
    const res = await fetch(`${SUPA_URL}/auth/v1/user`, {
      method: "PUT",
      headers: { "apikey": SUPA_KEY, "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(campos),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error_description || data.msg || "Erro ao atualizar usuário");
    return data;
  },
  recover: async (email) => {
    const res = await fetch(`${SUPA_URL}/auth/v1/recover`, {
      method: "POST",
      headers: { "apikey": SUPA_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ email }),
    });
    if (!res.ok) { const d = await res.json(); throw new Error(d.error_description || d.msg || "Erro ao enviar e-mail"); }
    return true;
  },
};

// ─── PALETA ───────────────────────────────────────────────────────────────────
const C = {
  bordo: "#6B1A2A", vinho: "#8B2635", vinhoLight: "#A63347",
  offWhite: "#F5F0EB", gold: "#C4993A",
  dark: "#1A0A0E", darkMid: "#2D1219", muted: "#9E7A82",
  surface: "#FBF8F5", border: "#E8DDD5",
};

const CADENCIAS = {
  novo:              { proximos:[0],             label:"Contato hoje" },
  em_andamento:      { proximos:[1],             label:"24h após interação" },
  objecao:           { proximos:[1],             label:"24h após interação" },
  tem_interesse:     { proximos:[3,5,7,21,30],   label:"3→5→7→21→30d → Funil Infinito" },
  nao_responde:      { proximos:[3,5,7],         label:"3→5→7d → Funil Infinito" },
  quarta_fase:       { proximos:[7,15,30,60,75], label:"7→15→30→60→75d → Agendamento" },
  consulta_agendada: { proximos:[],              label:"Pós-Venda (outra pessoa assume)" },
  funil_infinito:    { proximos:[],              label:"Campanhas e gatilhos sazonais" },
  reativacao:        { proximos:[],              label:"Abordagem pontual" },
};

const ETAPAS = [
  { id:"novo",              label:"Novo Lead",         cor:C.gold },
  { id:"em_andamento",      label:"Em Andamento",      cor:C.vinhoLight },
  { id:"objecao",           label:"Objeção",           cor:"#B05A2A" },
  { id:"tem_interesse",     label:"Tem Interesse",     cor:"#5B7FA6" },
  { id:"nao_responde",      label:"Não Responde",      cor:C.muted },
  { id:"consulta_agendada", label:"Consulta Agendada", cor:"#4A9B6F" },
  { id:"quarta_fase",       label:"4ª Fase",           cor:"#7B4FA6" },
  { id:"funil_infinito",    label:"Funil Infinito",    cor:"#3A7A8B" },
  { id:"reativacao",        label:"Reativação",        cor:C.bordo },
];

const MAP_STATUS = {
  "follow up 1":"em_andamento","follow up 2":"tem_interesse","follow up 3":"tem_interesse",
  "funil infinito":"funil_infinito","agendado":"consulta_agendada","paciente da base":"reativacao",
  "indicacao":"novo","novo":"novo","em andamento":"em_andamento","objecao":"objecao",
  "objeção":"objecao","nao responde":"nao_responde","não responde":"nao_responde",
  "4a fase":"quarta_fase","4ª fase":"quarta_fase","reativacao":"reativacao","reativação":"reativacao",
};

const ORIGENS = ["Instagram/DM","WhatsApp direto","Indicação","Paciente da Base","Outros"];
const ADMIN_EMAIL = "shirley@benettiangular.com.br";
const WHATSAPP_SUPORTE = "5518998184929";
const NOME_APP = "Cél";
const SLOGAN_APP = "a maneira mais fácil de fazer vendas";
const MSG_BLOQUEIO_DEFAULT = "Seu acesso está suspenso. Para reativar, fale com a equipe Cél pelo WhatsApp. Estamos prontos para te atender! 💬";

const hoje = () => new Date().toISOString().split("T")[0];
const addDias = (data, dias) => { const d=new Date(data); d.setDate(d.getDate()+dias); return d.toISOString().split("T")[0]; };
const calcProximoFollowUp = (etapaId, base, idx=0) => { const c=CADENCIAS[etapaId]; if(!c?.proximos?.length) return null; return addDias(base, c.proximos[Math.min(idx,c.proximos.length-1)]); };
const proximaEtapaAposCadencia = (e) => ({tem_interesse:"funil_infinito",nao_responde:"funil_infinito",quarta_fase:"consulta_agendada"}[e]||null);
const diasAtraso = (data) => { if(!data) return null; return Math.floor((new Date(hoje())-new Date(data))/86400000); };
const alertaFollowUp = (lead) => { const d=diasAtraso(lead.follow_up_data); if(d===null) return null; if(d>0) return {tipo:"atrasado",texto:`${d}d atrasado`}; if(d===0) return {tipo:"hoje",texto:"Hoje!"}; if(d>=-2) return {tipo:"proximo",texto:`Em ${Math.abs(d)}d`}; return null; };
const campoObrigatorio = (etapaId, lead) => { if(["tem_interesse","objecao"].includes(etapaId)&&(!lead.dor||!lead.desejo)) return "Preencha as 3 Perguntas de Ouro antes."; return null; };
const gerarCodigo = (prefixo, leads) => { const nums=leads.filter(l=>l.codigo?.startsWith(prefixo)).map(l=>parseInt(l.codigo.replace(prefixo,""))||0); const max=nums.length?Math.max(...nums):0; return `${prefixo}${String(max+1).padStart(3,"0")}`; };
const filtrarPorPeriodo = (leads, periodo) => { const agora=new Date(); const ini=new Date(); if(periodo==="hoje") ini.setHours(0,0,0,0); else if(periodo==="semana") ini.setDate(agora.getDate()-7); else if(periodo==="mes") ini.setDate(1); else if(periodo==="trim") ini.setMonth(agora.getMonth()-3); else if(periodo==="ano") ini.setMonth(0,1); else return leads; return leads.filter(l=>new Date(l.criado_em)>=ini); };

const FLUXO_WHATSAPP = [
  {passo:1,label:"Conexão",descricao:"Apresentação + acolhimento. Nunca comece perguntando o que a pessoa quer."},
  {passo:2,label:"Identificação",descricao:"Descubra como ela chegou até você e o que chamou atenção."},
  {passo:3,label:"3 Perguntas de Ouro",descricao:"Dor → Desejo → Urgência. Nessa ordem, sem pular."},
  {passo:4,label:"Espelho",descricao:"Repita com suas palavras o que ela disse. Ela precisa se sentir ouvida."},
  {passo:5,label:"Solução",descricao:"Apresente o procedimento como solução para a dor, não como produto."},
  {passo:6,label:"Convite",descricao:"Convide para a consulta. Ofereça 2 datas, não pergunte 'quando você pode'."},
  {passo:7,label:"Confirmação",descricao:"Confirme 24h antes. Reforce o que ela vai ganhar, não o que vai fazer."},
];

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
const Badge = ({children,cor,small}) => <span style={{background:cor+"22",color:cor,border:`1px solid ${cor}44`,borderRadius:20,padding:small?"2px 8px":"3px 10px",fontSize:small?10:11,fontWeight:600,whiteSpace:"nowrap"}}>{children}</span>;
const AlertaTag = ({alerta}) => { if(!alerta) return null; return <Badge cor={{atrasado:"#C0392B",hoje:C.gold,proximo:"#4A9B6F"}[alerta.tipo]} small>{alerta.texto}</Badge>; };
const Modal = ({children,onClose}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(26,10,14,0.75)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={onClose}>
    <div onClick={e=>e.stopPropagation()} style={{background:C.surface,borderRadius:12,maxWidth:700,width:"100%",maxHeight:"92vh",overflowY:"auto",boxShadow:"0 24px 60px rgba(0,0,0,0.45)",border:`1px solid ${C.border}`}}>{children}</div>
  </div>
);
const ModalHeader = ({title,sub,onClose}) => (
  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"18px 24px",borderBottom:`1px solid ${C.border}`,background:C.darkMid,borderRadius:"12px 12px 0 0"}}>
    <div><div style={{color:C.offWhite,fontWeight:700,fontSize:16}}>{title}</div>{sub&&<div style={{color:C.muted,fontSize:12,marginTop:2}}>{sub}</div>}</div>
    <button onClick={onClose} style={{background:"none",border:"none",color:C.muted,fontSize:20,cursor:"pointer"}}>✕</button>
  </div>
);
const iStyle={width:"100%",padding:"9px 12px",background:C.offWhite,border:`1px solid ${C.border}`,borderRadius:8,color:C.dark,fontSize:13,outline:"none",boxSizing:"border-box"};
const lStyle={color:C.muted,fontSize:11,fontWeight:600,display:"block",marginBottom:5,textTransform:"uppercase",letterSpacing:"0.04em"};
const Logo = ({size=32,showName=false}) => (
  <div style={{display:"flex",alignItems:"center",gap:10}}>
    <div style={{width:size,height:size,borderRadius:"50%",background:`linear-gradient(135deg,${C.bordo},${C.gold})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:size*0.45,fontWeight:900,color:C.offWhite,flexShrink:0}}>C</div>
    {showName&&<div><div style={{color:C.offWhite,fontWeight:800,fontSize:size*0.5,letterSpacing:"-0.02em",lineHeight:1}}>{NOME_APP}</div><div style={{color:C.muted,fontSize:size*0.3,lineHeight:1.2}}>{SLOGAN_APP}</div></div>}
  </div>
);
const BotaoSuporte = ({small}) => {
  const msg=encodeURIComponent("Olá! Preciso de suporte com o Cél CRM.");
  return <a href={`https://wa.me/${WHATSAPP_SUPORTE}?text=${msg}`} target="_blank" rel="noreferrer" style={{display:"inline-flex",alignItems:"center",gap:6,padding:small?"5px 12px":"7px 16px",background:"#25D36622",border:"1px solid #25D36644",borderRadius:20,color:"#25D366",fontSize:small?11:12,fontWeight:700,textDecoration:"none",whiteSpace:"nowrap"}}>💬 Suporte Cél</a>;
};
const Spinner = () => <div style={{display:"flex",alignItems:"center",justifyContent:"center",padding:48,color:C.muted,fontSize:14}}>Carregando...</div>;

// ─── LOGIN ─────────────────────────────────────────────────────────────────────
function Login({onLogin}){
  const [modo,setModo]=useState("login");
  const [email,setEmail]=useState("");
  const [senha,setSenha]=useState("");
  const [nome,setNome]=useState("");
  const [erro,setErro]=useState("");
  const [sucesso,setSucesso]=useState("");
  const [loading,setLoading]=useState(false);

  const handleLogin=async()=>{
    setErro(""); setLoading(true);
    try {
      const data = await supaAuth.signIn(email, senha);
      const token = data.access_token;
      // Buscar perfil com token do próprio usuário
      const perfis = await apiFetch(`perfis?id=eq.${data.user.id}&select=*`, {}, token);
      const perfil = perfis?.[0];
      if (!perfil) { setErro("Perfil não encontrado. Contate o suporte."); setLoading(false); return; }
      if (perfil.status === "pendente") { setErro("Cadastro em análise. A equipe Cél entrará em contato em breve."); setLoading(false); return; }
      if (perfil.status === "suspenso") { onLogin({ suspenso:true, perfil, token }); setLoading(false); return; }
      if (perfil.email === ADMIN_EMAIL) { onLogin({ admin:true, perfil, token }); setLoading(false); return; }
      onLogin({ perfil, token });
    } catch(e) {
      setErro("Email ou senha incorretos.");
    }
    setLoading(false);
  };

  const handleCadastro=async()=>{
    setErro(""); setLoading(true);
    if(!nome||!email||!senha){ setErro("Preencha todos os campos."); setLoading(false); return; }
    if(senha.length < 6){ setErro("A senha precisa ter pelo menos 6 caracteres."); setLoading(false); return; }
    try {
      const prefixo = nome.split(" ").filter(Boolean).map(p=>p[0]).join("").toUpperCase().substring(0,3)||"CLI";
      await supaAuth.signUp(email, senha, { nome, prefixo });
      setSucesso("Cadastro enviado! A equipe Cél entrará em contato pelo WhatsApp para liberar seu acesso.");
      setModo("login"); setNome(""); setEmail(""); setSenha("");
    } catch(e) {
      setErro(e.message||"Erro ao cadastrar. Tente novamente.");
    }
    setLoading(false);
  };

  return(
    <div style={{minHeight:"100vh",background:C.dark,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',sans-serif"}}>
      <div style={{background:C.darkMid,border:`1px solid ${C.bordo}44`,borderRadius:16,padding:"44px 40px",width:400,boxShadow:"0 24px 60px rgba(0,0,0,0.5)"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <Logo size={56}/>
          <div style={{color:C.offWhite,fontSize:26,fontWeight:800,marginTop:14,letterSpacing:"-0.02em"}}>{NOME_APP}</div>
          <div style={{color:C.muted,fontSize:13,marginTop:4}}>{SLOGAN_APP}</div>
        </div>
        <div style={{display:"flex",background:C.dark,borderRadius:10,padding:4,marginBottom:24}}>
          {[{id:"login",label:"Entrar"},{id:"cadastro",label:"Solicitar acesso"}].map(m=>(
            <button key={m.id} onClick={()=>{setModo(m.id);setErro("");setSucesso("");}} style={{flex:1,padding:"8px",border:"none",borderRadius:8,background:modo===m.id?C.bordo:"transparent",color:modo===m.id?"#fff":C.muted,fontSize:13,fontWeight:modo===m.id?700:400,cursor:"pointer"}}>{m.label}</button>
          ))}
        </div>
        {sucesso&&<div style={{background:"#4A9B6F22",border:"1px solid #4A9B6F44",borderRadius:8,padding:"12px 14px",marginBottom:16,color:"#4A9B6F",fontSize:13,lineHeight:1.6}}>{sucesso}</div>}
        {modo==="cadastro"&&<div style={{marginBottom:14}}><label style={{...lStyle,color:C.muted}}>Nome da Clínica</label><input style={{...iStyle,background:C.dark,border:`1px solid ${C.bordo}55`,color:C.offWhite}} placeholder="Ex: Clínica Dra. Maria" value={nome} onChange={e=>{setNome(e.target.value);setErro("");}}/></div>}
        {[{label:"Email",value:email,set:setEmail,type:"email",ph:"clinica@email.com"},{label:"Senha",value:senha,set:setSenha,type:"password",ph:"mínimo 6 caracteres"}].map(f=>(
          <div key={f.label} style={{marginBottom:14}}><label style={{...lStyle,color:C.muted}}>{f.label}</label><input type={f.type} value={f.value} placeholder={f.ph} onChange={e=>{f.set(e.target.value);setErro("");}} onKeyDown={e=>e.key==="Enter"&&(modo==="login"?handleLogin():handleCadastro())} style={{...iStyle,background:C.dark,border:`1px solid ${C.bordo}55`,color:C.offWhite}}/></div>
        ))}
        {erro&&<div style={{color:"#E74C3C",fontSize:13,marginBottom:12,padding:"10px 12px",background:"#E74C3C11",borderRadius:8}}>{erro}</div>}
        <button onClick={modo==="login"?handleLogin:handleCadastro} disabled={loading} style={{width:"100%",padding:12,marginTop:4,background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,border:"none",borderRadius:8,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer",opacity:loading?0.7:1}}>
          {loading?"Aguarde...":(modo==="login"?"Entrar":"Solicitar acesso")}
        </button>
        <div style={{marginTop:20,textAlign:"center"}}><BotaoSuporte small/></div>
      </div>
    </div>
  );
}

// ─── TELA SUSPENSA ────────────────────────────────────────────────────────────
function TelaSuspensa({perfil,token,onVoltar}){
  const msg=encodeURIComponent(`Olá! Sou ${perfil.nome} e gostaria de reativar meu acesso ao ${NOME_APP}.`);
  useEffect(()=>{ supaAuth.signOut(token); },[]);
  return(
    <div style={{minHeight:"100vh",background:C.dark,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter',sans-serif",padding:24}}>
      <div style={{background:C.darkMid,border:`1px solid ${C.bordo}44`,borderRadius:16,padding:"44px 40px",maxWidth:440,width:"100%",textAlign:"center",boxShadow:"0 24px 60px rgba(0,0,0,0.5)"}}>
        <div style={{fontSize:48,marginBottom:16}}>🔒</div>
        <div style={{color:C.offWhite,fontSize:20,fontWeight:700,marginBottom:12}}>Acesso suspenso</div>
        <div style={{color:C.muted,fontSize:14,lineHeight:1.8,marginBottom:28}}>{perfil.bloqueio_msg||MSG_BLOQUEIO_DEFAULT}</div>
        <a href={`https://wa.me/${WHATSAPP_SUPORTE}?text=${msg}`} target="_blank" rel="noreferrer" style={{display:"inline-block",padding:"12px 28px",background:"#25D366",borderRadius:8,color:"#fff",fontSize:14,fontWeight:700,textDecoration:"none",marginBottom:14}}>💬 Falar com a equipe {NOME_APP}</a>
        <br/><button onClick={onVoltar} style={{background:"transparent",border:"none",color:C.muted,fontSize:13,cursor:"pointer",marginTop:8}}>← Voltar ao login</button>
      </div>
    </div>
  );
}

// ─── PAINEL ADMIN ─────────────────────────────────────────────────────────────
function PainelAdmin({token,onLogout}){
  const [perfis,setPerfis]=useState([]);
  const [loading,setLoading]=useState(true);
  const [editandoMsg,setEditandoMsg]=useState(null);
  const [msgTemp,setMsgTemp]=useState("");
  const [busca,setBusca]=useState("");
  const [filtroStatus,setFiltroStatus]=useState("todos");

  useEffect(()=>{ carregarPerfis(); },[]);

  const carregarPerfis=async()=>{
    setLoading(true);
    try{ const rows=await apiFetch("perfis?select=*",{},token); setPerfis(rows||[]); }
    catch(e){ console.error(e); }
    setLoading(false);
  };

  const atualizar=async(id,campos)=>{
    try{
      await apiFetch(`perfis?id=eq.${id}`,{method:"PATCH",body:JSON.stringify(campos),prefer:"return=minimal"},token);
      setPerfis(ps=>ps.map(p=>p.id===id?{...p,...campos}:p));
    }catch(e){ alert("Erro ao atualizar."); }
  };

  const [enviandoReset,setEnviandoReset]=useState({});
  const handleEnviarReset=async(email,id)=>{
    setEnviandoReset(s=>({...s,[id]:"loading"}));
    try{
      await supaAuth.recover(email);
      setEnviandoReset(s=>({...s,[id]:"ok"}));
      setTimeout(()=>setEnviandoReset(s=>({...s,[id]:null})),4000);
    }catch(e){
      setEnviandoReset(s=>({...s,[id]:"erro"}));
      setTimeout(()=>setEnviandoReset(s=>({...s,[id]:null})),4000);
    }
  };

  const visiveis=perfis.filter(p=>{
    if(p.email===ADMIN_EMAIL) return false;
    if(filtroStatus!=="todos"&&p.status!==filtroStatus) return false;
    if(busca&&!p.nome.toLowerCase().includes(busca.toLowerCase())&&!p.email.toLowerCase().includes(busca.toLowerCase())) return false;
    return true;
  });
  const pendentes=visiveis.filter(p=>p.status==="pendente");
  const statusCor={ativo:"#4A9B6F",pendente:C.gold,suspenso:"#C0392B"};
  const statusLabel={ativo:"Ativo",pendente:"Pendente",suspenso:"Suspenso"};

  return(
    <div style={{minHeight:"100vh",background:C.surface,fontFamily:"'Inter',sans-serif"}}>
      <div style={{background:C.dark,borderBottom:`1px solid ${C.bordo}44`,padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:54}}>
        <Logo size={32} showName/>
        <div style={{display:"flex",gap:10,alignItems:"center"}}><BotaoSuporte small/><button onClick={onLogout} style={{background:"transparent",border:`1px solid ${C.bordo}55`,color:C.muted,padding:"5px 12px",borderRadius:6,fontSize:12,cursor:"pointer"}}>Sair</button></div>
      </div>
      <div style={{padding:24,maxWidth:1000,margin:"0 auto"}}>
        {loading?<Spinner/>:<>
          {pendentes.length>0&&<div style={{background:`${C.gold}11`,border:`1px solid ${C.gold}44`,borderRadius:12,padding:"16px 20px",marginBottom:24,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{color:C.gold,fontWeight:700,fontSize:14}}>{pendentes.length} cadastro{pendentes.length>1?"s":""} aguardando aprovação</div><div style={{color:C.muted,fontSize:13,marginTop:2}}>{pendentes.map(p=>p.nome).join(", ")}</div></div>
            <button onClick={()=>setFiltroStatus("pendente")} style={{padding:"7px 14px",background:C.gold,border:"none",borderRadius:8,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>Ver agora</button>
          </div>}
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:10,marginBottom:24}}>
            {[{label:"Total",valor:visiveis.length,cor:C.bordo},{label:"Ativas",valor:visiveis.filter(p=>p.status==="ativo").length,cor:"#4A9B6F"},{label:"Pendentes",valor:visiveis.filter(p=>p.status==="pendente").length,cor:C.gold},{label:"Suspensas",valor:visiveis.filter(p=>p.status==="suspenso").length,cor:"#C0392B"}].map((m,i)=>(
              <div key={i} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:10,padding:"14px 16px"}}>
                <div style={{fontSize:26,fontWeight:800,color:m.cor}}>{m.valor}</div><div style={{fontSize:12,color:C.muted,marginTop:3}}>{m.label}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
            <input placeholder="Buscar por nome ou email..." value={busca} onChange={e=>setBusca(e.target.value)} style={{padding:"9px 14px",border:`1px solid ${C.border}`,borderRadius:8,fontSize:13,outline:"none",background:"#fff",color:C.dark,width:260}}/>
            <div style={{display:"flex",gap:6}}>{["todos","ativo","pendente","suspenso"].map(s=><button key={s} onClick={()=>setFiltroStatus(s)} style={{padding:"7px 14px",border:`1px solid ${filtroStatus===s?C.bordo:C.border}`,borderRadius:20,background:filtroStatus===s?`${C.bordo}11`:"#fff",color:filtroStatus===s?C.bordo:C.muted,fontSize:12,fontWeight:filtroStatus===s?700:400,cursor:"pointer"}}>{s==="todos"?"Todas":statusLabel[s]||s}</button>)}</div>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {visiveis.map(p=>(
              <div key={p.id} style={{background:"#fff",border:`1px solid ${C.border}`,borderLeft:`3px solid ${statusCor[p.status]||C.muted}`,borderRadius:10,padding:"14px 20px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:10}}>
                  <div><div style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontWeight:700,color:C.dark,fontSize:15}}>{p.nome}</span><Badge cor={statusCor[p.status]||C.muted} small>{statusLabel[p.status]||p.status}</Badge></div><div style={{color:C.muted,fontSize:13,marginTop:3}}>{p.email} · Desde {p.criado_em}</div></div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {p.status==="pendente"&&<button onClick={()=>atualizar(p.id,{status:"ativo"})} style={{padding:"7px 14px",border:"none",borderRadius:8,background:"#4A9B6F",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>✓ Aprovar</button>}
                    {p.status==="ativo"&&<button onClick={()=>atualizar(p.id,{status:"suspenso"})} style={{padding:"7px 14px",border:`1px solid #C0392B44`,borderRadius:8,background:"transparent",color:"#C0392B",fontSize:12,cursor:"pointer"}}>Suspender</button>}
                    {p.status==="suspenso"&&<button onClick={()=>atualizar(p.id,{status:"ativo"})} style={{padding:"7px 14px",border:"none",borderRadius:8,background:"#4A9B6F",color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>Reativar</button>}
                    <button onClick={()=>{setEditandoMsg(p.id);setMsgTemp(p.bloqueio_msg||MSG_BLOQUEIO_DEFAULT);}} style={{padding:"7px 14px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:12,cursor:"pointer"}}>✏ Mensagem</button>
                    <button onClick={()=>handleEnviarReset(p.email,p.id)} disabled={enviandoReset[p.id]==="loading"} style={{padding:"7px 14px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:enviandoReset[p.id]==="ok"?"#4A9B6F":enviandoReset[p.id]==="erro"?"#C0392B":C.muted,fontSize:12,cursor:"pointer"}}>
                      {enviandoReset[p.id]==="loading"?"Enviando...":(enviandoReset[p.id]==="ok"?"✓ E-mail enviado":(enviandoReset[p.id]==="erro"?"Erro":"🔑 Redefinir senha"))}
                    </button>
                  </div>
                </div>
                {editandoMsg===p.id&&<div style={{marginTop:14,padding:14,background:C.offWhite,borderRadius:8,border:`1px solid ${C.border}`}}>
                  <label style={{...lStyle,marginBottom:8}}>Mensagem de bloqueio para {p.nome}</label>
                  <textarea value={msgTemp} onChange={e=>setMsgTemp(e.target.value)} style={{...iStyle,height:80,resize:"vertical",marginBottom:10}}/>
                  <div style={{display:"flex",gap:8,justifyContent:"space-between",alignItems:"center"}}>
                    <button onClick={()=>setMsgTemp(MSG_BLOQUEIO_DEFAULT)} style={{background:"none",border:"none",color:C.bordo,fontSize:12,cursor:"pointer",textDecoration:"underline"}}>Restaurar padrão</button>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={()=>setEditandoMsg(null)} style={{padding:"7px 14px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:12,cursor:"pointer"}}>Cancelar</button>
                      <button onClick={()=>{atualizar(p.id,{bloqueio_msg:msgTemp});setEditandoMsg(null);}} style={{padding:"7px 14px",border:"none",borderRadius:8,background:C.bordo,color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>Salvar</button>
                    </div>
                  </div>
                </div>}
              </div>
            ))}
          </div>
        </>}
      </div>
    </div>
  );
}

// ─── MODAL IMPORTAR ───────────────────────────────────────────────────────────
function ModalImportar({onClose,onImportar,perfil,leadsExistentes}){
  const [preview,setPreview]=useState(null);
  const [nomeImport,setNomeImport]=useState("");
  const [erro,setErro]=useState("");
  const inputRef=useRef();
  const handleFile=(e)=>{
    const file=e.target.files[0]; if(!file) return;
    setNomeImport(file.name.replace(/\.(xlsx|xls)$/i,""));
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary",cellDates:true});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!rows.length){setErro("Planilha vazia.");return;}
        const mapeados=rows.map((row,i)=>{
          const get=(...ns)=>{for(const n of ns){const k=Object.keys(row).find(k=>k.toLowerCase().includes(n));if(k&&row[k]!=="")return String(row[k]).trim();}return "";};
          const statusRaw=get("status","etapa","situacao").toLowerCase();
          const etapa=MAP_STATUS[statusRaw]||"novo";
          const codRaw=get("cod","codigo","code");
          const codigo=codRaw||gerarCodigo(perfil.prefixo,[...leadsExistentes]);
          const followUpRaw=get("follow-up","followup","prox");
          let follow_up_data=""; if(followUpRaw){const d=new Date(followUpRaw);if(!isNaN(d))follow_up_data=d.toISOString().split("T")[0];}
          const dataRegRaw=get("data","registro","criado");
          let criado_em=hoje(); if(dataRegRaw){const d=new Date(dataRegRaw);if(!isNaN(d))criado_em=d.toISOString().split("T")[0];}
          return{clinica_id:perfil.id,codigo,nome:get("nome","name","paciente")||`Lead ${i+1}`,telefone:get("telefone","fone","celular","whatsapp"),origem:get("origem","source","canal")||"Outros",etapa,ultima_interacao:criado_em,follow_up_data:follow_up_data||calcProximoFollowUp(etapa,criado_em,0)||"",indice_follow_up:0,criado_em,observacoes:get("obs","observa","nota","anotacao"),dor:"",desejo:"",urgencia:"",historico:"[]",_statusOriginal:get("status","etapa","situacao")};
        });
        setPreview(mapeados);
      }catch(err){setErro("Erro ao ler arquivo.");}
    };
    reader.readAsBinaryString(file);
  };
  const etapaLabel=(id)=>ETAPAS.find(e=>e.id===id)?.label||id;
  const etapaCor=(id)=>ETAPAS.find(e=>e.id===id)?.cor||C.muted;
  return(
    <Modal onClose={onClose}>
      <ModalHeader title="Importar Planilha Excel" sub="Mapeia automaticamente para as etapas do CRM" onClose={onClose}/>
      <div style={{padding:24}}>
        {!preview?(
          <div>
            <div style={{border:`2px dashed ${C.border}`,borderRadius:12,padding:"40px 24px",textAlign:"center",cursor:"pointer",background:C.offWhite}} onClick={()=>inputRef.current?.click()}>
              <div style={{fontSize:36,marginBottom:12}}>📊</div>
              <div style={{fontWeight:700,color:C.dark,fontSize:15,marginBottom:6}}>Clique para selecionar o arquivo</div>
              <div style={{color:C.muted,fontSize:13}}>Suporta .xlsx e .xls</div>
              <input ref={inputRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
            </div>
            {erro&&<div style={{color:"#C0392B",fontSize:13,marginTop:12,padding:"10px 14px",background:"#C0392B11",borderRadius:8}}>⚠ {erro}</div>}
          </div>
        ):(
          <div>
            <div style={{marginBottom:14}}><label style={lStyle}>Nome da importação</label><input style={iStyle} value={nomeImport} onChange={e=>setNomeImport(e.target.value)} placeholder="Ex: Base março 2026"/></div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div><div style={{fontWeight:700,color:C.dark,fontSize:15}}>{preview.length} leads encontrados</div><div style={{color:C.muted,fontSize:13}}>Revise antes de importar</div></div>
              <button onClick={()=>setPreview(null)} style={{padding:"6px 12px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:13,cursor:"pointer"}}>← Trocar arquivo</button>
            </div>
            <div style={{maxHeight:300,overflowY:"auto",border:`1px solid ${C.border}`,borderRadius:10}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:C.darkMid}}>{["Código","Nome","Telefone","Etapa CRM","Status Original","Follow-up"].map(h=><th key={h} style={{padding:"8px 10px",color:C.offWhite,fontWeight:600,textAlign:"left",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
                <tbody>{preview.map((l,i)=><tr key={i} style={{borderBottom:`1px solid ${C.border}`,background:i%2===0?"#fff":C.offWhite}}><td style={{padding:"7px 10px",fontWeight:700,color:C.bordo}}>{l.codigo}</td><td style={{padding:"7px 10px",color:C.dark,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{l.nome}</td><td style={{padding:"7px 10px",color:C.muted}}>{l.telefone}</td><td style={{padding:"7px 10px"}}><Badge cor={etapaCor(l.etapa)} small>{etapaLabel(l.etapa)}</Badge></td><td style={{padding:"7px 10px",color:C.muted,fontStyle:"italic"}}>{l._statusOriginal}</td><td style={{padding:"7px 10px",color:C.muted}}>{l.follow_up_data}</td></tr>)}</tbody>
              </table>
            </div>
            <div style={{marginTop:14,display:"flex",justifyContent:"flex-end",gap:10}}>
              <button onClick={onClose} style={{padding:"9px 16px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:13,cursor:"pointer"}}>Cancelar</button>
              <button onClick={()=>{const importId=`imp_${Date.now()}`;onImportar(preview.map(({_statusOriginal,...l})=>({...l,import_id:importId,import_nome:nomeImport||"Importação"})));onClose();}} style={{padding:"9px 20px",border:"none",borderRadius:8,background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>✓ Importar {preview.length} leads</button>
            </div>
          </div>
        )}
      </div>
    </Modal>
  );
}

// ─── MODAL LEAD ───────────────────────────────────────────────────────────────
function ModalLead({lead,onClose,onSave,onDelete}){
  const [form,setForm]=useState({...lead});
  const [aba,setAba]=useState("dados");
  const [erroEtapa,setErroEtapa]=useState("");
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const handleEtapaChange=(novaEtapa)=>{
    const erro=campoObrigatorio(novaEtapa,form); if(erro){setErroEtapa(erro);return;} setErroEtapa("");
    const base=hoje();
    setForm(f=>({...f,etapa:novaEtapa,ultima_interacao:base,follow_up_data:calcProximoFollowUp(novaEtapa,base,0)||f.follow_up_data,indice_follow_up:0,historico:[...(f.historico||[]),{data:base,de:f.etapa,para:novaEtapa}]}));
  };
  const handleRegistrarContato=()=>{
    const c=CADENCIAS[form.etapa]; if(!c?.proximos?.length)return;
    const novoIdx=form.indice_follow_up+1; const base=hoje();
    const atingiu=novoIdx>=c.proximos.length;
    const proxEtapa=atingiu?proximaEtapaAposCadencia(form.etapa):null;
    const novoFU=!atingiu?calcProximoFollowUp(form.etapa,base,novoIdx):null;
    setForm(f=>({...f,ultima_interacao:base,indice_follow_up:novoIdx,follow_up_data:novoFU||f.follow_up_data,etapa:proxEtapa||f.etapa,historico:[...(f.historico||[]),{data:base,de:f.etapa,para:proxEtapa||f.etapa,nota:atingiu?"Cadência encerrada":`Follow-up ${novoIdx+1}`}]}));
  };
  const etapaAtual=ETAPAS.find(e=>e.id===form.etapa);
  const cadencia=CADENCIAS[form.etapa];
  const proximos=cadencia?.proximos||[];
  const atingiuFim=proximos.length>0&&form.indice_follow_up>=proximos.length;
  return(
    <Modal onClose={onClose}>
      <ModalHeader title={`${form.codigo?`[${form.codigo}] `:""}${form.nome||"Novo Lead"}`} sub={etapaAtual?.label} onClose={onClose}/>
      <div style={{display:"flex",borderBottom:`1px solid ${C.border}`,background:"#fff"}}>
        {[{id:"dados",label:"Dados"},{id:"regras",label:"Cadência"},{id:"perguntas",label:"3 Perguntas"},{id:"historico",label:"Histórico"}].map(a=><button key={a.id} onClick={()=>setAba(a.id)} style={{padding:"11px 16px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontWeight:aba===a.id?700:400,color:aba===a.id?C.bordo:C.muted,borderBottom:aba===a.id?`2px solid ${C.bordo}`:"2px solid transparent"}}>{a.label}</button>)}
      </div>
      <div style={{padding:24}}>
        {aba==="dados"&&<div style={{display:"grid",gap:14}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 2fr 1fr",gap:12}}>
            <div><label style={lStyle}>Código</label><input style={iStyle} value={form.codigo||""} onChange={e=>set("codigo",e.target.value)} placeholder="PAC001"/></div>
            <div><label style={lStyle}>Nome Completo</label><input style={iStyle} value={form.nome} onChange={e=>set("nome",e.target.value)}/></div>
            <div><label style={lStyle}>Telefone</label><input style={iStyle} value={form.telefone||""} onChange={e=>set("telefone",e.target.value)}/></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <div><label style={lStyle}>Origem</label><select style={iStyle} value={form.origem||"Outros"} onChange={e=>set("origem",e.target.value)}>{ORIGENS.map(o=><option key={o}>{o}</option>)}</select></div>
            <div><label style={lStyle}>Etapa</label><select style={iStyle} value={form.etapa} onChange={e=>handleEtapaChange(e.target.value)}>{ETAPAS.map(e=><option key={e.id} value={e.id}>{e.label}</option>)}</select>{erroEtapa&&<div style={{color:"#C0392B",fontSize:12,marginTop:4}}>⚠ {erroEtapa}</div>}</div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <div><label style={lStyle}>Última Interação</label><input style={iStyle} type="date" value={form.ultima_interacao||""} onChange={e=>set("ultima_interacao",e.target.value)}/></div>
            <div><label style={lStyle}>Próximo Follow-up</label><input style={iStyle} type="date" value={form.follow_up_data||""} onChange={e=>set("follow_up_data",e.target.value)}/></div>
          </div>
          <div><label style={lStyle}>Observações</label><textarea style={{...iStyle,height:72,resize:"vertical"}} value={form.observacoes||""} onChange={e=>set("observacoes",e.target.value)}/></div>
        </div>}
        {aba==="regras"&&<div>
          <div style={{padding:16,borderRadius:10,background:etapaAtual?.cor+"11",border:`1px solid ${etapaAtual?.cor}33`,marginBottom:20}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div><div style={{fontWeight:700,color:etapaAtual?.cor,fontSize:14,marginBottom:4}}>{etapaAtual?.label}</div><div style={{color:C.muted,fontSize:13}}>{cadencia?.label}</div></div>
              {proximos.length>0&&<Badge cor={etapaAtual?.cor}>{atingiuFim?"Fim da cadência":`Contato ${form.indice_follow_up+1} de ${proximos.length}`}</Badge>}
            </div>
            {proximos.length>0&&<div style={{marginTop:14}}><div style={{color:C.muted,fontSize:11,marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>Sequência</div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{proximos.map((d,i)=><div key={i} style={{padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:600,background:i<form.indice_follow_up?C.darkMid:i===form.indice_follow_up?etapaAtual?.cor:"#fff",color:i<form.indice_follow_up?C.muted:i===form.indice_follow_up?"#fff":C.muted,border:`1px solid ${i===form.indice_follow_up?etapaAtual?.cor:C.border}`,textDecoration:i<form.indice_follow_up?"line-through":"none"}}>{i<form.indice_follow_up?"✓ ":""}{d}d</div>)}<div style={{padding:"4px 12px",borderRadius:20,fontSize:12,background:"#fff",border:`1px solid ${C.border}`,color:C.muted}}>→ {proximaEtapaAposCadencia(form.etapa)?ETAPAS.find(e=>e.id===proximaEtapaAposCadencia(form.etapa))?.label:"Agendamento"}</div></div></div>}
            {proximos.length>0&&!atingiuFim&&<button onClick={handleRegistrarContato} style={{marginTop:16,padding:"9px 18px",border:"none",borderRadius:8,background:etapaAtual?.cor,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>✓ Registrar contato realizado</button>}
            {form.etapa==="consulta_agendada"&&<div style={{marginTop:10,padding:"10px 14px",background:"#4A9B6F22",borderRadius:8,color:"#4A9B6F",fontSize:13,fontWeight:600}}>✓ Passa para Pós-Venda. Outra pessoa assume.</div>}
            {form.etapa==="funil_infinito"&&<div style={{marginTop:10,padding:"10px 14px",background:"#3A7A8B22",borderRadius:8,color:"#3A7A8B",fontSize:13}}>Lead no funil infinito. Abordagem por campanhas e gatilhos.</div>}
          </div>
          <div style={{color:C.dark,fontWeight:700,fontSize:13,marginBottom:10}}>Fluxo WhatsApp — 7 Etapas</div>
          {FLUXO_WHATSAPP.map((p,i)=><div key={i} style={{display:"flex",gap:12,marginBottom:8,padding:"10px 14px",background:C.offWhite,borderRadius:8,border:`1px solid ${C.border}`}}><div style={{width:26,height:26,borderRadius:"50%",minWidth:26,background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:12}}>{p.passo}</div><div><div style={{fontWeight:700,color:C.dark,fontSize:13}}>{p.label}</div><div style={{color:C.muted,fontSize:12,lineHeight:1.5}}>{p.descricao}</div></div></div>)}
        </div>}
        {aba==="perguntas"&&<div>
          <p style={{color:C.muted,fontSize:13,marginBottom:18,lineHeight:1.6}}>Obrigatório para mover para Tem Interesse ou Objeção.</p>
          {[{key:"dor",label:"1. Qual é a DOR?",hint:"O que está incomodando?",cor:C.bordo},{key:"desejo",label:"2. Qual é o DESEJO?",hint:"Como ela quer se sentir depois?",cor:"#5B7FA6"},{key:"urgencia",label:"3. Qual é a URGÊNCIA?",hint:"Evento, prazo ou motivo que acelera?",cor:C.gold}].map(p=><div key={p.key} style={{marginBottom:14,padding:14,borderRadius:10,border:`1px solid ${form[p.key]?p.cor+"44":C.border}`,background:form[p.key]?p.cor+"08":C.offWhite}}><label style={{...lStyle,color:p.cor,marginBottom:3}}>{p.label}</label><div style={{color:C.muted,fontSize:12,marginBottom:7,fontStyle:"italic"}}>{p.hint}</div><textarea style={{...iStyle,height:56,resize:"none"}} value={form[p.key]||""} onChange={e=>set(p.key,e.target.value)} placeholder="Escreva aqui..."/></div>)}
          {form.dor&&form.desejo&&<div style={{padding:14,background:`${C.gold}11`,border:`1px solid ${C.gold}33`,borderRadius:10}}><div style={{color:C.gold,fontWeight:700,fontSize:13,marginBottom:6}}>✨ Espelho</div><div style={{color:C.dark,fontSize:13,lineHeight:1.7}}>"Então, você está sentindo <strong>{form.dor}</strong> e quer <strong>{form.desejo}</strong>{form.urgencia?`, especialmente porque ${form.urgencia}`:""}. Faz sentido?"</div></div>}
        </div>}
        {aba==="historico"&&<div>
          {(!form.historico||form.historico.length===0)&&<div style={{color:C.muted,fontSize:13,fontStyle:"italic"}}>Nenhuma interação registrada ainda.</div>}
          {[...(form.historico||[])].reverse().map((h,i)=><div key={i} style={{display:"flex",gap:12,marginBottom:8,padding:"10px 14px",background:C.offWhite,borderRadius:8,border:`1px solid ${C.border}`}}><div style={{color:C.muted,fontSize:12,minWidth:88}}>{h.data}</div><div><span style={{color:C.dark,fontSize:13}}>{ETAPAS.find(e=>e.id===h.de)?.label} → <strong>{ETAPAS.find(e=>e.id===h.para)?.label}</strong></span>{h.nota&&<div style={{color:C.muted,fontSize:12,marginTop:2}}>{h.nota}</div>}</div></div>)}
        </div>}
      </div>
      <div style={{padding:"14px 24px",borderTop:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",background:C.surface,borderRadius:"0 0 12px 12px"}}>
        <button onClick={()=>onDelete(lead.id)} style={{padding:"8px 14px",border:`1px solid #E74C3C44`,borderRadius:8,background:"transparent",color:"#E74C3C",fontSize:13,cursor:"pointer"}}>Excluir</button>
        <div style={{display:"flex",gap:10}}>
          <button onClick={onClose} style={{padding:"8px 14px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:13,cursor:"pointer"}}>Cancelar</button>
          <button onClick={async()=>{setSaving(true);await onSave(form);setSaving(false);onClose();}} disabled={saving} style={{padding:"8px 20px",border:"none",borderRadius:8,background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",opacity:saving?0.7:1}}>{saving?"Salvando...":"Salvar"}</button>
        </div>
      </div>
    </Modal>
  );
}

// ─── CRM ─────────────────────────────────────────────────────────────────────
function CRM({perfil,token,onLogout}){
  const [leads,setLeads]=useState([]);
  const [loading,setLoading]=useState(true);
  const [leadSelecionado,setLeadSelecionado]=useState(null);
  const [novoLead,setNovoLead]=useState(false);
  const [importando,setImportando]=useState(false);
  const [filtroEtapa,setFiltroEtapa]=useState("todos");
  const [busca,setBusca]=useState("");
  const [abaAtiva,setAbaAtiva]=useState("followups");
  const [periodo,setPeriodo]=useState("mes");
  const [dataInicio,setDataInicio]=useState("");
  const [dataFim,setDataFim]=useState("");
  const [usarPersonalizado,setUsarPersonalizado]=useState(false);

  useEffect(()=>{ carregarLeads(); },[]);

  const carregarLeads=async()=>{
    setLoading(true);
    try{
      const rows=await apiFetch(`leads?clinica_id=eq.${perfil.id}&select=*&order=criado_em.desc`,{},token);
      setLeads((rows||[]).map(l=>({...l,historico:typeof l.historico==="string"?JSON.parse(l.historico||"[]"):l.historico||[]})));
    }catch(e){console.error(e);}
    setLoading(false);
  };

  const leadsFiltrados=leads.filter(l=>{
    if(filtroEtapa!=="todos"&&l.etapa!==filtroEtapa)return false;
    if(busca&&!l.nome?.toLowerCase().includes(busca.toLowerCase())&&!l.codigo?.toLowerCase().includes(busca.toLowerCase())&&!l.telefone?.includes(busca))return false;
    return true;
  });
  const followupsHoje=leads.filter(l=>{const d=diasAtraso(l.follow_up_data);return d!==null&&d>=0;}).sort((a,b)=>diasAtraso(b.follow_up_data)-diasAtraso(a.follow_up_data));
  const leadsNoPeriodo=usarPersonalizado&&dataInicio&&dataFim?leads.filter(l=>l.criado_em>=dataInicio&&l.criado_em<=dataFim):filtrarPorPeriodo(leads,periodo);
  const importacoes=[...new Map(leads.filter(l=>l.import_id).map(l=>[l.import_id,{id:l.import_id,nome:l.import_nome||"Importação",qtd:leads.filter(x=>x.import_id===l.import_id).length}])).values()];
  const totalPorEtapa=(id)=>leads.filter(l=>l.etapa===id).length;

  const salvarLead=async(form)=>{
    const payload={codigo:form.codigo,nome:form.nome,telefone:form.telefone,origem:form.origem,etapa:form.etapa,ultima_interacao:form.ultima_interacao,follow_up_data:form.follow_up_data,indice_follow_up:form.indice_follow_up,observacoes:form.observacoes,dor:form.dor,desejo:form.desejo,urgencia:form.urgencia,historico:JSON.stringify(form.historico||[])};
    try{
      await apiFetch(`leads?id=eq.${form.id}`,{method:"PATCH",body:JSON.stringify(payload),prefer:"return=minimal"},token);
      setLeads(ls=>ls.map(l=>l.id===form.id?{...l,...form}:l));
    }catch(e){alert("Erro ao salvar.");}
  };

  const excluirLead=async(id)=>{
    try{
      await apiFetch(`leads?id=eq.${id}`,{method:"DELETE",prefer:"return=minimal"},token);
      setLeads(ls=>ls.filter(l=>l.id!==id)); setLeadSelecionado(null);
    }catch(e){alert("Erro ao excluir.");}
  };

  const excluirImportacao=async(importId)=>{
    if(!window.confirm("Excluir esta importação e todos os leads dela?"))return;
    try{
      await apiFetch(`leads?import_id=eq.${importId}&clinica_id=eq.${perfil.id}`,{method:"DELETE",prefer:"return=minimal"},token);
      setLeads(ls=>ls.filter(l=>l.import_id!==importId));
    }catch(e){alert("Erro ao excluir importação.");}
  };

  const criarLead=async(form)=>{
    const codigo=form.codigo||gerarCodigo(perfil.prefixo,leads);
    const payload={clinica_id:perfil.id,codigo,nome:form.nome,telefone:form.telefone,origem:form.origem,etapa:form.etapa,ultima_interacao:hoje(),follow_up_data:calcProximoFollowUp(form.etapa,hoje(),0)||hoje(),indice_follow_up:0,criado_em:hoje(),observacoes:form.observacoes||"",dor:"",desejo:"",urgencia:"",historico:"[]"};
    try{
      const rows=await apiFetch("leads",{method:"POST",body:JSON.stringify(payload)},token);
      if(rows?.[0]) setLeads(ls=>[{...rows[0],historico:[]}, ...ls]);
    }catch(e){alert("Erro ao criar lead.");}
    setNovoLead(false);
  };

  const importarLeads=async(novos)=>{
    try{
      const rows=await apiFetch("leads",{method:"POST",body:JSON.stringify(novos)},token);
      if(rows) setLeads(ls=>[...(rows.map(r=>({...r,historico:[]}))), ...ls]);
    }catch(e){alert("Erro ao importar leads.");}
  };

  const handleLogout=async()=>{ await supaAuth.signOut(token); onLogout(); };

  const [alterandoSenha,setAlterandoSenha]=useState(false);
  const [novaSenha,setNovaSenha]=useState("");
  const [senhaMsg,setSenhaMsg]=useState("");
  const [senhaErro,setSenhaErro]=useState("");
  const [senhaLoading,setSenhaLoading]=useState(false);

  const handleAlterarSenha=async()=>{
    setSenhaErro(""); setSenhaMsg("");
    if(novaSenha.length<6){setSenhaErro("A senha precisa ter pelo menos 6 caracteres.");return;}
    setSenhaLoading(true);
    try{
      await supaAuth.updateUser(token,{password:novaSenha});
      setSenhaMsg("Senha alterada com sucesso!");
      setNovaSenha("");
    }catch(e){ setSenhaErro(e.message||"Erro ao alterar senha."); }
    setSenhaLoading(false);
  };

  const leadVazio={id:null,codigo:"",nome:"",telefone:"",origem:"Instagram/DM",etapa:"novo",ultima_interacao:hoje(),follow_up_data:hoje(),indice_follow_up:0,criado_em:hoje(),observacoes:"",dor:"",desejo:"",urgencia:"",historico:[]};
  const periodos=[{id:"hoje",label:"Hoje"},{id:"semana",label:"Semana"},{id:"mes",label:"Mês"},{id:"trim",label:"Trimestre"},{id:"ano",label:"Ano"},{id:"todos",label:"Tudo"}];

  return(
    <div style={{minHeight:"100vh",background:C.surface,fontFamily:"'Inter',sans-serif",display:"flex",flexDirection:"column"}}>
      <div style={{background:C.dark,borderBottom:`1px solid ${C.bordo}44`,padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:54}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}><Logo size={30} showName/><span style={{color:C.muted,fontSize:11,marginLeft:4}}>| {perfil.nome}</span></div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {followupsHoje.length>0&&<button onClick={()=>setAbaAtiva("followups")} style={{background:`${C.gold}22`,border:`1px solid ${C.gold}44`,borderRadius:20,padding:"4px 12px",fontSize:12,color:C.gold,fontWeight:700,cursor:"pointer"}}>{followupsHoje.length} follow-up{followupsHoje.length>1?"s":""} hoje</button>}
          <BotaoSuporte small/>
          <button onClick={()=>{setAlterandoSenha(true);setSenhaMsg("");setSenhaErro("");setNovaSenha("");}} style={{background:"transparent",border:`1px solid ${C.bordo}55`,color:C.muted,padding:"5px 12px",borderRadius:6,fontSize:12,cursor:"pointer"}}>🔑 Senha</button>
          <button onClick={handleLogout} style={{background:"transparent",border:`1px solid ${C.bordo}55`,color:C.muted,padding:"5px 12px",borderRadius:6,fontSize:12,cursor:"pointer"}}>Sair</button>
        </div>
      </div>
      <div style={{background:"#fff",borderBottom:`1px solid ${C.border}`,display:"flex",padding:"0 24px",gap:2}}>
        {[{id:"followups",label:`Follow-ups Hoje${followupsHoje.length>0?` (${followupsHoje.length})`:""}`},{id:"pipeline",label:"Pipeline"},{id:"metricas",label:"Métricas"}].map(a=><button key={a.id} onClick={()=>setAbaAtiva(a.id)} style={{padding:"12px 16px",border:"none",background:"none",cursor:"pointer",fontSize:13,fontWeight:abaAtiva===a.id?700:400,color:abaAtiva===a.id?C.bordo:C.muted,borderBottom:abaAtiva===a.id?`2px solid ${C.bordo}`:"2px solid transparent"}}>{a.label}</button>)}
      </div>
      <div style={{flex:1,padding:24,maxWidth:1100,margin:"0 auto",width:"100%",boxSizing:"border-box"}}>
        {loading?<Spinner/>:<>
          {abaAtiva==="followups"&&<div style={{maxWidth:680}}>
            <div style={{marginBottom:20}}><h2 style={{color:C.dark,fontSize:20,fontWeight:800,margin:"0 0 4px"}}>Follow-ups do Dia</h2><p style={{color:C.muted,fontSize:14,margin:0}}>Contatos que precisam acontecer hoje.</p></div>
            {followupsHoje.length===0&&<div style={{textAlign:"center",padding:48,color:C.muted,fontSize:14,background:"#fff",borderRadius:12,border:`1px solid ${C.border}`}}>✓ Nenhum follow-up pendente para hoje.</div>}
            {followupsHoje.map(lead=>{const etapa=ETAPAS.find(e=>e.id===lead.etapa);const alerta=alertaFollowUp(lead);const cadencia=CADENCIAS[lead.etapa];return(<div key={lead.id} onClick={()=>setLeadSelecionado(lead)} style={{background:"#fff",border:`1px solid ${C.border}`,borderLeft:`4px solid ${etapa?.cor||C.muted}`,borderRadius:10,padding:"14px 18px",cursor:"pointer",marginBottom:10,display:"flex",gap:14,alignItems:"flex-start"}}><div style={{width:38,height:38,borderRadius:"50%",minWidth:38,background:`${etapa?.cor||C.muted}22`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:700,color:etapa?.cor||C.muted}}>{lead.nome?.charAt(0)}</div><div style={{flex:1}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div><span style={{fontWeight:700,color:C.dark,fontSize:15}}>{lead.nome}</span>{lead.codigo&&<span style={{marginLeft:8,fontSize:12,fontWeight:700,color:C.bordo}}>{lead.codigo}</span>}</div><AlertaTag alerta={alerta}/></div><div style={{color:C.muted,fontSize:12,marginTop:2}}>{lead.telefone} · {lead.origem}</div><div style={{display:"flex",gap:6,marginTop:8,flexWrap:"wrap"}}>{etapa&&<Badge cor={etapa.cor} small>{etapa.label}</Badge>}{cadencia?.proximos?.length>0&&<Badge cor={C.muted} small>Contato {lead.indice_follow_up+1}/{cadencia.proximos.length}</Badge>}{lead.observacoes&&<span style={{color:C.muted,fontSize:12,fontStyle:"italic"}}>{lead.observacoes.substring(0,50)}{lead.observacoes.length>50?"...":""}</span>}</div></div></div>);})}
          </div>}
          {abaAtiva==="pipeline"&&<>
            <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
              <input placeholder="Buscar por nome, código ou telefone..." value={busca} onChange={e=>setBusca(e.target.value)} style={{padding:"9px 14px",border:`1px solid ${C.border}`,borderRadius:8,fontSize:13,outline:"none",background:"#fff",color:C.dark,width:260}}/>
              <select value={filtroEtapa} onChange={e=>setFiltroEtapa(e.target.value)} style={{padding:"9px 14px",border:`1px solid ${C.border}`,borderRadius:8,fontSize:13,outline:"none",background:"#fff",color:C.dark}}><option value="todos">Todas as etapas</option>{ETAPAS.map(e=><option key={e.id} value={e.id}>{e.label} ({totalPorEtapa(e.id)})</option>)}</select>
              <div style={{marginLeft:"auto",display:"flex",gap:8}}>
                <button onClick={()=>setImportando(true)} style={{padding:"9px 16px",border:`1px solid ${C.border}`,borderRadius:8,background:"#fff",color:C.dark,fontSize:13,fontWeight:600,cursor:"pointer"}}>📊 Importar Excel</button>
                <button onClick={()=>setNovoLead(true)} style={{padding:"9px 18px",background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,border:"none",borderRadius:8,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer"}}>+ Novo lead</button>
              </div>
            </div>
            {importacoes.length>0&&<div style={{marginBottom:20,padding:14,background:"#fff",border:`1px solid ${C.border}`,borderRadius:10}}>
              <div style={{color:C.dark,fontWeight:700,fontSize:13,marginBottom:10}}>Planilhas importadas</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {importacoes.map(imp=><div key={imp.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 12px",background:C.offWhite,border:`1px solid ${C.border}`,borderRadius:8}}>
                  <span style={{fontSize:13,color:C.dark}}>📊 {imp.nome}</span><span style={{fontSize:11,color:C.muted}}>{imp.qtd} leads</span>
                  <button onClick={()=>excluirImportacao(imp.id)} style={{background:"none",border:"none",color:"#C0392B",cursor:"pointer",fontSize:14,lineHeight:1,padding:"0 2px"}}>✕</button>
                </div>)}
              </div>
            </div>}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(120px,1fr))",gap:8,marginBottom:24}}>
              {ETAPAS.map(e=><div key={e.id} onClick={()=>setFiltroEtapa(filtroEtapa===e.id?"todos":e.id)} style={{padding:"10px 12px",borderRadius:10,cursor:"pointer",background:filtroEtapa===e.id?e.cor+"22":"#fff",border:`1px solid ${filtroEtapa===e.id?e.cor:C.border}`}}><div style={{fontSize:22,fontWeight:800,color:e.cor}}>{totalPorEtapa(e.id)}</div><div style={{fontSize:11,color:C.muted,marginTop:2,lineHeight:1.3}}>{e.label}</div></div>)}
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {leadsFiltrados.length===0&&<div style={{textAlign:"center",padding:48,color:C.muted,fontSize:14}}>Nenhum lead encontrado</div>}
              {leadsFiltrados.map(lead=>{const etapa=ETAPAS.find(e=>e.id===lead.etapa);const alerta=alertaFollowUp(lead);const semPerguntas=!lead.dor||!lead.desejo;return(<div key={lead.id} onClick={()=>setLeadSelecionado(lead)} style={{background:"#fff",border:`1px solid ${C.border}`,borderLeft:`3px solid ${etapa?.cor||C.border}`,borderRadius:10,padding:"12px 18px",cursor:"pointer",display:"flex",alignItems:"center",gap:14}}><div style={{width:36,height:36,borderRadius:"50%",minWidth:36,background:`${etapa?.cor||C.muted}22`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:etapa?.cor||C.muted}}>{lead.nome?.charAt(0)}</div><div style={{flex:1,minWidth:0}}><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:700,color:C.dark,fontSize:14}}>{lead.nome}</span>{lead.codigo&&<span style={{fontSize:11,fontWeight:700,color:C.bordo,background:C.bordo+"11",padding:"2px 7px",borderRadius:10}}>{lead.codigo}</span>}{lead.import_nome&&<span style={{fontSize:10,color:C.muted,background:C.offWhite,padding:"2px 6px",borderRadius:6,border:`1px solid ${C.border}`}}>📊 {lead.import_nome}</span>}</div><div style={{color:C.muted,fontSize:12,marginTop:1}}>{lead.telefone} · {lead.origem}</div></div><div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",justifyContent:"flex-end"}}>{etapa&&<Badge cor={etapa.cor} small>{etapa.label}</Badge>}<AlertaTag alerta={alerta}/>{semPerguntas&&["tem_interesse","objecao"].includes(lead.etapa)&&<Badge cor={C.gold} small>⚠ Perguntas</Badge>}</div></div>);})}
            </div>
          </>}
          {abaAtiva==="metricas"&&<div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:24,flexWrap:"wrap",gap:12}}>
              <div><h2 style={{color:C.dark,fontSize:20,fontWeight:800,margin:"0 0 4px"}}>Métricas de Oportunidades</h2><p style={{color:C.muted,fontSize:14,margin:0}}>Histórico completo — filtre por período.</p></div>
              <div style={{display:"flex",flexDirection:"column",gap:8,alignItems:"flex-end"}}>
                <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                  {periodos.map(p=><button key={p.id} onClick={()=>{setPeriodo(p.id);setUsarPersonalizado(false);}} style={{padding:"6px 12px",border:`1px solid ${!usarPersonalizado&&periodo===p.id?C.bordo:C.border}`,borderRadius:20,background:!usarPersonalizado&&periodo===p.id?`${C.bordo}11`:"#fff",color:!usarPersonalizado&&periodo===p.id?C.bordo:C.muted,fontSize:12,fontWeight:!usarPersonalizado&&periodo===p.id?700:400,cursor:"pointer"}}>{p.label}</button>)}
                  <button onClick={()=>setUsarPersonalizado(v=>!v)} style={{padding:"6px 12px",border:`1px solid ${usarPersonalizado?C.bordo:C.border}`,borderRadius:20,background:usarPersonalizado?`${C.bordo}11`:"#fff",color:usarPersonalizado?C.bordo:C.muted,fontSize:12,fontWeight:usarPersonalizado?700:400,cursor:"pointer"}}>📅 Personalizado</button>
                </div>
                {usarPersonalizado&&<div style={{display:"flex",gap:8,alignItems:"center",background:"#fff",border:`1px solid ${C.bordo}44`,borderRadius:10,padding:"10px 14px"}}>
                  <div><label style={{...lStyle,marginBottom:3}}>De</label><input type="date" value={dataInicio} onChange={e=>setDataInicio(e.target.value)} style={{...iStyle,width:140,padding:"7px 10px"}}/></div>
                  <div style={{color:C.muted,fontSize:13,marginTop:16}}>→</div>
                  <div><label style={{...lStyle,marginBottom:3}}>Até</label><input type="date" value={dataFim} onChange={e=>setDataFim(e.target.value)} style={{...iStyle,width:140,padding:"7px 10px"}}/></div>
                </div>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:12,marginBottom:32}}>
              {[{label:"Total de leads",valor:leadsNoPeriodo.length,cor:C.bordo},{label:"Consultas agendadas",valor:leadsNoPeriodo.filter(l=>l.etapa==="consulta_agendada").length,cor:"#4A9B6F"},{label:"Em cadência ativa",valor:leadsNoPeriodo.filter(l=>["em_andamento","objecao","tem_interesse","nao_responde","quarta_fase"].includes(l.etapa)).length,cor:"#5B7FA6"},{label:"Funil Infinito",valor:leadsNoPeriodo.filter(l=>l.etapa==="funil_infinito").length,cor:"#3A7A8B"},{label:"Novos leads",valor:leadsNoPeriodo.filter(l=>l.etapa==="novo").length,cor:C.gold},{label:"Taxa de agendamento",valor:leadsNoPeriodo.length>0?Math.round(leadsNoPeriodo.filter(l=>l.etapa==="consulta_agendada").length/leadsNoPeriodo.length*100)+"%":"0%",cor:C.vinho}].map((m,i)=><div key={i} style={{background:"#fff",border:`1px solid ${C.border}`,borderRadius:12,padding:"16px 18px"}}><div style={{fontSize:28,fontWeight:800,color:m.cor}}>{m.valor}</div><div style={{fontSize:12,color:C.muted,marginTop:4,lineHeight:1.3}}>{m.label}</div></div>)}
            </div>
            <div style={{color:C.dark,fontWeight:700,fontSize:14,marginBottom:14}}>Distribuição por etapa</div>
            {ETAPAS.map(e=>{const qtd=leadsNoPeriodo.filter(l=>l.etapa===e.id).length;const pct=leadsNoPeriodo.length>0?(qtd/leadsNoPeriodo.length)*100:0;return(<div key={e.id} style={{marginBottom:10}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:13,color:C.dark}}>{e.label}</span><span style={{fontSize:13,fontWeight:700,color:e.cor}}>{qtd}</span></div><div style={{background:C.border,borderRadius:4,height:6,overflow:"hidden"}}><div style={{height:"100%",width:`${pct}%`,background:e.cor,borderRadius:4,transition:"width 0.4s"}}/></div></div>);})}
          </div>}
        </>}
      </div>
      {leadSelecionado&&<ModalLead lead={leadSelecionado} onClose={()=>setLeadSelecionado(null)} onSave={salvarLead} onDelete={excluirLead}/>}
      {novoLead&&<ModalLead lead={leadVazio} onClose={()=>setNovoLead(false)} onSave={criarLead} onDelete={()=>setNovoLead(false)}/>}
      {importando&&<ModalImportar onClose={()=>setImportando(false)} onImportar={importarLeads} perfil={perfil} leadsExistentes={leads}/>}
      {alterandoSenha&&(
        <Modal onClose={()=>setAlterandoSenha(false)}>
          <ModalHeader title="Alterar senha" sub={perfil.email} onClose={()=>setAlterandoSenha(false)}/>
          <div style={{padding:24}}>
            <label style={lStyle}>Nova senha</label>
            <input type="password" value={novaSenha} onChange={e=>{setNovaSenha(e.target.value);setSenhaErro("");setSenhaMsg("");}} placeholder="mínimo 6 caracteres" style={{...iStyle,marginBottom:12}}/>
            {senhaErro&&<div style={{color:"#C0392B",fontSize:13,marginBottom:12,padding:"9px 12px",background:"#C0392B11",borderRadius:8}}>{senhaErro}</div>}
            {senhaMsg&&<div style={{color:"#4A9B6F",fontSize:13,marginBottom:12,padding:"9px 12px",background:"#4A9B6F11",borderRadius:8}}>✓ {senhaMsg}</div>}
            <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
              <button onClick={()=>setAlterandoSenha(false)} style={{padding:"9px 16px",border:`1px solid ${C.border}`,borderRadius:8,background:"transparent",color:C.muted,fontSize:13,cursor:"pointer"}}>Cancelar</button>
              <button onClick={handleAlterarSenha} disabled={senhaLoading} style={{padding:"9px 20px",border:"none",borderRadius:8,background:`linear-gradient(135deg,${C.bordo},${C.vinho})`,color:"#fff",fontSize:13,fontWeight:700,cursor:"pointer",opacity:senhaLoading?0.7:1}}>{senhaLoading?"Salvando...":"Alterar senha"}</button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────
export default function App(){
  const [sessao,setSessao]=useState(null);
  if(!sessao) return <Login onLogin={setSessao}/>;
  if(sessao.suspenso) return <TelaSuspensa perfil={sessao.perfil} token={sessao.token} onVoltar={()=>setSessao(null)}/>;
  if(sessao.admin) return <PainelAdmin token={sessao.token} onLogout={()=>setSessao(null)}/>;
  return <CRM perfil={sessao.perfil} token={sessao.token} onLogout={()=>setSessao(null)}/>;
}
