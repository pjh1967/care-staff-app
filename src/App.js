// ============================================================
//  요양보호사 근무 요청 앱 (직원용)
//  care-schedule-staff-app / src/App.js 로 사용
// ============================================================
import { useState, useEffect } from "react";

// ── Google API 설정 (관리자 앱과 동일하게 입력) ─────────────
const GAPI_CONFIG = {
  CLIENT_ID    : "575088677348-i8s653ni326sj9e7jpl6ikjrgbbrdfup.apps.googleusercontent.com",
  API_KEY      : "AIzaSyA17IdNlmepK2eK3riUzqH489BVJ-uGyww",
  SPREADSHEET_ID: "1xp3IJmB1jyrVY0DrDYdx2MXh4Xo68uRCTQkmh_xufhw",
  SCOPES       : "https://www.googleapis.com/auth/spreadsheets",
};
const SHEET_REQUEST = "요청입력";
const SHEET_CONFIG  = "설정";

// ── 색상 ─────────────────────────────────────────────────────
const C = {
  bg    : "#0a0f1e",
  panel : "#111827",
  dark  : "#0d1b2e",
  border: "#1e3a5f",
  teal  : "#00b4a6",
  amber : "#f59e0b",
  red   : "#ef4444",
  white : "#f0f4f8",
  gray  : "#64748b",
  steel : "#1e40af",
};

const TYPES     = ["주","야","공","V"];
const TYPE_LABEL= {주:"주간",야:"야간",공:"비번",V:"연차"};
const TYPE_COLOR= {주:"#F4B942",야:"#5B5EA6",공:"#A9D18E",V:"#FFD966"};
const TYPE_FG   = {주:"#111",야:"#fff",공:"#111",V:"#111"};
const WD        = ["일","월","화","수","목","금","토"];

// ── Sheets API ────────────────────────────────────────────────
const Sheets = {
  _ready:false, _token:null, _tokenExpiry:null, _tokenClient:null,

  async init(){
    return new Promise((res,rej)=>{
      if(Sheets._ready){res();return;}
      const s=document.createElement("script");
      s.src="https://apis.google.com/js/api.js";
      s.onload=()=>window.gapi.load("client",async()=>{
        try{
          await window.gapi.client.init({
            apiKey:GAPI_CONFIG.API_KEY,
            discoveryDocs:["https://sheets.googleapis.com/$discovery/rest?version=v4"],
          });
          Sheets._ready=true; res();
        }catch(e){rej(e);}
      });
      s.onerror=rej;
      document.head.appendChild(s);
    });
  },

  async signIn(){
    return new Promise((res,rej)=>{
      const s=document.createElement("script");
      s.src="https://accounts.google.com/gsi/client";
      s.onload=()=>{
        Sheets._tokenClient=window.google.accounts.oauth2.initTokenClient({
          client_id:GAPI_CONFIG.CLIENT_ID,
          scope:GAPI_CONFIG.SCOPES,
          callback:(r)=>{
            if(r.error){rej(r.error);return;}
            Sheets._token=r.access_token;
            Sheets._tokenExpiry=Date.now()+(r.expires_in||3600)*1000;
            window.gapi.client.setToken({access_token:r.access_token});
            res();
          },
        });
        Sheets._tokenClient.requestAccessToken({prompt:"consent"});
      };
      s.onerror=rej;
      document.head.appendChild(s);
    });
  },

  async read(range){
    const r=await window.gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId:GAPI_CONFIG.SPREADSHEET_ID, range,
    });
    return r.result.values||[];
  },

  async write(range,values){
    await window.gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId:GAPI_CONFIG.SPREADSHEET_ID,
      range, valueInputOption:"RAW",
      resource:{values},
    });
  },

  // 직원 목록 읽기
  async readStaff(){
    const rows=await this.read(`${SHEET_CONFIG}!A1:L60`);
    const staff=[];
    rows.forEach(r=>{
      if(r[0]==="STAFF"&&r[2]){
        staff.push({
          name:String(r[2]).trim(),
          role:String(r[3]).trim()||"요양보호사",
        });
      }
    });
    return staff;
  },

  // 기존 요청 읽기
  async readMyRequests(name){
    const rows=await this.read(`${SHEET_REQUEST}!A1:G200`);
    const reqs={};
    rows.forEach(r=>{
      if(String(r[0]).trim()===name&&r[4]){
        reqs[Number(r[4])]=String(r[5]).trim();
      }
    });
    return reqs;
  },

  // 요청 저장 (기존 삭제 후 새로 저장)
  async saveRequests(name,role,requests,year,month){
    // 전체 데이터 읽기
    const rows=await this.read(`${SHEET_REQUEST}!A1:G200`);
    const lastRow=Math.max(rows.length,4);

    // 기존 해당 직원 행 제거
    const others=rows.filter((r,i)=>i<4||String(r[0]).trim()!==name);

    // 새 요청 행 추가
    const newRows=Object.entries(requests)
      .sort(([a],[b])=>Number(a)-Number(b))
      .map(([d,t])=>[name,role,"여","0",Number(d),t,"1"]);

    // 헤더 유지 + 기타 직원 + 새 요청
    const allRows=[...others,...newRows];

    // 전체 다시 쓰기
    if(allRows.length>4){
      await this.write(`${SHEET_REQUEST}!A5:G${allRows.length+5}`,
        allRows.slice(4).map(r=>[r[0]||"",r[1]||"",r[2]||"",r[3]||"",r[4]||"",r[5]||"",r[6]||""]));
    }
  },
};

// ── 메인 앱 ──────────────────────────────────────────────────
export default function StaffApp(){
  const [step,setStep]     = useState("login");   // login|select|input|done
  const [staff,setStaff]   = useState([]);
  const [myName,setMyName] = useState("");
  const [myRole,setMyRole] = useState("");
  const [requests,setReqs] = useState({});
  const [year,setYear]     = useState(new Date().getFullYear());
  const [month,setMonth]   = useState(new Date().getMonth()+1);
  const [loading,setLoading]= useState(false);
  const [msg,setMsg]       = useState("");

  const total = new Date(year,month,0).getDate();

  // 로그인
  const login = async()=>{
    setLoading(true); setMsg("Google 로그인 중...");
    try{
      await Sheets.init();
      await Sheets.signIn();
      const s=await Sheets.readStaff();
      setStaff(s);
      setStep("select");
      setMsg("");
    }catch(e){
      setMsg("로그인 실패: "+e.message);
    }
    setLoading(false);
  };

  // 직원 선택
  const selectEmp = async(name,role)=>{
    setLoading(true); setMsg("기존 요청 불러오는 중...");
    setMyName(name); setMyRole(role);
    try{
      const prev=await Sheets.readMyRequests(name);
      setReqs(prev);
      setStep("input");
    }catch(e){ setReqs({}); setStep("input"); }
    setMsg(""); setLoading(false);
  };

  // 날짜 클릭 순환
  const toggle=(d)=>{
    setReqs(prev=>{
      const cur=prev[d];
      const idx=TYPES.indexOf(cur);
      const next={...prev};
      if(idx===-1) next[d]="주";
      else if(idx<TYPES.length-1) next[d]=TYPES[idx+1];
      else delete next[d];
      return next;
    });
  };

  // 저장
  const save=async()=>{
    setLoading(true); setMsg("저장 중...");
    try{
      await Sheets.saveRequests(myName,myRole,requests,year,month);
      setStep("done");
      setMsg("");
    }catch(e){ setMsg("저장 실패: "+e.message); }
    setLoading(false);
  };

  const summary={주:0,야:0,공:0,V:0};
  Object.values(requests).forEach(v=>{if(summary[v]!==undefined)summary[v]++;});

  const S={
    wrap:{minHeight:"100vh",background:C.bg,color:C.white,
          fontFamily:"'맑은 고딕',sans-serif",padding:16},
    card:{background:C.panel,borderRadius:16,padding:24,
          maxWidth:480,margin:"0 auto",border:`1px solid ${C.border}`},
    title:{fontSize:24,fontWeight:700,color:C.teal,marginBottom:8,textAlign:"center"},
    sub:  {fontSize:14,color:C.gray,textAlign:"center",marginBottom:24},
    btn:  (bg)=>({background:bg,color:"#fff",border:"none",borderRadius:10,
                  padding:"14px 24px",fontSize:16,fontWeight:700,
                  cursor:"pointer",width:"100%",marginTop:8}),
    msg:  {fontSize:13,color:C.amber,textAlign:"center",marginTop:8},
  };

  // ── 로그인 화면 ──
  if(step==="login") return(
    <div style={S.wrap}>
      <div style={S.card}>
        <div style={{fontSize:48,textAlign:"center",marginBottom:8}}>🏥</div>
        <div style={S.title}>요양보호사 근무 요청</div>
        <div style={S.sub}>근무 희망 일정을 입력해 주세요</div>
        <div style={{display:"flex",gap:8,marginBottom:16,justifyContent:"center"}}>
          <select value={year} onChange={e=>setYear(Number(e.target.value))}
            style={{background:C.dark,color:C.white,border:`1px solid ${C.border}`,
                    borderRadius:8,padding:"8px 12px",fontSize:16}}>
            {[2025,2026,2027].map(y=><option key={y}>{y}</option>)}
          </select>
          <select value={month} onChange={e=>setMonth(Number(e.target.value))}
            style={{background:C.dark,color:C.white,border:`1px solid ${C.border}`,
                    borderRadius:8,padding:"8px 12px",fontSize:16}}>
            {Array.from({length:12},(_,i)=>(
              <option key={i+1} value={i+1}>{i+1}월</option>
            ))}
          </select>
        </div>
        <button onClick={login} disabled={loading} style={S.btn(C.teal)}>
          {loading?"로그인 중...":"🔗 Google 로그인"}
        </button>
        {msg&&<div style={S.msg}>{msg}</div>}
      </div>
    </div>
  );

  // ── 직원 선택 화면 ──
  if(step==="select") return(
    <div style={S.wrap}>
      <div style={S.card}>
        <div style={S.title}>본인 이름 선택</div>
        <div style={S.sub}>{year}년 {month}월 근무 요청</div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {staff.map(s=>(
            <button key={s.name} onClick={()=>selectEmp(s.name,s.role)}
              style={{...S.btn(C.steel),textAlign:"left",
                      display:"flex",justifyContent:"space-between",
                      alignItems:"center"}}>
              <span>{s.name}</span>
              <span style={{fontSize:12,color:"#93c5fd"}}>{s.role}</span>
            </button>
          ))}
        </div>
        {msg&&<div style={S.msg}>{msg}</div>}
      </div>
    </div>
  );

  // ── 요청 입력 화면 (달력) ──
  if(step==="input") return(
    <div style={S.wrap}>
      <div style={{maxWidth:520,margin:"0 auto"}}>
        {/* 헤더 */}
        <div style={{background:C.panel,borderRadius:16,padding:16,
                     border:`1px solid ${C.border}`,marginBottom:12}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontSize:22,fontWeight:700,color:C.white}}>{myName}</div>
              <div style={{fontSize:13,color:C.gray}}>{myRole} · {year}년 {month}월</div>
            </div>
            <button onClick={()=>setStep("select")}
              style={{background:"transparent",color:C.gray,border:`1px solid ${C.border}`,
                      borderRadius:8,padding:"6px 12px",fontSize:13,cursor:"pointer"}}>
              ← 변경
            </button>
          </div>

          {/* 요약 */}
          <div style={{display:"flex",gap:6,marginTop:10,flexWrap:"wrap"}}>
            {TYPES.map(t=>(
              <span key={t} style={{background:TYPE_COLOR[t],color:TYPE_FG[t],
                                     borderRadius:6,padding:"3px 10px",
                                     fontSize:14,fontWeight:700}}>
                {TYPE_LABEL[t]} {summary[t]}일
              </span>
            ))}
          </div>
        </div>

        {/* 범례 */}
        <div style={{background:C.dark,borderRadius:10,padding:10,
                     border:`1px solid ${C.border}`,marginBottom:10,
                     fontSize:13,color:C.gray,textAlign:"center"}}>
          날짜 클릭 →&nbsp;
          {TYPES.map(t=>(
            <span key={t} style={{background:TYPE_COLOR[t],color:TYPE_FG[t],
                                   borderRadius:4,padding:"1px 8px",
                                   fontSize:13,fontWeight:700,margin:"0 2px"}}>
              {TYPE_LABEL[t]}
            </span>
          ))}
          &nbsp;→ 다시 클릭 시 삭제
        </div>

        {/* 달력 */}
        <div style={{background:C.panel,borderRadius:16,padding:14,
                     border:`1px solid ${C.border}`,marginBottom:12}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
            {WD.map(w=>(
              <div key={w} style={{textAlign:"center",fontSize:13,fontWeight:700,
                                    color:w==="일"?"#f87171":w==="토"?"#93c5fd":C.gray,
                                    padding:"4px 0",borderBottom:`1px solid ${C.border}`}}>
                {w}
              </div>
            ))}
            {Array.from({length:new Date(year,month-1,1).getDay()},(_,i)=>(
              <div key={`e${i}`}/>
            ))}
            {Array.from({length:total},(_,i)=>{
              const d=i+1;
              const wd=new Date(year,month-1,d).getDay();
              const req=requests[d];
              return(
                <div key={d} onClick={()=>toggle(d)}
                  style={{background:req?TYPE_COLOR[req]:wd===0?"#1a1a2e":"#1a2d4a",
                           color:req?TYPE_FG[req]:wd===0?"#6b7280":C.white,
                           borderRadius:8,padding:"8px 2px",textAlign:"center",
                           cursor:"pointer",border:`1px solid ${req?TYPE_COLOR[req]:C.border}`,
                           transition:"all 0.15s",userSelect:"none",minHeight:52}}>
                  <div style={{fontSize:16,fontWeight:700}}>{d}</div>
                  <div style={{fontSize:11,marginTop:2}}>
                    {req?TYPE_LABEL[req]:WD[wd]}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* 버튼 */}
        <button onClick={save} disabled={loading}
          style={{background:C.teal,color:"#fff",border:"none",borderRadius:12,
                  padding:"16px",fontSize:18,fontWeight:700,
                  cursor:"pointer",width:"100%",marginBottom:8}}>
          {loading?"저장 중...":"✅ 요청 저장"}
        </button>
        <button onClick={()=>setReqs({})}
          style={{background:"#5b1a1a",color:"#fca5a5",border:"none",borderRadius:12,
                  padding:"10px",fontSize:14,cursor:"pointer",width:"100%"}}>
          🗑 전체 초기화
        </button>
        {msg&&<div style={S.msg}>{msg}</div>}
      </div>
    </div>
  );

  // ── 완료 화면 ──
  return(
    <div style={S.wrap}>
      <div style={S.card}>
        <div style={{fontSize:64,textAlign:"center",marginBottom:16}}>✅</div>
        <div style={S.title}>요청 저장 완료!</div>
        <div style={S.sub}>
          {myName}님의 {year}년 {month}월 근무 요청이<br/>저장되었습니다.
        </div>
        <div style={{background:C.dark,borderRadius:10,padding:14,
                     border:`1px solid ${C.border}`,marginBottom:16}}>
          {TYPES.map(t=>summary[t]>0&&(
            <div key={t} style={{display:"flex",justifyContent:"space-between",
                                  padding:"4px 0",borderBottom:`1px solid ${C.border}`}}>
              <span style={{color:C.gray}}>{TYPE_LABEL[t]}</span>
              <span style={{color:TYPE_COLOR[t],fontWeight:700}}>{summary[t]}일</span>
            </div>
          ))}
        </div>
        <button onClick={()=>{setStep("select");setReqs({});}}
          style={S.btn(C.steel)}>
          다른 직원 요청 입력
        </button>
        <button onClick={()=>{setStep("login");setReqs({});setMyName("");}}
          style={S.btn(C.dark)}>
          처음으로
        </button>
      </div>
    </div>
  );
}