import { useState, useRef, useLayoutEffect } from "react";
import * as XLSX from "xlsx";

/* ── 폰트 주입 ───────────────────────────────────────────────────────────── */
(function inject() {
  if (document.querySelector("#bs7-link")) return;
  const lk = document.createElement("link");
  lk.id = "bs7-link"; lk.rel = "stylesheet";
  lk.href = "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css";
  document.head.appendChild(lk);
  // 나눔고딕 폰트 추가
  const nk = document.createElement("link");
  nk.id = "noto-sans-kr"; nk.rel = "stylesheet";
  nk.href = "https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap";
  document.head.appendChild(nk);
  const st = document.createElement("style");
  st.id = "bs7-style";
  st.textContent = `
    *{box-sizing:border-box;margin:0;padding:0;}
    body{font-family:'Pretendard',-apple-system,sans-serif!important;background:#0d1117;}
    ::-webkit-scrollbar{width:5px;}::-webkit-scrollbar-track{background:transparent;}
    ::-webkit-scrollbar-thumb{background:#30363d;border-radius:3px;}
    input[type=color]{-webkit-appearance:none;border:1px solid #30363d;border-radius:4px;overflow:hidden;cursor:pointer;}
    input[type=color]::-webkit-color-swatch-wrapper{padding:0;}input[type=color]::-webkit-color-swatch{border:none;}
    input[type=range]{-webkit-appearance:none;appearance:none;height:3px;border-radius:2px;background:#30363d;outline:none;cursor:pointer;}
    input[type=range]::-webkit-slider-thumb{-webkit-appearance:none;width:13px;height:13px;border-radius:50%;background:#58a6ff;cursor:pointer;}
    select,button{outline:none;font-family:'Pretendard',-apple-system,sans-serif;}
  `;
  document.head.appendChild(st);
})();

/* jsPDF 제거 - HTML 방식으로 대체 */

/* ── 사이즈 ──────────────────────────────────────────────────────────────── */
const MM = 11.811;
const mkSz = (label, sub, note, wMM, hMM, pw) => {
  const ph = Math.round(pw * hMM / wMM);
  return { label, sub, note, wMM, hMM, w:pw, h:ph, ew:Math.round(wMM*MM), eh:Math.round(hMM*MM) };
};
const SIZES = [
  mkSz("가로형",    "117×108mm","공공기관·단체모임",             117,108,390),
  mkSz("세로형 L",  "110×172mm","공공기관·세미나·학교·단체모임", 110,172,320),
  mkSz("세로형 M",  "110×152mm","공공기관·세미나·학교·단체모임", 110,152,320),
  mkSz("세로형 S",  "95×147mm", "공공기관·세미나·학교·단체모임",  95,147,280),
  mkSz("집게 가로", "107×75mm", "속지 96×66mm",                 107, 75,370),
  mkSz("집게 세로", "73×107mm", "속지 66×96mm",                  73,107,240),
];

/* 커스텀 사이즈 생성 (mm → 캔버스px, 출력px) */
const mkCustomSz = (wMM, hMM) => {
  const pw = Math.min(400, Math.round(300 * wMM / Math.max(wMM, hMM)));
  const ph = Math.round(pw * hMM / wMM);
  return { label:"직접입력", sub:`${wMM}×${hMM}mm`, note:"사용자 정의", wMM, hMM, w:pw, h:ph, ew:Math.round(wMM*MM), eh:Math.round(hMM*MM) };
};
const CUSTOM_IDX = -1; /* 커스텀 사이즈 식별자 */

const A4 = { w:210, h:297, margin:8, gap:4 };
const calcLayout = (sz) => {
  const cols = Math.max(1, Math.floor((A4.w - A4.margin*2 + A4.gap) / (sz.wMM + A4.gap)));
  const rows = Math.max(1, Math.floor((A4.h - A4.margin*2 + A4.gap) / (sz.hMM + A4.gap)));
  return { cols, rows, perPage: cols * rows };
};

const C = {
  bg:"#0d1117", surface:"#161b22", border:"#21262d", border2:"#30363d",
  text:"#c9d1d9", muted:"#8b949e", dim:"#484f58",
  blue:"#58a6ff", green:"#3fb950", accent:"#1f6feb", purple:"#8b5cf6", orange:"#f0883e",
};
const mkZone = (ex={}) => ({ visible:true, fontSize:18, ox:0, oy:0, ...ex });

/* ── 핵심 수정: FileReader로 base64 변환 ─────────────────────────────────
   blob URL(URL.createObjectURL)은 샌드박스 iframe에서 canvas.drawImage가
   보안 정책으로 차단됨. base64 data URL은 어디서나 동작함.
────────────────────────────────────────────────────────────────────────── */
function fileToImage(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error("이미지 변환 실패"));
      img.src = e.target.result;   // ← base64 data URL (blob URL 아님)
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsDataURL(file);    // ← FileReader로 base64 변환
  });
}


/* ── 샘플 CSV 안내 박스 (복사 방식) ─────────────────────────────────────── */
const SAMPLE_CSV = "소속,직위,이름\n중소기업중앙회,팀장,홍길동\n○○기업,대표이사,김철수\n△△협회,사무국장,이영희";

function SampleCsvBox() {
  const [copied, setCopied] = useState(false);
  const copy = () => {
    navigator.clipboard.writeText(SAMPLE_CSV).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }).catch(() => {
      // clipboard API 실패 시 textarea 선택
      const ta = document.getElementById("sample-csv-ta");
      if (ta) { ta.select(); document.execCommand("copy"); setCopied(true); setTimeout(()=>setCopied(false),2000); }
    });
  };
  return (
    <div style={{background:"#0a0e14",border:`1px solid ${C.border2}`,borderRadius:6,padding:10,marginTop:4}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
        <span style={{fontSize:10,color:C.muted,fontWeight:600}}>📋 샘플 형식 (복사 후 메모장에 붙여넣어 .csv로 저장)</span>
        <button onClick={copy} style={{fontSize:10,padding:"3px 8px",background:copied?"#238636":C.accent,color:"#fff",border:"none",borderRadius:4,cursor:"pointer",flexShrink:0}}>
          {copied ? "✓ 복사됨" : "복사"}
        </button>
      </div>
      <textarea id="sample-csv-ta" readOnly value={SAMPLE_CSV}
        style={{width:"100%",height:72,background:"#060a0f",border:`1px solid ${C.border}`,borderRadius:4,color:"#7ee787",fontSize:11,fontFamily:"monospace",padding:"6px 8px",resize:"none",outline:"none"}}
      />
    </div>
  );
}


/* ── 출력 미리보기 모달 ─────────────────────────────────────────────────── */
function PdfPreviewModal({ preview, onClose }) {
  if (!preview) return null;
  const iframeRef = useRef(null);
  const isHtml = preview.type === "html";

  const handlePrint = () => {
    if (iframeRef.current && iframeRef.current.contentWindow) {
      iframeRef.current.contentWindow.focus();
      iframeRef.current.contentWindow.print();
    }
  };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:9999,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:12,padding:16}}>
      {/* 헤더 */}
      <div style={{display:"flex",alignItems:"center",gap:8,width:"100%",maxWidth:860}}>
        <div style={{flex:1}}>
          <div style={{fontSize:12,color:C.blue,fontWeight:700}}>
            {isHtml ? "📄 A4 명찰 출력 미리보기" : "🖼 PNG 미리보기"}
          </div>
          <div style={{fontSize:10,color:C.muted,marginTop:2}}>{preview.name}</div>
        </div>
        {isHtml && (
          <button onClick={handlePrint}
            style={{padding:"7px 16px",background:"#238636",color:"#fff",border:"none",borderRadius:6,fontSize:12,fontWeight:700,cursor:"pointer",flexShrink:0}}>
            🖨 인쇄 / PDF 저장
          </button>
        )}
        <button onClick={onClose}
          style={{padding:"7px 14px",background:C.border2,color:C.text,border:"none",borderRadius:6,fontSize:12,cursor:"pointer",flexShrink:0}}>
          ✕ 닫기
        </button>
      </div>

      {/* 미리보기 본문 */}
      <div style={{width:"100%",maxWidth:860,flex:1,minHeight:0,borderRadius:8,overflow:"hidden",background:"#fff",boxShadow:"0 8px 40px rgba(0,0,0,0.6)"}}>
        {isHtml
          ? <iframe ref={iframeRef} srcDoc={preview.html}
              style={{width:"100%",height:"100%",minHeight:"60vh",border:"none",display:"block"}}
              title="명찰 미리보기"/>
          : <img src={preview.url} alt="badge"
              style={{maxWidth:"100%",maxHeight:"70vh",display:"block",margin:"0 auto",padding:16}}/>
        }
      </div>

      {isHtml && (
        <div style={{fontSize:10,color:C.muted,textAlign:"center",lineHeight:1.8}}>
          <strong style={{color:C.text}}>🖨 인쇄 / PDF 저장</strong> 버튼 클릭 →
          프린터 선택 또는 <strong style={{color:C.text}}>PDF로 저장</strong> 선택<br/>
          인쇄 설정에서 <strong style={{color:C.text}}>여백: 없음, 배율: 100%</strong>로 설정하면 정확한 크기로 출력됩니다.
        </div>
      )}
    </div>
  );
}

/* ── 렌더러 ──────────────────────────────────────────────────────────────── */
function wrapText(ctx, text, cx, cy, maxW, lh) {
  if (!text) return;
  let line="", lines=[];
  for (const ch of String(text)) {
    const t = line+ch;
    if (ctx.measureText(t).width > maxW && line) { lines.push(line); line=ch; } else line=t;
  }
  if (line) lines.push(line);
  const sy = cy - ((lines.length-1)*lh)/2;
  lines.forEach((l,i) => ctx.fillText(l, cx, sy+i*lh));
}

function renderBadge(canvas, { baseW, baseH, bgImg, logoImg, border, eventName, textColor, zones, fontFamily, topHeightRatio }, data={}) {
  const ctx = canvas.getContext("2d");
  const W=canvas.width, H=canvas.height;
  const s=Math.min(W/baseW, H/baseH);
  ctx.clearRect(0,0,W,H);

  /* 배경 */
  if (bgImg) {
    const sc=Math.max(W/bgImg.width, H/bgImg.height);
    ctx.drawImage(bgImg, (W-bgImg.width*sc)/2, (H-bgImg.height*sc)/2, bgImg.width*sc, bgImg.height*sc);
  } else {
    const g=ctx.createLinearGradient(0,0,0,H);
    g.addColorStop(0,"#e8edf5"); g.addColorStop(1,"#d0d8e8");
    ctx.fillStyle=g; ctx.fillRect(0,0,W,H);
    ctx.fillStyle="#9fadbf"; ctx.font=`${11*s}px ${fontFamily},sans-serif`;
    ctx.textAlign="center"; ctx.textBaseline="middle";
    ctx.fillText("배경 이미지를 업로드하세요", W/2, H/2);
  }

  /* 테두리 */
  const bt=border.thickness*s;
  if (bt>0) { ctx.strokeStyle=border.color; ctx.lineWidth=bt; ctx.strokeRect(bt/2,bt/2,W-bt,H-bt); }

  const pad=W*0.1, topH=H*topHeightRatio, botH=H*0.22, midH=H-topH-botH;
  const z=zones;

  /* 행사명 */
  if (z.top.visible && eventName) {
    const fs=z.top.fontSize*s;
    ctx.fillStyle=textColor; ctx.font=`700 ${fs}px ${fontFamily},'Noto Sans KR',sans-serif`;
    ctx.textAlign="center"; ctx.textBaseline="middle";
    wrapText(ctx, eventName, W/2+z.top.ox*s, topH/2+z.top.oy*s, W-pad*0.4, fs*1.4);
  }

  /* 소속·직위·이름·필드1·필드2 */
  const midDefs=[
    {key:"org",val:data.org,weight:"500"},
    {key:"position",val:data.position,weight:"400"},
    {key:"name",val:data.name,weight:"700"},
    {key:"field1",val:data.field1,weight:"400"},
    {key:"field2",val:data.field2,weight:"400"},
  ].filter(d=>z[d.key]?.visible && d.val);
  if (midDefs.length) {
    const sp=midH/(midDefs.length+1);
    midDefs.forEach((d,i)=>{
      const zz=z[d.key], fs=zz.fontSize*s;
      ctx.fillStyle=textColor; ctx.font=`${d.weight} ${fs}px ${fontFamily},'Noto Sans KR',sans-serif`;
      ctx.textAlign="center"; ctx.textBaseline="middle";
      ctx.fillText(d.val, W/2+zz.ox*s, topH+sp*(i+1)+zz.oy*s, W-pad*2);
    });
  }

  /* 로고 */
  const botTop=topH+midH, lz=z.logo;
  if (lz.visible && logoImg) {
    const maxW=W*0.55, maxH=botH*0.7;
    const sc=Math.min(maxW/logoImg.width, maxH/logoImg.height);
    const lw=logoImg.width*sc, lh=logoImg.height*sc;
    ctx.drawImage(logoImg, (W-lw)/2+lz.ox*s, botTop+(botH-lh)/2+lz.oy*s, lw, lh);
  }
}

/* ── 엣지 색상 감지 ──────────────────────────────────────────────────────── */
function detectEdgeColor(img) {
  const c=document.createElement("canvas"); c.width=60; c.height=60;
  const ctx=c.getContext("2d"); ctx.drawImage(img,0,0,60,60);
  const pts=[];
  for(let x=0;x<60;x+=4){pts.push([...ctx.getImageData(x,0,1,1).data.slice(0,3)]);pts.push([...ctx.getImageData(x,59,1,1).data.slice(0,3)]);}
  for(let y=4;y<56;y+=4){pts.push([...ctx.getImageData(0,y,1,1).data.slice(0,3)]);pts.push([...ctx.getImageData(59,y,1,1).data.slice(0,3)]);}
  const avg=pts.reduce((a,p)=>[a[0]+p[0],a[1]+p[1],a[2]+p[2]],[0,0,0]).map(v=>Math.round(Math.min(220,Math.max(0,v/pts.length*0.6))));
  return "#"+avg.map(v=>v.toString(16).padStart(2,"0")).join("");
}

/* ── UI 원자 ─────────────────────────────────────────────────────────────── */
const baseInp={background:"#0d1117",border:"1px solid #30363d",borderRadius:4,color:"#c9d1d9",padding:"6px 8px",fontSize:12,width:"100%",outline:"none"};
function Sec({title,children,accent}){return(<div style={{marginBottom:16}}><div style={{fontSize:9.5,fontWeight:700,color:accent||C.muted,textTransform:"uppercase",letterSpacing:"0.13em",marginBottom:7,paddingBottom:5,borderBottom:`1px solid ${C.border}`}}>{title}</div>{children}</div>);}
function Row({label,children}){return(<div style={{display:"flex",alignItems:"center",gap:6,marginBottom:7}}>{label&&<span style={{fontSize:11,color:C.muted,flexShrink:0,width:60}}>{label}</span>}<div style={{flex:1,display:"flex",alignItems:"center",gap:5,justifyContent:"flex-end"}}>{children}</div></div>);}
function Toggle({on,onChange}){return(<div onClick={()=>onChange(!on)} style={{width:32,height:17,background:on?"#238636":C.border2,borderRadius:9,cursor:"pointer",position:"relative",flexShrink:0,transition:"background .2s"}}><div style={{width:13,height:13,background:"#fff",borderRadius:"50%",position:"absolute",top:2,left:on?17:2,transition:"left .2s"}}/></div>);}
function Slider({min=0,max=100,value,onChange,unit=""}){return(<div style={{display:"flex",alignItems:"center",gap:6,flex:1}}><input type="range" min={min} max={max} value={value} onChange={e=>onChange(+e.target.value)} style={{flex:1}}/><span style={{fontSize:11,color:C.muted,minWidth:28,textAlign:"right"}}>{value}{unit}</span></div>);}

function ImgUploadBtn({ id, label, accept, onChange, hasImage, loading }) {
  return (
    <label htmlFor={id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 11px",background:C.bg,border:`1px dashed ${loading?C.orange:hasImage?"#238636":C.border2}`,borderRadius:6,cursor:"pointer",marginBottom:6,userSelect:"none"}}>
      <span style={{fontSize:15}}>{loading?"⏳":hasImage?"✓":"↑"}</span>
      <span style={{fontSize:12,color:loading?C.orange:hasImage?C.green:C.muted}}>
        {loading?"변환 중...":hasImage?label+" (재업로드 가능)":label}
      </span>
      <input id={id} type="file" accept={accept} style={{display:"none"}}
        onChange={e=>{ const f=e.target.files[0]; e.target.value=""; if(f) onChange(f); }}/>
    </label>
  );
}

/* ── A4 미리보기 ──────────────────────────────────────────────────────────── */
function A4Preview({sz,total}){
  const {cols,rows,perPage}=calcLayout(sz);
  const pages=total>0?Math.ceil(total/perPage):1;
  const sW=120,sH=170,mPx=(A4.margin/A4.w)*sW,gPx=(A4.gap/A4.w)*sW;
  const bW=(sz.wMM/A4.w)*sW,bH=(sz.hMM/A4.h)*sH;
  const cells=[];
  for(let r=0;r<rows;r++) for(let c=0;c<cols;c++) cells.push({x:mPx+c*(bW+gPx),y:mPx+r*(bH+gPx),idx:r*cols+c});
  return(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:8}}>
      <svg width={sW} height={sH} style={{background:"#fff",borderRadius:2,boxShadow:"0 2px 8px rgba(0,0,0,.4)"}}>
        {cells.map((c,i)=><rect key={i} x={c.x} y={c.y} width={bW} height={bH} fill={total>0&&c.idx<(total%perPage||perPage)?"#bfdbfe":"#e2e8f0"} stroke="#94a3b8" strokeWidth="0.5" rx="1"/>)}
        {cells.map((c,i)=><text key={`t${i}`} x={c.x+bW/2} y={c.y+bH/2} textAnchor="middle" dominantBaseline="middle" fontSize="5" fill="#1e40af" fontFamily="sans-serif" fontWeight="600">{c.idx+1}</text>)}
      </svg>
      <div style={{fontSize:10,color:C.muted,textAlign:"center",lineHeight:1.6}}>
        <span style={{color:C.text,fontWeight:700}}>{cols}열×{rows}행</span> = {perPage}매/페이지
        {total>0&&<><br/>총 <span style={{color:C.green,fontWeight:700}}>{total}명</span> → <span style={{color:C.orange,fontWeight:700}}>{pages}p</span> PDF</>}
      </div>
    </div>
  );
}

/* ── 위치 / ZoneCard ──────────────────────────────────────────────────────── */
function PosControl({ox,oy,onChange,range=120}){
  const changed=ox!==0||oy!==0;
  return(
    <div style={{background:"#0a0e14",border:`1px solid ${changed?C.purple:C.border}`,borderRadius:6,padding:"9px 10px"}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
        <span style={{fontSize:10,color:changed?C.purple:C.muted,fontWeight:700}}>📍 위치 조정</span>
        {changed&&<button onClick={()=>onChange(0,0)} style={{fontSize:9,padding:"2px 7px",background:"#3d1b7a",color:"#c4b5fd",border:"none",borderRadius:10,cursor:"pointer"}}>초기화</button>}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:7}}>
        <span style={{fontSize:10,color:C.dim,width:24,textAlign:"center"}}>↔</span>
        <input type="range" min={-range} max={range} value={ox} onChange={e=>onChange(+e.target.value,oy)} style={{flex:1}}/>
        <span style={{fontSize:10,color:ox!==0?C.purple:C.muted,width:32,textAlign:"right",fontFamily:"monospace"}}>{ox>0?"+":""}{ox}</span>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <span style={{fontSize:10,color:C.dim,width:24,textAlign:"center"}}>↕</span>
        <input type="range" min={-range} max={range} value={oy} onChange={e=>onChange(ox,+e.target.value)} style={{flex:1}}/>
        <span style={{fontSize:10,color:oy!==0?C.purple:C.muted,width:32,textAlign:"right",fontFamily:"monospace"}}>{oy>0?"+":""}{oy}</span>
      </div>
      {changed&&<div style={{marginTop:6,fontSize:9,color:"#7c5cbf",textAlign:"right",fontFamily:"monospace"}}>({ox>0?"+":""}{ox}, {oy>0?"+":""}{oy})</div>}
    </div>
  );
}
function ZoneCard({title,zone,onChange,showFontSize=true,fsRange=[8,72]}){
  const [open,setOpen]=useState(false);
  const changed=zone.ox!==0||zone.oy!==0;
  return(
    <div style={{background:C.bg,border:`1px solid ${open?C.border2:C.border}`,borderRadius:6,marginBottom:6,overflow:"hidden"}}>
      <div style={{display:"flex",alignItems:"center",padding:"9px 10px",cursor:"pointer",gap:8}} onClick={()=>setOpen(v=>!v)}>
        <span style={{fontSize:11,color:C.text,fontWeight:600,flex:1}}>{title}</span>
        {changed&&<span style={{fontSize:9,padding:"2px 6px",background:"#3d1b7a",color:"#c4b5fd",borderRadius:10}}>위치조정됨</span>}
        <Toggle on={zone.visible} onChange={v=>onChange({...zone,visible:v})}/>
        <span style={{fontSize:10,color:C.dim}}>{open?"▲":"▼"}</span>
      </div>
      {open&&(
        <div style={{padding:"0 10px 10px",borderTop:`1px solid ${C.border}`}}>
          {showFontSize&&(
            <div style={{display:"flex",alignItems:"center",gap:6,marginTop:9,marginBottom:8}}>
              <span style={{fontSize:11,color:C.muted,width:60,flexShrink:0}}>글자 크기</span>
              <Slider min={fsRange[0]} max={fsRange[1]} value={zone.fontSize} onChange={v=>onChange({...zone,fontSize:v})}/>
            </div>
          )}
          <PosControl ox={zone.ox} oy={zone.oy} onChange={(x,y)=>onChange({...zone,ox:x,oy:y})}/>
        </div>
      )}
    </div>
  );
}

/* ════════════════════════════════════════════════════════════════════════════
   메인 컴포넌트
════════════════════════════════════════════════════════════════════════════ */
export default function BadgePrintSystem() {
  const [szIdx,       setSzIdx]       = useState(1);      // -1 = 커스텀
  const [customW,     setCustomW]     = useState("100");  // mm (문자열로 저장 → 입력 중 자유롭게)
  const [customH,     setCustomH]     = useState("150");  // mm
  const [bgImg,       setBgImg]       = useState(null);
  const [logoImg,     setLogoImg]     = useState(null);
  const [bgLoading,   setBgLoading]   = useState(false);
  const [logoLoading, setLogoLoading] = useState(false);
  const [border,      setBorder]      = useState({color:"#1e3a8a",thickness:4});
  const [autoColor,   setAutoColor]   = useState(true);
  const [eventName,   setEventName]   = useState("2026 중소기업 CEO 포럼");
  const [textColor,   setTextColor]   = useState("#1a1a2e");
  const [fontFamily,  setFontFamily]  = useState("Pretendard");
  const [zones,       setZones]       = useState({
    top:      mkZone({fontSize:22}),
    org:      mkZone({fontSize:15}),
    position: mkZone({fontSize:13}),
    name:     mkZone({fontSize:28}),
    field1:   {visible:false,fontSize:12,ox:0,oy:0},
    field2:   {visible:false,fontSize:12,ox:0,oy:0},
    logo:     {visible:true,fontSize:14,ox:0,oy:0},
  });
  const [topHeightRatio, setTopHeightRatio] = useState(0.22);
  const [csvData,  setCsvData]  = useState([]);
  const [headers,  setHeaders]  = useState([]);
  const [colMap,   setColMap]   = useState({org:-1,position:-1,name:-1,field1:-1,field2:-1});
  const [pidx,     setPidx]     = useState(0);
  const [tab,      setTab]      = useState("design");
  const [exporting,setExporting]= useState(false);
  const [pdfPreview,setPdfPreview]= useState(null); // {type, url, name}
  const [expProg,  setExpProg]  = useState({cur:0,total:0});
  const [expMsg,   setExpMsg]   = useState("");

  const canvasRef = useRef(null);
  const safeW = Math.min(400, Math.max(10, parseFloat(customW) || 10));
  const safeH = Math.min(400, Math.max(10, parseFloat(customH) || 10));
  const sz = szIdx === CUSTOM_IDX
    ? mkCustomSz(safeW, safeH)
    : SIZES[szIdx];

  /* ── 미리보기 데이터 ──────────────────────────────────────────────────── */
  const getRowData = (row) => ({
    org:      colMap.org>=0      && row ? String(row[colMap.org]     ??"") : "",
    position: colMap.position>=0 && row ? String(row[colMap.position]??"") : "",
    name:     colMap.name>=0     && row ? String(row[colMap.name]    ??"") : "",
    field1:   colMap.field1>=0   && row ? String(row[colMap.field1]  ??"") : "",
    field2:   colMap.field2>=0   && row ? String(row[colMap.field2]  ??"") : "",
  });
  const previewData = csvData.length>0 ? getRowData(csvData[pidx])
    : {org:"중소기업중앙회", position:"팀장", name:"홍 길 동"};

  /* ── useLayoutEffect: paint 전에 동기적으로 캔버스 갱신 ─────────────── */
  useLayoutEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    // 크기가 바뀔 때만 재설정 (같은 값 대입해도 캔버스가 지워지므로 체크)
    if (canvas.width !== sz.w)  canvas.width  = sz.w;
    if (canvas.height !== sz.h) canvas.height = sz.h;
    renderBadge(canvas, {baseW:sz.w, baseH:sz.h, bgImg, logoImg, border, eventName, textColor, zones, fontFamily, topHeightRatio}, previewData);
  }); // 의존성 없음 → 모든 상태 변경에 반응

  /* ── 이미지 업로드 (FileReader base64) ───────────────────────────────── */
  const handleBgFile = (file) => {
    setBgLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        setBgImg(img);           // state 저장 → 리렌더 → useLayoutEffect
        setBgLoading(false);
        if (autoColor) setBorder(p => ({...p, color:detectEdgeColor(img)}));
      };
      img.onerror = () => { setBgLoading(false); alert("이미지 로드 실패"); };
      img.src = e.target.result; // base64 data URL
    };
    reader.onerror = () => { setBgLoading(false); alert("파일 읽기 실패"); };
    reader.readAsDataURL(file);  // blob URL 대신 base64
  };

  const handleLogoFile = (file) => {
    setLogoLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => { setLogoImg(img); setLogoLoading(false); };
      img.onerror = () => { setLogoLoading(false); alert("로고 로드 실패"); };
      img.src = e.target.result;
    };
    reader.onerror = () => { setLogoLoading(false); alert("파일 읽기 실패"); };
    reader.readAsDataURL(file);
  };

  /* ── CSV ─────────────────────────────────────────────────────────────── */
  const handleCsv = (e) => {
    const f=e.target.files[0]; if(!f) return;
    const r=new FileReader();
    r.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(!raw.length) return;
        const hdrs=raw[0].map(String);
        const rows=raw.slice(1).filter(r=>r.some(v=>String(v).trim()));
        setHeaders(hdrs); setCsvData(rows); setPidx(0);
        const find=kws=>hdrs.findIndex(h=>kws.some(k=>h.includes(k)));
        setColMap({org:find(["소속","업체","기관","회사"]),position:find(["직위","직급","직책"]),name:find(["이름","성명","성함"]),field1:-1,field2:-1});
      }catch(err){alert("파일 파싱 오류: "+err.message);}
    };
    r.readAsBinaryString(f);
  };

  /* ── 단일 PNG 출력 - base64 data URL 방식 ──────────────────────────── */
  const exportOne = () => {
    const c=document.createElement("canvas"); c.width=sz.ew; c.height=sz.eh;
      renderBadge(c,{baseW:sz.w,baseH:sz.h,bgImg,logoImg,border,eventName,textColor,zones,fontFamily,topHeightRatio},previewData);
    const dataUrl = c.toDataURL("image/png");
    setPdfPreview({ type:"image", url:dataUrl, name:`badge_${previewData.name||"preview"}.png` });
  };

  /* ── A4 대량 출력 (HTML → srcdoc iframe → 인쇄/PDF저장) ─────────────
     샌드박스 환경 제약:
       - blob URL, PDF data URI → iframe 표시 불가
       - srcdoc 속성으로 HTML 직접 삽입 → 정상 동작
       - iframe.contentWindow.print() → 브라우저 인쇄 다이얼로그 → PDF 저장 가능
  ───────────────────────────────────────────────────────────────────── */
  const exportPDF = async () => {
    if (!csvData.length) return;
    setExporting(true);

    const { cols, rows, perPage } = calcLayout(sz);
    const total = csvData.length;
    const cfg = { baseW:sz.w, baseH:sz.h, bgImg, logoImg, border, eventName, textColor, zones };

    // ① 모든 명찰을 JPEG base64로 렌더링 (150dpi 저해상도)
    const images = [];
    const printScale = 0.5; // 150dpi (300dpi 기준 50%)
    const printW = Math.round(sz.ew * printScale);
    const printH = Math.round(sz.eh * printScale);
    
    for (let i = 0; i < total; i++) {
      setExpProg({ cur:i+1, total });
      setExpMsg(`명찰 렌더링 중... (${i+1}/${total})`);
      const cv = document.createElement("canvas");
      cv.width = printW; cv.height = printH;
      renderBadge(cv, {...cfg, fontFamily, topHeightRatio}, getRowData(csvData[i]));
      images.push(cv.toDataURL("image/jpeg", 0.50));
      if (i % 2 === 1) await new Promise(r => setTimeout(r, 0));
    }

    setExpMsg("HTML 페이지 조합 중...");
    await new Promise(r => setTimeout(r, 0));

    // ② A4 페이지 HTML 생성
    // 각 명찰 크기 (mm 단위)
    const bW = sz.wMM, bH = sz.hMM;
    const margin = A4.margin, gap = A4.gap;

    // 페이지 묶기
    const pages = [];
    for (let p = 0; p < Math.ceil(total / perPage); p++) {
      pages.push(images.slice(p * perPage, (p + 1) * perPage));
    }

    let pageHtml = "";
    for (let pIdx = 0; pIdx < pages.length; pIdx++) {
      const pageImgs = pages[pIdx];
      setExpProg({ cur: Math.min(total + pIdx, total), total: total + pages.length });
      setExpMsg(`PDF 페이지 생성 중... (${pIdx+1}/${pages.length})`);
      const cells = pageImgs.map((src, idx) => {
        const col = idx % cols;
        const row = Math.floor(idx / cols);
        const x = margin + col * (bW + gap);
        const y = margin + row * (bH + gap);
        return `<img src="${src}" style="position:absolute;left:${x}mm;top:${y}mm;width:${bW}mm;height:${bH}mm;display:block;" />`;
      }).join("");
      pageHtml += `<div class="page">${cells}</div>`;
      if (pIdx % 2 === 1) await new Promise(r => setTimeout(r, 0));
    }

    const html = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>명찰 출력 — ${eventName||"badge"}</title>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { background:#888; }
  .page {
    position: relative;
    width: 210mm;
    height: 297mm;
    background: #fff;
    margin: 0 auto 8mm;
    overflow: hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.3);
  }
  @media print {
    body { background: none; }
    .page {
      margin: 0;
      box-shadow: none;
      page-break-after: always;
    }
    .page:last-child { page-break-after: avoid; }
  }
</style>
</head>
<body>${pageHtml}</body>
</html>`;

    setPdfPreview({
      type: "html",
      html,
      name: `명찰_${eventName||"badge"}_${total}명`,
    });
    setExporting(false); setExpMsg(""); setExpProg({ cur:0, total:0 });
  };

  /* ── 레이아웃 계산 ───────────────────────────────────────────────────── */
  const layout=calcLayout(sz);
  const dScale=Math.min(1, 500/sz.h, 360/sz.w);
  const dW=Math.round(sz.w*dScale), dH=Math.round(sz.h*dScale);
  const upZ=(k,v)=>setZones(p=>({...p,[k]:v}));

  /* ── 렌더 ────────────────────────────────────────────────────────────── */
  return (
    <>
    <PdfPreviewModal preview={pdfPreview} onClose={()=>setPdfPreview(null)}/>
    <div style={{display:"flex",height:"100vh",fontFamily:"Pretendard,sans-serif",background:C.bg,color:C.text,overflow:"hidden"}}>

      {/* 사이드바 */}
      <div style={{width:305,background:C.surface,borderRight:`1px solid ${C.border}`,display:"flex",flexDirection:"column",flexShrink:0}}>
        <div style={{padding:"15px 16px 11px",borderBottom:`1px solid ${C.border}`}}>
          <div style={{fontSize:9.5,fontWeight:700,letterSpacing:"0.16em",color:C.blue,textTransform:"uppercase"}}>Badge Print System</div>
          <div style={{fontSize:16,fontWeight:700,color:"#f0f6fc",marginTop:2}}>명찰 출력 시스템</div>
        </div>
        <div style={{display:"flex",borderBottom:`1px solid ${C.border}`}}>
          {[["design","🎨 디자인"],["data","📋 데이터"],["export","📄 출력"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"9px 0",fontSize:11,fontWeight:600,border:"none",background:tab===k?C.bg:"transparent",color:tab===k?C.blue:C.muted,borderBottom:`2px solid ${tab===k?C.blue:"transparent"}`,cursor:"pointer"}}>{l}</button>
          ))}
        </div>

        <div style={{flex:1,overflowY:"auto",padding:14}}>

          {/* 디자인 탭 */}
          {tab==="design"&&(<>
            <Sec title="명찰 사이즈">
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {SIZES.map((s,i)=>(
                  <button key={i} onClick={()=>setSzIdx(i)} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:szIdx===i?"#0d2137":C.bg,border:`1px solid ${szIdx===i?"#388bfd":C.border}`,borderRadius:5,cursor:"pointer",textAlign:"left"}}>
                    <div style={{width:6,height:6,borderRadius:"50%",background:szIdx===i?C.blue:C.dim,flexShrink:0}}/>
                    <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:szIdx===i?"#fff":C.text}}>{s.label}</div><div style={{fontSize:10,color:C.muted}}>{s.sub}</div></div>
                    <div style={{fontSize:9,color:C.dim,textAlign:"right",maxWidth:110,lineHeight:1.4}}>{s.note}</div>
                  </button>
                ))}
                {/* 직접 입력 버튼 */}
                <button onClick={()=>setSzIdx(CUSTOM_IDX)} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:szIdx===CUSTOM_IDX?"#1a0a2e":C.bg,border:`1px solid ${szIdx===CUSTOM_IDX?C.purple:C.border}`,borderRadius:5,cursor:"pointer",textAlign:"left"}}>
                  <div style={{width:6,height:6,borderRadius:"50%",background:szIdx===CUSTOM_IDX?C.purple:C.dim,flexShrink:0}}/>
                  <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:szIdx===CUSTOM_IDX?"#c4b5fd":C.text}}>직접 입력</div><div style={{fontSize:10,color:C.muted}}>가로×세로 직접 지정</div></div>
                  <div style={{fontSize:9,color:C.dim}}>Custom</div>
                </button>
              </div>
              {/* 커스텀 사이즈 입력 영역 */}
              {szIdx===CUSTOM_IDX&&(
                <div style={{marginTop:8,background:"#0f0620",border:`1px solid ${C.purple}`,borderRadius:7,padding:"12px 12px 10px"}}>
                  <div style={{fontSize:10,color:"#c4b5fd",fontWeight:700,marginBottom:10,letterSpacing:"0.05em"}}>✏️ 사이즈 직접 입력 (mm)</div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{flex:1}}>
                      <div style={{fontSize:10,color:C.muted,marginBottom:4}}>가로 (width)</div>
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <input type="number" min="10" max="400" value={customW}
                          onChange={e=>setCustomW(e.target.value)}
                          onBlur={e=>{const v=Math.min(400,Math.max(10,parseFloat(e.target.value)||10));setCustomW(String(v));}}
                          style={{...{background:"#0d1117",border:`1px solid ${C.purple}`,borderRadius:4,color:"#c4b5fd",padding:"6px 8px",fontSize:13,fontWeight:700,width:"100%",outline:"none",textAlign:"center"}}}
                        />
                        <span style={{fontSize:11,color:C.muted,flexShrink:0}}>mm</span>
                      </div>
                    </div>
                    <div style={{fontSize:18,color:C.dim,paddingTop:18}}>×</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:10,color:C.muted,marginBottom:4}}>세로 (height)</div>
                      <div style={{display:"flex",alignItems:"center",gap:5}}>
                        <input type="number" min="10" max="400" value={customH}
                          onChange={e=>setCustomH(e.target.value)}
                          onBlur={e=>{const v=Math.min(400,Math.max(10,parseFloat(e.target.value)||10));setCustomH(String(v));}}
                          style={{...{background:"#0d1117",border:`1px solid ${C.purple}`,borderRadius:4,color:"#c4b5fd",padding:"6px 8px",fontSize:13,fontWeight:700,width:"100%",outline:"none",textAlign:"center"}}}
                        />
                        <span style={{fontSize:11,color:C.muted,flexShrink:0}}>mm</span>
                      </div>
                    </div>
                  </div>
                  <div style={{marginTop:10,padding:"7px 10px",background:"#0d1117",borderRadius:5,border:`1px solid ${C.border}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:10}}>
                      <span style={{color:C.muted}}>출력 해상도 (300dpi)</span>
                      <span style={{color:"#c4b5fd",fontWeight:700,fontFamily:"monospace"}}>
                        {Math.round(customW*MM)} × {Math.round(customH*MM)} px
                      </span>
                    </div>
                  </div>
                </div>
              )}
            </Sec>

            <Sec title="배경 / 로고 이미지">
              <ImgUploadBtn id="bg-up"   label="배경 이미지 업로드" accept="image/*" onChange={handleBgFile}   hasImage={!!bgImg}   loading={bgLoading}/>
              <ImgUploadBtn id="logo-up" label="로고 이미지 업로드" accept="image/*" onChange={handleLogoFile} hasImage={!!logoImg} loading={logoLoading}/>
            </Sec>

            <Sec title="테두리 설정">
              <Row label="색상">
                <input type="color" value={border.color} onChange={e=>{setBorder(p=>({...p,color:e.target.value}));setAutoColor(false);}} style={{width:30,height:24,borderRadius:4}}/>
                <span style={{fontSize:10,color:C.muted,fontFamily:"monospace"}}>{border.color}</span>
              </Row>
              <Row label="자동감지"><span style={{fontSize:11,color:C.muted}}>배경 색 자동 인식</span><Toggle on={autoColor} onChange={setAutoColor}/></Row>
              <Row label="두께"><Slider min={0} max={24} value={border.thickness} onChange={v=>setBorder(p=>({...p,thickness:v}))} unit="px"/></Row>
            </Sec>

            <Sec title="글자 설정">
              <Row label="색상">
                <input type="color" value={textColor} onChange={e=>setTextColor(e.target.value)} style={{width:30,height:24,borderRadius:4}}/>
                <span style={{fontSize:10,color:C.muted,fontFamily:"monospace"}}>{textColor}</span>
              </Row>
              <Row label="폰트">
                <select value={fontFamily} onChange={e=>setFontFamily(e.target.value)} style={{...baseInp,padding:"5px 7px",fontSize:11}}>
                  <option value="Pretendard">Pretendard</option>
                  <option value="Noto Sans KR">Noto Sans KR</option>
                  <option value="Arial">Arial</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Courier New">Courier New</option>
                </select>
              </Row>
            </Sec>

            <Sec title="행사명 (상단)" accent={C.blue}>
              <input value={eventName} onChange={e=>setEventName(e.target.value)} placeholder="행사명 입력..." style={{...baseInp,marginBottom:8}}/>
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
                <span style={{fontSize:11,color:C.muted,width:60,flexShrink:0}}>영역 높이</span>
                <Slider min={0.15} max={0.45} value={topHeightRatio} onChange={v=>setTopHeightRatio(v)} unit=""/>
              </div>
              <ZoneCard title="행사명" zone={zones.top} onChange={v=>upZ("top",v)} fsRange={[8,120]}/>
            </Sec>
            <Sec title="중단 구역" accent={C.green}>
              <ZoneCard title="소속 / 업체명" zone={zones.org}      onChange={v=>upZ("org",v)}/>
              <ZoneCard title="직위 / 직급"   zone={zones.position} onChange={v=>upZ("position",v)}/>
              <ZoneCard title="이름"          zone={zones.name}     onChange={v=>upZ("name",v)} fsRange={[10,80]}/>
              <ZoneCard title="추가 필드 1"   zone={zones.field1}   onChange={v=>upZ("field1",v)}/>
              <ZoneCard title="추가 필드 2"   zone={zones.field2}   onChange={v=>upZ("field2",v)}/>
            </Sec>
            <Sec title="로고 위치 (하단)" accent={C.orange}>
              <ZoneCard title="로고" zone={zones.logo} onChange={v=>upZ("logo",v)} showFontSize={false}/>
            </Sec>
          </>)}

          {/* 데이터 탭 */}
          {tab==="data"&&(<>
            <Sec title="데이터 파일 업로드">
              <p style={{fontSize:11,color:C.muted,lineHeight:1.7,marginBottom:10}}>엑셀(.xlsx) 또는 CSV 파일을 업로드하세요.<br/><strong style={{color:C.text}}>첫 번째 행은 헤더</strong>로 인식됩니다.</p>
              <label htmlFor="csv-up" style={{display:"flex",alignItems:"center",gap:8,padding:"8px 11px",background:C.bg,border:`1px dashed ${csvData.length>0?"#238636":C.border2}`,borderRadius:6,cursor:"pointer",marginBottom:6}}>
                <span style={{fontSize:15}}>{csvData.length>0?"✓":"↑"}</span>
                <span style={{fontSize:12,color:csvData.length>0?C.green:C.muted}}>엑셀 / CSV 업로드</span>
                <input id="csv-up" type="file" accept=".xlsx,.xls,.csv" onChange={handleCsv} style={{display:"none"}}/>
              </label>
              {csvData.length>0
                ?<div style={{fontSize:11,color:C.green,padding:"6px 10px",background:"#0d2318",borderRadius:4,border:"1px solid #238636"}}>✓ {csvData.length}건 로드 완료</div>
                :<SampleCsvBox/>
              }
            </Sec>
            {headers.length>0&&(
              <Sec title="열 매핑">
                {[["org","소속 / 업체명"],["position","직위 / 직급"],["name","이름"],["field1","추가 필드 1"],["field2","추가 필드 2"]].map(([k,lbl])=>(
                  <div key={k} style={{marginBottom:9}}>
                    <div style={{fontSize:11,color:C.muted,marginBottom:3}}>{lbl}</div>
                    <select value={colMap[k]} onChange={e=>setColMap(p=>({...p,[k]:+e.target.value}))} style={{...baseInp,padding:"5px 7px",fontSize:11}}>
                      <option value={-1}>— 사용 안함</option>
                      {headers.map((h,i)=><option key={i} value={i}>{h}</option>)}
                    </select>
                  </div>
                ))}
              </Sec>
            )}
            {csvData.length>0&&(
              <Sec title="데이터 탐색">
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                  <button onClick={()=>setPidx(v=>Math.max(0,v-1))} disabled={pidx===0} style={{padding:"5px 10px",background:pidx===0?C.surface:C.border2,color:pidx===0?C.dim:C.text,border:`1px solid ${C.border}`,borderRadius:4,cursor:pidx===0?"not-allowed":"pointer"}}>◀</button>
                  <span style={{flex:1,textAlign:"center",fontSize:12,color:C.muted}}>{pidx+1} / {csvData.length}명</span>
                  <button onClick={()=>setPidx(v=>Math.min(csvData.length-1,v+1))} disabled={pidx===csvData.length-1} style={{padding:"5px 10px",background:pidx===csvData.length-1?C.surface:C.border2,color:pidx===csvData.length-1?C.dim:C.text,border:`1px solid ${C.border}`,borderRadius:4,cursor:pidx===csvData.length-1?"not-allowed":"pointer"}}>▶</button>
                </div>
                <div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:10,maxHeight:200,overflowY:"auto"}}>
                  {headers.map((h,i)=>(
                    <div key={i} style={{display:"flex",gap:8,paddingBottom:5,borderBottom:`1px solid ${C.border}`,marginBottom:5}}>
                      <span style={{fontSize:10,color:C.dim,width:65,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{h}</span>
                      <span style={{fontSize:11,color:C.text}}>{String(csvData[pidx]?.[i]??"")}</span>
                    </div>
                  ))}
                </div>
              </Sec>
            )}
          </>)}

          {/* 출력 탭 */}
          {tab==="export"&&(<>
            <Sec title="A4 레이아웃">
              <A4Preview sz={sz} total={csvData.length}/>
              <div style={{marginTop:10,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"10px 12px"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px 12px",fontSize:11}}>
                  {[["출력 용지","A4 (210×297mm)"],["명찰 사이즈",sz.sub],["페이지당",`${layout.cols}×${layout.rows}=${layout.perPage}매`],["총 페이지",csvData.length>0?`${Math.ceil(csvData.length/layout.perPage)}p`:"-"],["해상도","300 dpi"],["형식","PDF"]].map(([l,v])=>(
                    <div key={l}><div style={{fontSize:9,color:C.dim,marginBottom:1}}>{l}</div><div style={{fontWeight:700,color:C.text}}>{v}</div></div>
                  ))}
                </div>
              </div>
            </Sec>
            <Sec title="단일 PNG 저장">
              <button onClick={exportOne} style={{width:"100%",padding:"8px",background:C.accent,color:"#fff",border:"none",borderRadius:6,fontSize:12,fontWeight:600,cursor:"pointer"}}>📥 현재 미리보기 PNG 저장</button>
            </Sec>
            <Sec title="대량 출력 — A4 PDF">
              {csvData.length===0
                ?<div style={{padding:12,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,fontSize:11,color:C.muted,lineHeight:1.7}}>📋 데이터 탭에서 CSV / 엑셀 파일을<br/>먼저 업로드해 주세요.</div>
                :(<>
                  <div style={{fontSize:12,color:C.text,marginBottom:10}}>총 <span style={{color:C.green,fontWeight:700}}>{csvData.length}명</span> → <span style={{color:C.orange,fontWeight:700}}>A4 {Math.ceil(csvData.length/layout.perPage)}페이지</span> PDF</div>
                  {exporting&&(<div style={{marginBottom:12}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.muted,marginBottom:5}}><span>{expMsg}</span><span>{expProg.cur}/{expProg.total}</span></div>
                    <div style={{height:6,background:C.border2,borderRadius:3}}><div style={{height:6,background:`linear-gradient(90deg,${C.blue},${C.green})`,borderRadius:3,width:`${(expProg.cur/expProg.total)*100}%`,transition:"width .2s"}}/></div>
                  </div>)}
                  <button onClick={exportPDF} disabled={exporting} style={{width:"100%",padding:"10px",background:exporting?C.border:"#238636",color:exporting?C.dim:"#fff",border:"none",borderRadius:6,fontSize:13,fontWeight:700,cursor:exporting?"not-allowed":"pointer"}}>
                    {exporting?`처리 중... (${expProg.cur}/${expProg.total}명)`:`📄 A4 PDF 생성 — 전체 ${csvData.length}명`}
                  </button>
                  <p style={{fontSize:10,color:C.muted,marginTop:8,lineHeight:1.6}}>인쇄 시 <strong style={{color:C.text}}>실제 크기(100%)</strong>로 설정하세요.</p>
                </>)
              }
            </Sec>
          </>)}

        </div>
      </div>

      {/* 미리보기 영역 */}
      <div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,padding:24,overflow:"hidden",background:"radial-gradient(ellipse at 60% 40%, #161b22 0%, #0d1117 75%)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:10,padding:"3px 10px",background:"#0d2137",border:"1px solid #388bfd",borderRadius:10,color:C.blue,fontWeight:700}}>{sz.label}</span>
          <span style={{fontSize:10,color:C.muted}}>{sz.sub}</span>
          <span style={{fontSize:10,color:C.dim}}>· {sz.ew}×{sz.eh}px · 300dpi</span>
        </div>

        <div style={{boxShadow:"0 4px 6px rgba(0,0,0,.3),0 20px 60px rgba(0,0,0,.6)",borderRadius:2,overflow:"hidden",width:dW,height:dH,flexShrink:0}}>
          <canvas ref={canvasRef} style={{display:"block",width:dW,height:dH}}/>
        </div>

        {csvData.length>0&&(
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <button onClick={()=>setPidx(v=>Math.max(0,v-1))} disabled={pidx===0} style={{padding:"5px 12px",background:pidx===0?C.surface:C.border2,color:pidx===0?C.dim:C.text,border:`1px solid ${C.border}`,borderRadius:4,cursor:pidx===0?"not-allowed":"pointer"}}>◀</button>
            <span style={{fontSize:12,color:C.muted,minWidth:80,textAlign:"center"}}>{pidx+1} / {csvData.length}명</span>
            <button onClick={()=>setPidx(v=>Math.min(csvData.length-1,v+1))} disabled={pidx===csvData.length-1} style={{padding:"5px 12px",background:pidx===csvData.length-1?C.surface:C.border2,color:pidx===csvData.length-1?C.dim:C.text,border:`1px solid ${C.border}`,borderRadius:4,cursor:pidx===csvData.length-1?"not-allowed":"pointer"}}>▶</button>
          </div>
        )}

        <div style={{display:"flex",gap:8}}>
          <button onClick={exportOne} style={{padding:"8px 16px",background:C.accent,color:"#fff",border:"none",borderRadius:6,fontSize:12,fontWeight:600,cursor:"pointer"}}>🖼 PNG 저장</button>
          {csvData.length>0&&(
            <button onClick={exportPDF} disabled={exporting} style={{padding:"8px 16px",background:exporting?C.border:"#238636",color:exporting?C.dim:"#fff",border:"none",borderRadius:6,fontSize:12,fontWeight:600,cursor:exporting?"not-allowed":"pointer"}}>
              {exporting?`PDF 생성 중 ${expProg.cur}/${expProg.total}`:`📄 A4 PDF (${csvData.length}명)`}
            </button>
          )}
        </div>

        {exporting&&(
          <div style={{width:300,padding:"10px 14px",background:C.surface,border:`1px solid ${C.border}`,borderRadius:8}}>
            <div style={{fontSize:11,color:C.muted,marginBottom:6}}>{expMsg}</div>
            <div style={{height:5,background:C.border2,borderRadius:3}}>
              <div style={{height:5,background:`linear-gradient(90deg,${C.blue},${C.green})`,borderRadius:3,width:`${(expProg.cur/expProg.total)*100}%`,transition:"width .2s"}}/>
            </div>
            <div style={{fontSize:10,color:C.dim,marginTop:5,textAlign:"right"}}>{expProg.cur} / {expProg.total}명</div>
          </div>
        )}
      </div>
    </div>
    </>
  );
}
