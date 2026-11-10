// 共用：API、跳脫、星空背景
const API=(window.OCTO_API||'').trim();
const GROUPS=['對唱','新人','獨唱','長青'];
const GICON={'對唱':'💞','新人':'🌱','獨唱':'🎤','長青':'🌿'};
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
const label=p=>p.club?`${p.name}（${p.club}）`:p.name;
async function apiGet(){const r=await fetch(API+(API.includes('?')?'&':'?')+'t='+Date.now());return r.json()}
async function apiPost(body){const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(body)});return r.json()}
function starfield(canvas,opt={}){
  const cx=canvas.getContext('2d');let stars=[];const S={parts:[],hot:null};
  function fit(){canvas.width=canvas.offsetWidth;canvas.height=canvas.offsetHeight;
    stars=Array.from({length:opt.n||140},()=>({x:Math.random()*canvas.width,y:Math.random()*canvas.height,r:Math.random()*1.6+.3,p:Math.random()*6.28}))}
  function loop(){cx.clearRect(0,0,canvas.width,canvas.height);const t=Date.now()/600;
    stars.forEach(s=>{cx.globalAlpha=.35+.6*Math.abs(Math.sin(t+s.p));cx.fillStyle='#fff';cx.beginPath();cx.arc(s.x,s.y,s.r,0,6.28);cx.fill()});
    if(S.hot){const h=S.hot();if(h&&Math.random()<.7)S.parts.push({x:h.x+(Math.random()-.5)*h.w,y:h.y,vx:(Math.random()-.5),vy:-Math.random()*4-2,l:1,c:'#ffcf6b',s:Math.random()*2.5+1})}
    S.parts=S.parts.filter(p=>p.l>0);S.parts.forEach(p=>{p.x+=p.vx;p.y+=p.vy;p.vy+=.12;p.l-=.012;cx.globalAlpha=Math.max(p.l,0);cx.fillStyle=p.c;cx.shadowBlur=10;cx.shadowColor=p.c;cx.beginPath();cx.arc(p.x,p.y,p.s,0,6.28);cx.fill();cx.shadowBlur=0});
    cx.globalAlpha=1;requestAnimationFrame(loop)}
  S.burstAt=(ox,oy,cnt)=>{for(let i=0;i<cnt;i++){const a=Math.random()*6.28,v=Math.random()*9+2;S.parts.push({x:ox,y:oy,vx:Math.cos(a)*v,vy:Math.sin(a)*v-2,l:1,c:['#ffe9a8','#d4af37','#b57bff','#ffffff','#ff9ad5','#ff8a65'][i%6],s:Math.random()*3+1})}};
  S.burstEl=(el,cnt)=>{const r=el.getBoundingClientRect(),c=canvas.getBoundingClientRect();S.burstAt(r.left-c.left+r.width/2,r.top-c.top+r.height/2,cnt||220)};
  window.addEventListener('resize',fit);fit();loop();S.fit=fit;return S;
}
