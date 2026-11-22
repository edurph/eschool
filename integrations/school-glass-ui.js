/* School Light Glass UI V2 — include inside the SAME document as the school app.
   UI-only: no network calls, changes to school data, or login interception. */
(function (global) {
  'use strict';
  if (global.SchoolGlassUI) return;
  const KEY = 'sk-school-glass-v2';
  const defaults = { theme: 'light', sound: true, volume: 0.75, motion: true, notifications: true };
  let options = { ...defaults };
  try {
    const stored = JSON.parse(localStorage.getItem(KEY) || '{}');
    options = { ...options, ...stored };
  } catch (_) {}
  let audioContext;
  let lastInteraction = 0;
  let toastTimer;
  let toastNode;
  const css = `
    :root {
      --sg-glass: rgba(255,255,255,.72); --sg-border: rgba(255,255,255,.82);
      --sg-text:#243457; --sg-shadow:0 12px 35px rgba(44,75,130,.12);
    }
    .sg-glass { color:var(--sg-text);background:var(--sg-glass);
      border:1px solid var(--sg-border);box-shadow:var(--sg-shadow);
      backdrop-filter:blur(16px) saturate(140%);-webkit-backdrop-filter:blur(16px) saturate(140%); }
    .sg-interactive { transition:transform .22s ease,box-shadow .22s ease,filter .22s ease; }
    @media (hover:hover) { .sg-motion .sg-interactive:hover { transform:translateY(-2px);filter:brightness(1.035);box-shadow:0 12px 28px rgba(75,105,180,.15); } }
    .sg-interactive:focus-visible,.sg-beacon:focus-visible {outline:3px solid #3b82f6;outline-offset:3px}
    .sg-beacon { position:fixed;right:18px;bottom:calc(18px + env(safe-area-inset-bottom,0px));z-index:2147481000;
      border:1px solid rgba(255,255,255,.85);width:44px;height:44px;border-radius:50%;
      background:rgba(255,255,255,.8);box-shadow:0 5px 20px rgba(77,96,150,.19);
      backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
      cursor:pointer;font:600 19px system-ui;color:#3d5aa7; }
    .sg-beacon[data-muted="false"]::after {content:"";position:absolute;inset:-4px;border-radius:50%;
      border:1px solid rgba(96,165,250,.35);box-shadow:0 0 12px rgba(96,165,250,.23);pointer-events:none}
    .sg-toast {position:fixed;z-index:2147481001;left:50%;bottom:calc(78px + env(safe-area-inset-bottom,0px));
      transform:translateX(-50%);max-width:min(92vw,440px);min-width:180px;padding:13px 18px;
      border-radius:18px;font:500 14px/1.4 system-ui;text-align:center;pointer-events:none;}
    .sg-toast[data-type="error"] {border-color:rgba(239,68,68,.35)}
    .sg-toast[data-type="success"] {border-color:rgba(34,197,94,.35)}
    @media (prefers-reduced-motion:reduce) { .sg-interactive {transition:none!important;transform:none!important} }
    .sg-no-motion .sg-interactive {transition:none!important;transform:none!important}
  `;
  function save() { try {localStorage.setItem(KEY, JSON.stringify(options));} catch (_) {} }
  function apply() {
    document.documentElement.classList.toggle('sg-motion',!!options.motion);
    document.documentElement.classList.toggle('sg-no-motion',!options.motion);
    const b=document.getElementById('sg-sound-beacon');
    if(b) {b.dataset.muted=String(!options.sound);b.setAttribute('aria-label',options.sound?'Matikan bunyi UI':'Aktifkan bunyi UI');b.setAttribute('aria-pressed',String(options.sound));b.textContent=options.sound?'♫':'♪';}
  }
  function beep(kind='click') {
    if(!options.sound || options.volume<=0) return;
    try {
      const AC=global.AudioContext||global.webkitAudioContext;
      if(!AC) return;
      audioContext=audioContext||new AC();
      if(audioContext.state==='suspended') audioContext.resume().catch(()=>{});
      const tones=kind==='success'?[620,830]:kind==='error'?[310,240]:[530];
      const now=audioContext.currentTime;
      tones.forEach((freq,i)=>{
        const osc=audioContext.createOscillator(), gain=audioContext.createGain(), start=now+i*.075;
        osc.type='sine';osc.frequency.setValueAtTime(freq,start);
        gain.gain.setValueAtTime(.0001,start);
        gain.gain.exponentialRampToValueAtTime(Math.max(.0002,options.volume*.055),start+.008);
        gain.gain.exponentialRampToValueAtTime(.0001,start+.09);
        osc.connect(gain);gain.connect(audioContext.destination);osc.start(start);osc.stop(start+.095);
      });
    } catch (_) {}
  }
  function toast(message,type='info',duration=2800) {
    if(!options.notifications || !message) return;
    if(!toastNode) {toastNode=document.createElement('div');toastNode.className='sg-toast sg-glass';toastNode.setAttribute('role','status');toastNode.setAttribute('aria-live','polite');document.body.appendChild(toastNode);}
    clearTimeout(toastTimer);
    toastNode.textContent=String(message);toastNode.dataset.type=type;toastNode.hidden=false;
    toastTimer=setTimeout(()=>{toastNode.hidden=true;},Math.max(800,duration));
    if(type==='success'||type==='error') beep(type);
  }
  function mount() {
    if(!document.getElementById('sg-glass-style')) {
      const style=document.createElement('style');style.id='sg-glass-style';style.textContent=css;document.head.appendChild(style);
    }
    if(!document.getElementById('sg-sound-beacon')) {
      const button=document.createElement('button');button.type='button';button.id='sg-sound-beacon';button.className='sg-beacon';
      button.addEventListener('click',()=>{options.sound=!options.sound;save();apply();if(options.sound)beep();});
      document.body.appendChild(button);
    }
    // Only genuine user clicks; does not misinterpret programmatic updates as successful saves.
    document.addEventListener('click',(event)=>{
      const el=event.target && event.target.closest && event.target.closest('button,[role="button"],input[type="submit"],input[type="button"]');
      if(!el || el.id==='sg-sound-beacon' || el.disabled || el.getAttribute('aria-disabled')==='true') return;
      const t=Date.now();if(t-lastInteraction<45)return;lastInteraction=t;beep('click');
    },{passive:true});
    apply();
  }
  const api={
    init:mount,
    play:beep,
    notify:toast,
    success:(message)=>toast(message,'success'),
    error:(message)=>toast(message,'error'),
    getSettings:()=>({...options}),
    setSettings:(next)=>{
      const theme=next && next.theme;
      options={...options,...(next||{}),theme:theme==='dark'||theme==='auto'?'light':options.theme};
      options.volume=Math.min(1,Math.max(0,Number(options.volume)||0));
      save();apply();return {...options};
    }
  };
  global.SchoolGlassUI=api;
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',mount,{once:true});
  else mount();
})(window);
