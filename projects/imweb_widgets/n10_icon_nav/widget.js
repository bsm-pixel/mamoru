(function(){
  var SVG={
    scissors:'<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/>',
    brush:'<rect x="3" y="4" width="18" height="4" rx="1"/><line x1="7" y1="8" x2="7" y2="14"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="15" y1="8" x2="15" y2="14"/><line x1="19" y1="8" x2="19" y2="14"/>',
    box:'<path d="M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>',
    wrench:'<path d="M14.7 6.3a1 1 0 000 1.4l1.6 1.6a1 1 0 001.4 0l3.77-3.77a6 6 0 01-7.94 7.94L6.7 20.2a2.1 2.1 0 01-3-3l6.73-6.73a6 6 0 017.94-7.94L14.7 6.3z"/>',
    chat:'<path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/>',
    star:'<polygon points="12 2 15.1 8.3 22 9.3 17 14.1 18.2 21 12 17.8 5.8 21 7 14.1 2 9.3 8.9 8.3 12 2"/>',
    home:'<path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
    phone:'<path d="M22 16.92v3a2 2 0 01-2.18 2 19.8 19.8 0 01-8.63-3.07 19.5 19.5 0 01-6-6 19.8 19.8 0 01-3.07-8.67A2 2 0 014.11 2h3a2 2 0 012 1.72c.13.96.36 1.9.7 2.81a2 2 0 01-.45 2.11L8.09 9.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0122 16.92z"/>'
  };
  function pick(k){
    k=(k||'').toLowerCase();
    if(/가위집|케이스|case|box|package|pouch/.test(k))return 'box';
    if(/가위|scissor|미용|cut/.test(k))return 'scissors';
    if(/빗|브러시|brush|comb/.test(k))return 'brush';
    if(/복원|수리|repair|wrench|tool|as/.test(k))return 'wrench';
    if(/상담|컨설|consult|chat|message|talk|문의/.test(k))return 'chat';
    if(/후기|리뷰|review|star|별|평점/.test(k))return 'star';
    if(/홈|home|메인/.test(k))return 'home';
    if(/전화|연락|phone|call/.test(k))return 'phone';
    return null;
  }
  function initOne(root){
    var chips=root.querySelectorAll('.mm-nav__chip');
    for(var i=0;i<chips.length;i++){
      var ico=chips[i].querySelector('.mm-nav__ico');
      if(!ico)continue;
      /* 커스텀 SVG 우선. textContent라 아임웹이 escape해도 원본 SVG 문자열이 잡혀 innerHTML로 정상 렌더 */
      var srcEl=chips[i].querySelector('.mm-nav__svgsrc'),custom=srcEl?String(srcEl.textContent||'').trim():'';
      if(custom){ ico.innerHTML=custom; continue; }
      var name=pick(chips[i].getAttribute('data-icon'));
      if(name)ico.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'+SVG[name]+'</svg>';
    }
  }
  function init(){var l=document.querySelectorAll('.mm-nav');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
