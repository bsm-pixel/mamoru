(function(){
  var DAYS=['일','월','화','수','목','금','토'];
  function toMin(s){if(!s)return null;var pm=/오후|PM/i.test(s),am=/오전|AM/i.test(s);var m=String(s).match(/(\d{1,2})\D+(\d{2})/);if(!m)return null;var h=+m[1],mi=+m[2];if(pm&&h<12)h+=12;if(am&&h===12)h=0;return h*60+mi;}
  function initOne(root){
    var data=root.querySelector('.mm-hr__data');
    var stateEl=root.querySelector('.mm-hr__state');
    var detailEl=root.querySelector('.mm-hr__detail');
    if(!data)return;
    var openLabel=data.getAttribute('data-open')||'영업 중';
    var closeLabel=data.getAttribute('data-close')||'영업 종료';
    var rows=data.querySelectorAll('span');
    var now=new Date(),today=DAYS[now.getDay()],curMin=now.getHours()*60+now.getMinutes();
    var todayRow=null;
    for(var i=0;i<rows.length;i++){var d=rows[i].getAttribute('data-day')||'';if(d.indexOf(today)>=0){todayRow=rows[i];break;}}
    var isOpen=false,detail='';
    if(todayRow){
      var off=todayRow.getAttribute('data-off')==='true';
      var o=toMin(todayRow.getAttribute('data-o')),c=toMin(todayRow.getAttribute('data-c'));
      if(off){detail='오늘 휴무';}
      else if(o!=null&&c!=null){
        isOpen=curMin>=o&&curMin<c;
        function hm(x){var h=Math.floor(x/60),m=x%60;return (h<10?'0'+h:h)+':'+(m<10?'0'+m:m);}
        detail=hm(o)+' ~ '+hm(c);
      }
    }else{detail='오늘 영업정보 없음';}
    root.classList.toggle('is-open',isOpen);
    if(stateEl)stateEl.textContent=isOpen?openLabel:closeLabel;
    if(detailEl)detailEl.textContent=detail;
  }
  function init(){var l=document.querySelectorAll('.mm-hr');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
