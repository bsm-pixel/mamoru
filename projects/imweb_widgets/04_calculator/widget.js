(function(){
  function num(v){return parseFloat(String(v).replace(/[^0-9.]/g,''))||0;}
  function fmt(n){return Math.round(n).toLocaleString('en-US');}
  function initOne(root){
    var spans=root.querySelectorAll('.mm-calc__data span');
    var sel=root.querySelector('.mm-calc__model');
    var qty=root.querySelector('.mm-calc__qty');
    var gradesWrap=root.querySelector('.mm-calc__grades');
    var totalEl=root.querySelector('.mm-calc__total');
    if(!sel||!qty||!totalEl)return;
    // 모델 옵션
    var models=[];
    for(var i=0;i<spans.length;i++){var nm=spans[i].getAttribute('data-name')||('모델'+(i+1));var pr=num(spans[i].getAttribute('data-price'));models.push({name:nm,price:pr});var o=document.createElement('option');o.value=i;o.textContent=nm;sel.appendChild(o);}
    // 등급
    var grades=[];
    for(var g=1;g<=3;g++){var gn=root.getAttribute('data-g'+g+'n');if(gn){grades.push({name:gn,disc:num(root.getAttribute('data-g'+g+'d'))});}}
    var gi=0,gbtns=[];
    for(var k=0;k<grades.length;k++){(function(idx){var b=document.createElement('button');b.type='button';b.className='mm-calc__gbtn'+(idx===0?' is-on':'');b.textContent=grades[idx].name;b.addEventListener('click',function(){gi=idx;for(var x=0;x<gbtns.length;x++)gbtns[x].classList.toggle('is-on',x===idx);calc();});gradesWrap.appendChild(b);gbtns.push(b);})(k);}
    function calc(){
      var m=models[sel.value|0]||{price:0};
      var q=Math.max(1,num(qty.value));
      var disc=grades[gi]?grades[gi].disc:0;
      var total=m.price*q*(1-disc/100);
      totalEl.textContent=fmt(total)+'원';
    }
    sel.addEventListener('change',calc);qty.addEventListener('input',calc);
    calc();
  }
  function init(){var l=document.querySelectorAll('.mm-calc');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();

/* 모서리(radius): 인라인 style {{}}는 아임웹 저장거부 → data-radius 속성값을 JS로 적용(깜빡임 없음, CSS 기본은 각지게) */
(function(){function apR(){var es=document.querySelectorAll("[data-radius]");for(var i=0;i<es.length;i++){var v=String(es[i].getAttribute("data-radius")||"").trim();if(v){if(String(parseFloat(v))===v)v+="px";es[i].style.borderRadius=v;}}}if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",apR);else apR();})();
