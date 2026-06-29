(function(){
  function splitTags(s){return String(s||'').split(',').map(function(t){return t.trim();}).filter(Boolean);}
  function initOne(root){
    var chipsWrap=root.querySelector('.mm-flt__chips');
    var cards=root.querySelectorAll('.mm-flt__card');
    var allLabel=root.getAttribute('data-all')||'전체';
    if(!chipsWrap||!cards.length)return;
    // 태그 수집
    var set={},order=[];
    for(var i=0;i<cards.length;i++){var ts=splitTags(cards[i].getAttribute('data-tags'));for(var j=0;j<ts.length;j++){if(!set[ts[j]]){set[ts[j]]=1;order.push(ts[j]);}}}
    var chips=[];
    function mkChip(label,tag,on){var b=document.createElement('button');b.type='button';b.className='mm-flt__chip'+(on?' is-on':'');b.textContent=label;b.addEventListener('click',function(){select(tag);});chipsWrap.appendChild(b);chips.push({el:b,tag:tag});}
    mkChip(allLabel,'__all',true);
    for(var k=0;k<order.length;k++)mkChip(order[k],order[k],false);
    function select(tag){
      for(var c=0;c<chips.length;c++)chips[c].el.classList.toggle('is-on',chips[c].tag===tag);
      for(var m=0;m<cards.length;m++){var has=tag==='__all'||splitTags(cards[m].getAttribute('data-tags')).indexOf(tag)>=0;cards[m].classList.toggle('is-hide',!has);}
    }
    if(order.length===0)chipsWrap.style.display='none';
  }
  function init(){var l=document.querySelectorAll('.mm-flt');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
