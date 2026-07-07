(function(){
  function stars(n){n=Math.max(0,Math.min(5,parseInt(n,10)||0));var s='';for(var i=0;i<5;i++)s+=(i<n?'★':'☆');return s;}
  function initOne(root){
    var track=root.querySelector('.mm-rev__track');
    var cards=root.querySelectorAll('.mm-rev__card');
    var dotsWrap=root.querySelector('.mm-rev__dots');
    var total=cards.length;
    if(!track||!total)return;
    // 별점 렌더
    for(var i=0;i<total;i++){var st=cards[i].querySelector('.mm-rev__stars');if(st)st.textContent=stars(st.getAttribute('data-rating'));}
    var cur=0,timer=null;
    var speed=(parseFloat(root.getAttribute('data-speed'))||4)*1000;
    // 점
    var dots=[];
    if(dotsWrap&&total>1){for(var d=0;d<total;d++){(function(idx){var b=document.createElement('button');b.className='mm-rev__dot'+(idx===0?' is-on':'');b.setAttribute('aria-label',(idx+1)+'번 후기');b.addEventListener('click',function(){go(idx);restart();});dotsWrap.appendChild(b);dots.push(b);})(d);}}
    function go(n){cur=(n%total+total)%total;track.style.transform='translateX(-'+(cur*100)+'%)';for(var k=0;k<dots.length;k++)dots[k].classList.toggle('is-on',k===cur);}
    function next(){go(cur+1);}
    function start(){if(total>1&&speed>0){stop();timer=setInterval(next,speed);}}
    function stop(){if(timer){clearInterval(timer);timer=null;}}
    function restart(){start();}
    root.addEventListener('mouseenter',stop);root.addEventListener('mouseleave',start);
    // 스와이프
    var sx=0;
    track.addEventListener('touchstart',function(e){sx=e.changedTouches[0].clientX;stop();},{passive:true});
    track.addEventListener('touchend',function(e){var dx=e.changedTouches[0].clientX-sx;if(Math.abs(dx)>40){dx<0?next():go(cur-1);}start();},{passive:true});
    go(0);start();
  }
  function init(){var l=document.querySelectorAll('.mm-rev');if(!l.length){return setTimeout(init,50);}for(var i=0;i<l.length;i++)initOne(l[i]);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
