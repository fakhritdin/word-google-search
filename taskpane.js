(function(){
  const q=document.getElementById('q'); const status=document.getElementById('status');
  function query(){return q.value.trim();}
  function openSearch(base){const s=query();if(!s){status.textContent='Аввал қидирув сўровини киритинг ёки Word матнини танланг.';return;} const url=base+encodeURIComponent(s); window.open(url,'_blank','noopener,noreferrer');}
  document.getElementById('web').addEventListener('click',()=>openSearch('https://www.google.com/search?q='));
  document.getElementById('scholar').addEventListener('click',()=>openSearch('https://scholar.google.com/scholar?q='));
  document.getElementById('selection').addEventListener('click',async()=>{try{const r=await Word.run(async context=>{const range=context.document.getSelection();range.load('text');await context.sync();return range.text;});if(r&&r.trim()){q.value=r.trim();status.textContent='Word’дан танланган матн олинди.';}else status.textContent='Аввал Word ҳужжатида матнни белгиланг.';}catch(e){status.textContent='Матнни олиш учун Word ҳужжатида матнни белгиланг ва қайта урининг.';}});
  if(typeof Office!=='undefined'){Office.onReady(()=>{});}
})();
