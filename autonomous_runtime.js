(()=>{'use strict';
// Keep legacy generator code unchanged while redirecting its remaining runtime fetches locally.
const LOCAL_FONT='./vendor/fonts/DejaVuSans.ttf';
const originalFetch=window.fetch.bind(window);
window.fetch=async function(input,init){
  const url=typeof input==='string'?input:(input&&input.url)||'';
  if(url.includes('/dejavu-fonts-ttf@2.37.3/ttf/DejaVuSans.ttf')) return originalFetch(LOCAL_FONT,init);
  return originalFetch(input,init);
};
if(window.pdfjsLib?.GlobalWorkerOptions) window.pdfjsLib.GlobalWorkerOptions.workerSrc='./vendor/pdfjs/pdf.worker.min.js';
window.addEventListener('load',()=>{if(window.pdfjsLib?.GlobalWorkerOptions)window.pdfjsLib.GlobalWorkerOptions.workerSrc='./vendor/pdfjs/pdf.worker.min.js';},{once:true});
})();
