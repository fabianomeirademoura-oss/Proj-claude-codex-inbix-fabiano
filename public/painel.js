// Comportamento do painel no navegador. Arquivo à parte: a CSP proíbe script inline.
// Sem este arquivo, tudo continua funcionando; ele só melhora o uso no celular.

// No celular o menu é uma fileira de abas que rola de lado: mostra a aba da página atual.
document.addEventListener('DOMContentLoaded', function () {
  var nav = document.querySelector('.nav');
  var ativa = nav && nav.querySelector('a.ativo');
  if (ativa && nav.scrollWidth > nav.clientWidth) nav.scrollLeft = ativa.offsetLeft - nav.offsetLeft - (nav.clientWidth - ativa.offsetWidth) / 2;
});

// App instalável (PWA): registra o service worker.
if ('serviceWorker' in navigator) {
  window.addEventListener('load', function () { navigator.serviceWorker.register('/sw.js').catch(function () { /* sem PWA, o painel funciona igual */ }); });
}
