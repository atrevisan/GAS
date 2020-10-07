function onOpen(e) {
  Lib.setClassesDropdown();
  Lib.setTiposProventosDropdown();
}

function onEdit(e) {
  var ss = e.source,
    active = ss.getActiveSheet().getName(),
    rangePosition = e.range.getA1Notation();

  if (active === "Carteira" && rangePosition.startsWith("J")) Lib.atualizarPesos();
}

function limparCarteira() {
  Lib.limparCarteira();
}

function adicionarProvento() {
  Lib.adicionarProvento();
}

function addOrdem() {
  Lib.adicionarOrdem();
}

