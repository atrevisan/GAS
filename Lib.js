function transpose(a) {
    return a[0].map(function (_, c) {
      return a.map(function (r) {
        return r[c];
      });
    });
    // or in more modern dialect
    // return a[0].map((_, c) => a.map(r => r[c]));
  }
  
  function adicionarEvolucao() {
    var planilhaDados = getPlanilha("Dados");
    var planilhaAux2 = getPlanilha("Aux2");
    var planilhaEvolucao = getPlanilha("Evolução");
  
    if (planilhaAux2.getLastRow() > 1) {
      var chts = planilhaEvolucao.getCharts();
      for (var i = 0; i < chts.length; i++) {
        planilhaEvolucao.removeChart(chts[i]);
      }
  
      planilhaAux2.deleteRows(1, planilhaAux2.getLastRow());
    }
  
    var dates = planilhaDados
      .getRange(`G8:G${planilhaDados.getLastRow()}`)
      .getValues()
      .map(function (r) {
        return r[0];
      })
      .sort(function (a, b) {
        return a - b;
      });
  
    var menorData = dates[0];
    var maiorData = dates[dates.length - 1];
  
    var col = 1;
    for (var y = menorData.getFullYear(); y <= maiorData.getFullYear(); y++) {
      for (
        var m = y === menorData.getFullYear() ? menorData.getMonth() + 1 : 1;
        y === maiorData.getFullYear() ? m <= maiorData.getMonth() + 1 : m <= 12;
        m++
      ) {
        var data = `${
          y === maiorData.getFullYear() && maiorData.getMonth() + 1 === m
            ? new Date().getDay() === 6
              ? new Date().getDate() - 1
              : new Date().getDay() === 0
              ? new Date().getDate() - 2
              : new Date().getDate()
            : "28"
        }/${m}/${y}`;
  
        var carteiraAtePeriodo = planilhaDados
          .getRange(`B8:H${planilhaDados.getLastRow()}`)
          .getValues()
          .filter(function (r) {
            return r[5] <= new Date(`${m}/30/${y}`);
          })
          .map(function (r) {
            return [
              `=${r[1]} * INDEX(GOOGLEFINANCE("${r[0]}"; "price"; "${data}"); 2; 2)`,
            ];
          });
  
        planilhaAux2
          .getRange(3, col, carteiraAtePeriodo.length, 1)
          .setValues(carteiraAtePeriodo);
  
        planilhaAux2.getRange(1, col).setValue(data);
  
        planilhaAux2
          .getRange(2, col)
          .setValue(
            `=SUM(INDIRECT(SUBSTITUTE(CONCAT(ADDRESS(3;${col}); CONCAT(":"; ADDRESS(${
              carteiraAtePeriodo.length + 2
            }; ${col}))); "$"; "")))`
          );
  
        col++;
      }
    }
  
    planilhaAux2
      .getRange(1, col + 2, col, 2)
      .setValues(transpose(planilhaAux2.getRange(1, 1, 2, col).getValues()));
  
    for (var i = 1; i < col; i++) {
      var data = planilhaAux2.getRange(i, col + 2).getValue();
  
      var aportesNoMes = planilhaDados
        .getRange(`B8:H${planilhaDados.getLastRow()}`)
        .getValues()
        .filter(function (r) {
          return (
            r[5].getFullYear() === data.getFullYear() &&
            r[5].getMonth() + 1 === data.getMonth() + 1 &&
            r[4] > 0
          );
        })
        .map(function (r) {
          return r[4];
        });
  
      if (aportesNoMes && aportesNoMes.length > 0) {
        var totalAportadoMes = aportesNoMes.reduce(function (acum, curr) {
          return acum + curr;
        });
  
        planilhaAux2.getRange(i, col + 4).setValue(totalAportadoMes);
      }
    }
  
    var chart = planilhaEvolucao
      .newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(planilhaAux2.getRange(1, col + 2, col, 3))
      .setPosition(3, 2, 0, 0)
      .setOption("series", {
        0: { type: "area", color: "red", labelInLegend: "Evolução" },
        1: { type: "bars", color: "blue", labelInLegend: "Aportes" },
      })
      .setOption("useFirstColumnAsDomain", true)
      .build();
  
    planilhaEvolucao.insertChart(chart);
  }
  
  function modifyChart() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var chart = sheet.getCharts()[0];
    chart = chart
      .modify()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(sheet.getRange("a1:a18"))
      .addRange(sheet.getRange("b1:b18"))
      .addRange(sheet.getRange("c1:c18"))
      .setPosition(5, 5, 0, 0)
      .setOption("series", {
        0: { type: "bars", color: "red" },
        1: { type: "line", color: "blue", curveType: "function" },
      })
      .setOption("useFirstColumnAsDomain", true)
      .setOption("height", 280)
      .setOption("width", 480)
      .setOption("title", "Sample chart")
      .setOption("vAxis", {
        minValue: 0,
        maxValue: 0.4,
        gridlines: {
          count: 10,
        },
      })
      .build();
    sheet.updateChart(chart);
  }
  
  function calcularPercentagemTeto() {
    var sheetCarteira = getPlanilha("Carteira");
  
    var linhas = sheetCarteira.getLastRow();
  
    if (linhas === 7) return;
  
    for (var i = 8; i <= linhas; i++) {
      var preco = sheetCarteira.getRange(`D${i}`).getValue();
      var precoTeto = sheetCarteira.getRange(`I${i}`).getValue();
  
      var perc = preco / (precoTeto === 0 ? 1 : precoTeto) - 1;
  
      sheetCarteira.getRange(`Y${i}`).setValue(perc);
    }
  
    ordenarAtivos();
  }
  
  function getPieValuesAtivos(classe, planilhaCarteira) {
    var rows = planilhaCarteira
      .getRange(`B8:Z${planilhaCarteira.getLastRow()}`)
      .getValues();
  
    if (classe === -1)
      return rows.map(function (r) {
        return [r[0], r[4]];
      });
  
    var filteredRows = rows.filter(function (r) {
      return r[24] === classe;
    });
  
    if (filteredRows.length === 0) return [];
  
    var total = filteredRows
      .map(function (r) {
        return r[3];
      })
      .reduce(function (acum, curr) {
        return acum + curr;
      });
  
    return filteredRows.map(function (r) {
      return [r[0], r[3] / total];
    });
  }
  
  function getPieValuesClasses(planilhaCarteira) {
    var rows = planilhaCarteira
      .getRange(`B8:Z${planilhaCarteira.getLastRow()}`)
      .getValues();
  
    var countETF = rows
      .filter(function (r) {
        return r[24] === 0;
      })
      .map(function (r) {
        return r[3];
      })
      .reduce(function (acum, curr) {
        return acum + curr;
      });
  
    var countAcoes = rows
      .filter(function (r) {
        return r[24] === 1;
      })
      .map(function (r) {
        return r[3];
      })
      .reduce(function (acum, curr) {
        return acum + curr;
      });
  
    var countFII = rows
      .filter(function (r) {
        return r[24] === 2;
      })
      .map(function (r) {
        return r[3];
      })
      .reduce(function (acum, curr) {
        return acum + curr;
      });
  
    var total = countETF + countFII + countAcoes;
    return [
      ["ETF", countETF / total],
      ["Ações", countAcoes / total],
      ["FII", countFII / total],
    ];
  }
  
  function getPieValuesSetores(classe, planilhaCarteira) {
    var rows = planilhaCarteira
      .getRange(`B8:Z${planilhaCarteira.getLastRow()}`)
      .getValues();
  
    if (classe !== -1) {
      rows = rows.filter(function (r) {
        return r[24] === classe;
      });
    }
  
    var setores = {};
  
    var total = 0;
    for (var i = 0; i < rows.length; i++) {
      var setor = rows[i][12];
      var valor = rows[i][3];
  
      if (setor in setores) {
        total += valor;
        setores[setor] += valor;
      } else if (setor && setor !== "") {
        total += valor;
        setores[setor] = valor;
      }
    }
  
    return Object.keys(setores).map(function (k) {
      return [k, setores[k] / total];
    });
  }
  
  function adicionarPie(valores, planilhaAlocacao, range, posicao) {
    if (valores.length > 0) {
      planilhaAlocacao.getRange(range).setValues(valores);
  
      var chart = planilhaAlocacao
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(planilhaAlocacao.getRange(range))
        .setPosition(posicao, 2, 0, 0)
        .build();
  
      planilhaAlocacao.insertChart(chart);
    }
  }
  
  function adicionarPies() {
    var planilhaCarteira = getPlanilha("Carteira");
    var planilhaAlocacao = getPlanilha("Alocação");
  
    if (planilhaCarteira.getLastRow() === 7) return;
  
    var chts = planilhaAlocacao.getCharts();
    for (var i = 0; i < chts.length; i++) {
      planilhaAlocacao.removeChart(chts[i]);
    }
  
    var valoresTodosAtivos = getPieValuesAtivos(-1, planilhaCarteira);
    var valoresETF = getPieValuesAtivos(0, planilhaCarteira);
    var valoresAcoes = getPieValuesAtivos(1, planilhaCarteira);
    var valoresFII = getPieValuesAtivos(2, planilhaCarteira);
    var valoresClasses = getPieValuesClasses(planilhaCarteira);
    var valoresTodosSetores = getPieValuesSetores(-1, planilhaCarteira);
    var valoresSetoresETF = getPieValuesSetores(0, planilhaCarteira);
    var valoresSetoresAcoes = getPieValuesSetores(1, planilhaCarteira);
    var valoresSetoresFII = getPieValuesSetores(2, planilhaCarteira);
  
    adicionarPie(
      valoresTodosAtivos,
      planilhaAlocacao,
      `Y1:Z${valoresTodosAtivos.length}`,
      3
    );
    adicionarPie(valoresETF, planilhaAlocacao, `W1:X${valoresETF.length}`, 23);
    adicionarPie(
      valoresAcoes,
      planilhaAlocacao,
      `U1:V${valoresAcoes.length}`,
      43
    );
    adicionarPie(valoresFII, planilhaAlocacao, `S1:T${valoresFII.length}`, 63);
    adicionarPie(
      valoresClasses,
      planilhaAlocacao,
      `Q1:R${valoresClasses.length}`,
      83
    );
    adicionarPie(
      valoresTodosSetores,
      planilhaAlocacao,
      `O1:P${valoresTodosSetores.length}`,
      103
    );
    adicionarPie(
      valoresSetoresETF,
      planilhaAlocacao,
      `M1:N${valoresSetoresETF.length}`,
      123
    );
    adicionarPie(
      valoresSetoresAcoes,
      planilhaAlocacao,
      `K1:L${valoresSetoresAcoes.length}`,
      143
    );
    adicionarPie(
      valoresSetoresFII,
      planilhaAlocacao,
      `I1:J${valoresSetoresFII.length}`,
      163
    );
  }
  
  function getPlanilha(nome) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(nome);
    return sheet;
  }
  
  function limparCarteira() {
    var sheetDados = getPlanilha("Dados");
    var sheetCarteira = getPlanilha("Carteira");
  
    var linhas = sheetDados.getLastRow();
    for (var i = 8; i <= linhas; i++) {
      sheetDados.deleteRow(8);
    }
  
    linhas = sheetCarteira.getLastRow();
    for (var i = 8; i <= linhas; i++) {
      sheetCarteira.deleteRow(8);
    }
  }
  
  function adicionarProvento() {
    var sheetProventos = getPlanilha("Proventos");
    var planilhaEvolucao = getPlanilha("Evolução");
  
    if (sheetProventos.getLastRow() > 7) {
      var chts = planilhaEvolucao.getCharts();
      for (var i = 0; i < chts.length; i++) {
  
        if (chts[i].getOptions().get('title') === 'proventos')
          planilhaEvolucao.removeChart(chts[i]);
      }
  
      sheetProventos.getRange(1, 20, sheetProventos.getLastRow(), 2).clear();
    }
  
    sheetProventos
      .getRange(
        `B${sheetProventos.getLastRow() + 1}:E${sheetProventos.getLastRow() + 1}`
      )
      .setValues(sheetProventos.getRange("B3:E3").getValues())
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      )
      .setHorizontalAlignment("center");
  
    var dates = sheetProventos
      .getRange(`B8:E${sheetProventos.getLastRow()}`)
      .getValues()
      .map(function (r) {
        return r[2];
      })
      .sort(function (a, b) {
        return a - b;
      });
  
    var menorData = dates[0];
    var maiorData = dates[dates.length - 1];
  
    var linha = 1;
    for (var y = menorData.getFullYear(); y <= maiorData.getFullYear(); y++) {
      for (
        var m = y === menorData.getFullYear() ? menorData.getMonth() + 1 : 1;
        y === maiorData.getFullYear() ? m <= maiorData.getMonth() + 1 : m <= 12;
        m++
      ) {
        var proventosMes = sheetProventos
          .getRange(`B8:E${sheetProventos.getLastRow()}`)
          .getValues()
          .filter(function (r) {
            return r[2].getFullYear() === y && r[2].getMonth() + 1 === m;
          });
  
          if (proventosMes && proventosMes.length > 0) {
            proventosMes = proventosMes
            .map(function(r) {
              return r[1];
            })
            .reduce(function (acum, curr) {
              return acum + curr;
            });
  
            sheetProventos.getRange(linha, 20).setValue(`28/${m}/${y}`);
            sheetProventos.getRange(linha, 21).setValue(proventosMes);
  
            linha++;
          }
      }
    }
  
    var chart = planilhaEvolucao
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheetProventos.getRange(1, 20, linha, 2))
      .setPosition(24, 2, 0, 0)
      .setOption('title', 'proventos')
      .setOption("series", {
        0: { type: "bars", color: "blue", labelInLegend: "Proventos" },
      })
      .setOption("useFirstColumnAsDomain", true)
      .build();
  
    planilhaEvolucao.insertChart(chart);
  }
  
  function adicionarOrdem() {
    var sheetDados = getPlanilha("Dados");
    //var sheetCarteira = getPlanilha('Carteira');
  
    var range = sheetDados.getRange("B3:G3");
  
    // calcula valor total da ordem
    var ordem = range.getValues();
    ordem[0].splice(
      4,
      0,
      ordem[0][1] * ordem[0][2] +
        (ordem[0][1] < 0 ? -1 * ordem[0][3] : ordem[0][3])
    );
  
    var coordenadas = `B${sheetDados.getLastRow() + 1}:H${
      sheetDados.getLastRow() + 1
    }`;
    sheetDados
      .getRange(coordenadas)
      .setValues(ordem)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      )
      .setHorizontalAlignment("center");
  
    if (ordem[0][1] < 0)
      sheetDados.getRange(coordenadas).setBackground("#f4cccc");
    else sheetDados.getRange(coordenadas).setBackground("#cfebec");
  
    atualizarCarteira(ordem);
    calcularValorInvestido();
    atualizarPesos();
    adicionarPies();
    adicionarEvolucao();
    adicionarProvento();
  }
  
  function ordenarAtivos() {
    var sheetCarteira = getPlanilha("Carteira");
  
    if (sheetCarteira.getLastRow() === 7) return;
  
    // ordena pela percentagem de distancia do preço teto
    // sheetCarteira
    //   .getRange(`A8:Z${sheetCarteira.getLastRow()}`)
    //   .sort({ column: 25, ascending: true });
  
    // ordena pelos ativos mais longe da porcentagem alvo
    sheetCarteira
      .getRange(`A8:Z${sheetCarteira.getLastRow()}`)
      .sort({ column: 11, ascending: false });
  
    sheetCarteira
      .getRange(`J8:J${sheetCarteira.getLastRow()}`)
      .setBackground("#ffffff");
  
    var colors = ["#00ff00", "#22b60d", "#197906"];
    var lastRow = sheetCarteira.getLastRow();
  
    for (var i = 8; i <= lastRow; i++) {
      if (i === 11) break;
  
      sheetCarteira.getRange(`J${i}`).setBackground(colors[i - 8]);
    }
  
    //ordena pelo tipo do ativo
    sheetCarteira
      .getRange(`A8:Z${sheetCarteira.getLastRow()}`)
      .sort({ column: 26, ascending: true });
  }
  
  function atualizarCarteira(aporte) {
    var sheetCarteira = getPlanilha("Carteira");
  
    var ticker = aporte[0][0];
    // var quantidade = aporte[0][1];
    // var preco = aporte[0][2];
    // var custos = aporte[0][3];
    // var total = aporte[0][4];
    var classe = aporte[0][6];
  
    var novoAtivo = true;
  
    for (var i = 8; i < sheetCarteira.getLastRow() + 2; i++) {
      if (ticker === sheetCarteira.getRange(`B${i}`).getValue()) {
        var peso = sheetCarteira.getRange(`J${i}`).getValue();
        var setor = sheetCarteira.getRange(`N${i}`).getValue();
        atualizarLinhaCarteira(sheetCarteira, i, ticker, classe, peso, setor);
        novoAtivo = false;
      }
    }
  
    if (novoAtivo)
      atualizarLinhaCarteira(
        sheetCarteira,
        sheetCarteira.getLastRow() + 1,
        ticker,
        classe,
        10,
        ""
      );
  
    // seta formula para calcular valor de mercado
    sheetCarteira
      .getRange("D2")
      .setFormula(`=SUM(E8:E${sheetCarteira.getLastRow()})`);
  }
  
  function atualizarLinhaCarteira(
    sheetCarteira,
    linha,
    ticker,
    classe,
    peso,
    setor
  ) {
    var sheetDados = getPlanilha("Dados");
  
    var posicaoAtivo = calculaPosicaoAtivo(ticker, sheetDados);
    var valorInvestidoAtivo = posicaoAtivo[0];
    var quantidadeAtivo = posicaoAtivo[1];
  
    var valores = [
      [ticker, quantidadeAtivo, 0, 0, 0, 0, 0, 0, peso, 0, 0, 0, setor],
    ];
    sheetCarteira
      .getRange(`B${linha}:N${linha}`)
      .setValues(valores)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "black",
        SpreadsheetApp.BorderStyle.SOLID
      )
      .setBackground("#fff2cc")
      .setHorizontalAlignment("center");
  
    sheetCarteira.getRange(`D${linha}`).setFormula(`=GOOGLEFINANCE(B${linha})`);
    sheetCarteira.getRange(`E${linha}`).setFormula(`=C${linha}*D${linha}`);
    sheetCarteira.getRange(`F${linha}`).setFormula(`=E${linha} / $D$2`);
  
    sheetCarteira
      .getRange(`G${linha}`)
      .setFormula(
        `=E${linha} - ${valorInvestidoAtivo.toString().replace(".", ",")}`
      );
  
    sheetCarteira
      .getRange(`H${linha}`)
      .setFormula(
        `=G${linha} / ${valorInvestidoAtivo.toString().replace(".", ",")}`
      );
  
    sheetCarteira
      .getRange(`Z${linha}`)
      .setValue(classe === "FII" ? 2 : classe === "AÇÃO" ? 1 : 0);
  
    if (classe === "FII")
      sheetCarteira.getRange(`B${linha}:N${linha}`).setBackground("#ecd6d7");
    else if (classe === "AÇÃO")
      sheetCarteira.getRange(`B${linha}:N${linha}`).setBackground("#cfe9e9");
    else sheetCarteira.getRange(`B${linha}:N${linha}`).setBackground("#d6ecde");
  
    setSetoresDropdown(linha, classe);
    setPesosDropdown(linha);
  }
  
  function calcularValorInvestido() {
    var sheetCarteira = getPlanilha("Carteira");
    var sheetDados = getPlanilha("Dados");
  
    var valorInvestidoTotal = 0;
    var valorTotalEmVendas = 0;
    var valorTotalInvestidoVendas = 0;
  
    for (var i = 8; i < sheetCarteira.getLastRow() + 1; i++) {
      var ticker = sheetCarteira.getRange(`B${i}`).getValue();
  
      var posicaoAtivo = calculaPosicaoAtivo(ticker, sheetDados);
      var valorInvestidoAtivo = posicaoAtivo[0];
      var quantidadeInvestidoAtivo = posicaoAtivo[1];
      var valorEmVendasAtivo = posicaoAtivo[2];
      var valorInvestidoEmVendasAtivo = posicaoAtivo[3];
  
      valorInvestidoTotal += valorInvestidoAtivo;
      valorTotalEmVendas += valorEmVendasAtivo;
      valorTotalInvestidoVendas += valorInvestidoEmVendasAtivo;
  
      if (quantidadeInvestidoAtivo === 0) {
        sheetCarteira.deleteRow(i);
      }
    }
  
    sheetCarteira.getRange("D1").setValue(valorInvestidoTotal);
    sheetCarteira.getRange("K2").setValue(valorTotalEmVendas);
    sheetCarteira
      .getRange("L2")
      .setValue(
        valorTotalInvestidoVendas > 0
          ? valorTotalEmVendas / valorTotalInvestidoVendas
          : 0
      );
  }
  
  function calculaPosicaoAtivo(ticker, sheetDados) {
    var aportes = sheetDados
      .getRange(`B8:H${sheetDados.getLastRow()}`)
      .getValues()
      .filter(function (l) {
        return l[0] === ticker;
      });
  
    var valorTotalInvestido = 0;
    var quantidadeTotal = 0;
    var lucroTotalVendas = 0;
    var valorTotalInvestidoVendas = 0;
  
    for (var i = 0; i < aportes.length; i++) {
      var quantidade = aportes[i][1];
      var total = aportes[i][4];
  
      if (quantidade > 0) valorTotalInvestido += total;
      else {
        const novoValorTotalInvestido =
          (valorTotalInvestido * (quantidadeTotal + quantidade)) /
          quantidadeTotal;
        const valorInvestidoVenda = valorTotalInvestido - novoValorTotalInvestido;
        valorTotalInvestidoVendas += valorInvestidoVenda;
  
        lucroTotalVendas += total * -1 - valorInvestidoVenda;
  
        if (quantidadeTotal + quantidade > 0) {
          valorTotalInvestido = novoValorTotalInvestido;
        } else {
          valorTotalInvestido = 0;
        }
      }
      quantidadeTotal += quantidade;
    }
    return [
      valorTotalInvestido,
      quantidadeTotal,
      lucroTotalVendas,
      valorTotalInvestidoVendas,
    ];
  }
  
  function setClassesDropdown() {
    var sheetDados = getPlanilha("Dados");
    var range = sheetDados.getRange("G3");
  
    var arrayValues = [["AÇÃO"], ["FII"], ["ETF"]];
  
    // define the dropdown/validation rules
    var rangeRule = SpreadsheetApp.newDataValidation().requireValueInList(
      arrayValues
    );
  
    // set the dropdown validation for the row
    range.setDataValidation(rangeRule); // set range to your range
  }
  
  function setTiposProventosDropdown() {
    var sheetDados = getPlanilha("Proventos");
    var range = sheetDados.getRange("E3");
  
    var arrayValues = [["Dividendo"], ["JCP"], ["Outros"]];
  
    // define the dropdown/validation rules
    var rangeRule = SpreadsheetApp.newDataValidation().requireValueInList(
      arrayValues
    );
  
    // set the dropdown validation for the row
    range.setDataValidation(rangeRule); // set range to your range
  }
  
  function atualizarPesos() {
    var sheetCarteira = getPlanilha("Carteira");
  
    if (sheetCarteira.getLastRow() === 7) return;
  
    var rangeCarteira = sheetCarteira.getRange(
      `J8:J${sheetCarteira.getLastRow()}`
    );
  
    var pesos = rangeCarteira.getValues();
    var numeroAtivos = pesos.length;
  
    //Browser.msgBox(sheetCarteira.getLastRow());
  
    var percentagemAtivo = 100 / numeroAtivos / 100;
  
    var percentagemAlvoCadaAtivo = pesos.map(function (l) {
      return (percentagemAtivo * l[0]) / 10;
    });
  
    var percentagemTotal = percentagemAlvoCadaAtivo.reduce(function (acum, curr) {
      return acum + curr;
    });
  
    if (percentagemTotal < 1) {
      var percentagemFaltanteCadaAtivo = (1 - percentagemTotal) / numeroAtivos;
  
      percentagemAlvoCadaAtivo = percentagemAlvoCadaAtivo.map(function (perc) {
        return perc + percentagemFaltanteCadaAtivo;
      });
    }
  
    var percentagemAtualCadaAtivo = sheetCarteira
      .getRange(`F8:F${sheetCarteira.getLastRow()}`)
      .getValues();
  
    var percentagemFaltante = percentagemAtualCadaAtivo.map(function (
      atual,
      index
    ) {
      return percentagemAlvoCadaAtivo[index] - atual[0];
    });
  
    sheetCarteira
      .getRange(`K8:K${sheetCarteira.getLastRow()}`)
      .setNumberFormat("##.#%")
      .setValues(
        percentagemFaltante.map(function (l) {
          return [l];
        })
      );
  
    var totalMercado = sheetCarteira.getRange(`D${2}`).getValue();
    var valorFaltante = percentagemFaltante.map(function (p) {
      return totalMercado * p;
    });
  
    var precos = sheetCarteira
      .getRange(`D8:D${sheetCarteira.getLastRow()}`)
      .getValues();
    var quantidadeFaltante = precos.map(function (p, index) {
      return Math.floor(valorFaltante[index] / p[0]);
    });
  
    sheetCarteira.getRange(`L8:L${sheetCarteira.getLastRow()}`).setValues(
      quantidadeFaltante.map(function (q) {
        return [q];
      })
    );
  
    sheetCarteira.getRange(`M8:M${sheetCarteira.getLastRow()}`).setValues(
      valorFaltante.map(function (v) {
        return [v];
      })
    );
  
    ordenarAtivos();
  }
  
  function setSetoresDropdown(linha, classeAtivo) {
    var sheetCarteira = getPlanilha("Carteira");
    var rangeCarteira = sheetCarteira.getRange(`N${linha}`);
  
    var sheetAux = getPlanilha("Aux");
    const coluna =
      classeAtivo === "AÇÃO" ? "A" : classeAtivo === "ETF" ? "B" : "C";
    var setores = sheetAux
      .getRange(`${coluna}${3}:${coluna}${sheetAux.getLastRow()}`)
      .getValues()
      .filter(function (l) {
        return l[0] !== null && l[0] !== undefined;
      });
  
    // define the dropdown/validation rules
    var rangeRule = SpreadsheetApp.newDataValidation().requireValueInList(
      setores
    );
  
    // set the dropdown validation for the row
    rangeCarteira.setDataValidation(rangeRule); // set range to your range
  }
  
  function setPesosDropdown(linha) {
    var sheetCarteira = getPlanilha("Carteira");
    var rangeCarteira = sheetCarteira.getRange(`J${linha}`);
  
    var pesos = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10];
  
    // define the dropdown/validation rules
    var rangeRule = SpreadsheetApp.newDataValidation().requireValueInList(pesos);
  
    // set the dropdown validation for the row
    rangeCarteira.setDataValidation(rangeRule); // set range to your range
  }
  
  