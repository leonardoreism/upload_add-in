/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "block";
    document.getElementById("salvar").onclick = run;
  }
});

function montarTexto() {

    console.log("MontarTexto iniciou");

    let texto = "";

    texto += document.getElementById("problematizacao").checked
        ? "☑ Problematização\n"
        : "☐ Problematização\n";

    texto += document.getElementById("desenvolvimento").checked
        ? "☑ Desenvolvimento\n"
        : "☐ Desenvolvimento\n";

    texto += document.getElementById("sintese").checked
        ? "☑ Síntese\n"
        : "☐ Síntese\n";

    texto += document.getElementById("concluido").checked
        ? "☑ Concluído"
        : "☐ Concluído";

    return texto;
}
  
export async function run() {
  try {
    console.log("Run iniciou");
    await Excel.run(async (context) => {

      const range = context.workbook.getSelectedRange();

      range.values = [[montarTexto()]];

      range.format.wrapText = true;

      await context.sync();

    });

  } catch (error) {
    console.error(error);
  }
}
