/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Leggo solo l'indirizzo (come nell'originale)
      range.load("address");

      await context.sync();

      // Colore giallo (come l'originale)
      range.format.fill.color = "yellow";

      // 🔽 NUOVA PARTE: ordina alfabeticamente il range selezionato (A→Z)
      range.sort.apply([
        { key: 0, ascending: true }
      ]);

      await context.sync();

      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}


