/* global document, Office, Word */

let lastGeneratedText = ""; 

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 👉 Il bottone ora lancia il riepilogo
    document.getElementById("run").onclick = summarizeSelection;
    document.getElementById("improveText").onclick = improveSelectedText;

  }
});

/**
 * 🔵 Funzione che chiama Ollama localmente
 */
async function callOllama(prompt) {
  const response = await fetch("http://localhost:11434/api/generate", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "llama3",
      prompt: prompt,
      stream: false
    })
  });

  const data = await response.json();
  return data.response;
}

/**
 * 🟢 Funzione principale: riassume ciò che l’utente ha selezionato in Word
 */
async function summarizeSelection() {
  try {
    // Mostra overlay
    document.getElementById("loading-overlay").style.display = "block";

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text || selection.text.trim() === "") {
        alert("Seleziona un testo da riepilogare.");
        document.getElementById("loading-overlay").style.display = "none";
        return;
      }

      const prompt = `
      Rispondi obbligatoriamente in italiano.
      Riassumi senza aggiungere considerazioni.
      Valuta correttamente di non tralasciare dettagli importanti.
      Riassumi senza errori grammaticali, in modo chiaro conciso e professionale il seguente testo:
      ${selection.text}
      `;

      const summary = await callOllama(prompt);

     /* selection.insertText(
        "\n\n[📌 Sintesi AI]\n" + summary,
        Word.InsertLocation.after
      );*/

     lastGeneratedText = summary;

// commento nel documento Word
      selection.insertComment("[OfficeAI]\n" + summary);

// chat: messaggio utente + risposta AI
        appendHexabotMessage("Ho generato il riepilogo. Vuoi che lo migliori?", "ai");


// FA LA DOMANDA AUTOMATICA
        askForImprovements();




    });

    

    // ORDINA COLONNE A → Z
function sortAlphabetically(event) {
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("columnIndex");
    range.load("worksheet");

    await context.sync();

    const sheet = range.worksheet;
    const columnIndex = range.columnIndex;

    const used = sheet.getUsedRange();
    used.load("rowCount");

    await context.sync();

    const columnRange = sheet.getRangeByIndexes(
      0, columnIndex,
      used.rowCount,
      1
    );

    columnRange.sort.apply([
      { key: 0, ascending: true }
    ]);

    await context.sync();
    event.completed();
  });
}


// SPEZZA UNA FRASE NELLE COLONNE A DESTRA
function splitSentence(event) {
  Excel.run(async (context) => {
    const sentence = prompt("Inserisci la frase da spezzare:");

    const words = sentence.split(" ");
    const wordsPerColumn = 3; // puoi personalizzare

    let chunks = [];

    for (let i = 0; i < words.length; i += wordsPerColumn) {
      chunks.push(words.slice(i, i + wordsPerColumn).join(" "));
    }

    const range = context.workbook.getSelectedRange();
    const sheet = range.worksheet;

    range.load(["columnIndex", "rowIndex"]);
    await context.sync();

    for (let i = 0; i < chunks.length; i++) {
      const cell = sheet.getCell(range.rowIndex, range.columnIndex + i);
      cell.values = [[chunks[i]]];
    }

    await context.sync();
    event.completed();
  });
}

  } catch (error) {
    alert("Errore: " + error.message);
  } finally {
    // Nasconde overlay SEMPRE
    document.getElementById("loading-overlay").style.display = "none";
  }
}

async function improveSelectedText() {
  try {
    // Mostra overlay (uguale a summarize)
    document.getElementById("loading-overlay").style.display = "block";

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      if (!selection.text || selection.text.trim() === "") {
        alert("Seleziona un testo da migliorare.");
        document.getElementById("loading-overlay").style.display = "none";
        return;
      }

      const originalText = selection.text;

      const prompt = `
      Riscrivi il testo in modo formale e professionale con uno stile chiaro e moderno.
      Correggi eventuali errori grammaticali o lessicali.
      Mantieni esattamente lo stesso significato senza aggiungere informazioni.
      Non usare espressioni colloquiali, poco comuni o ridondanti.
      Evita frasi troppo lunghe, ridondanti o complesse.
      Usa un lessico adatto ad una comunicazione aziendale.
      Mantieni una lunghezza simile all’originale.
      Attieniti alla lingua del testo selezionato.
      EVITA LE RIPETIZIONI O PAROLE ASSONANTI.
      USA SINONONIMI SE UNA PAROLA È TROPPO SIMILE AD UN ALTRA.

      ❗ PRODUCI SOLO IL TESTO RISCRITTO.
      ❗ NON AZZARDARTI A SCRIVERLO IN INGLESE SE IL TESTO è IN ITALIANO.
      ❗ NON INSERIRE SPIEGAZIONI, COMMENTI O FRASI AGGIUNTIVE.

      Testo:
      ${originalText}
      `;

      const improvedText = await callOllama(prompt);

      // Inserisce il testo migliorato come COMMENTO (come summarize)
      selection.insertComment("[OfficeAI – Miglioramento]\n" + improvedText.trim());

      lastGeneratedText = improvedText;

        appendHexabotMessage("Testo migliorato. Vuoi che lo modifichi ulteriormente?", "ai");

        askForImprovements();


      await context.sync();
    });

  } catch (error) {
    console.error("Errore durante la generazione:", error);
    alert("Errore: " + error.message);
  } finally {
    // Nasconde overlay SEMPRE
    document.getElementById("loading-overlay").style.display = "none";
  }
}

function askForImprovements() {
    appendHexabotMessage(
        "Vuoi che lo migliori ulteriormente, lo accorci, lo renda più formale o più semplice?",
        "ai"
    );
}


document.getElementById("hexabot-send").onclick = sendHexabotMessage;

// INVIO invia – SHIFT+INVIO va a capo
document.getElementById("hexabot-text").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
        if (e.shiftKey) return; 
        e.preventDefault();
        sendHexabotMessage();
    }
});

async function sendHexabotMessage() {
    const textInput = document.getElementById("hexabot-text");
    const text = textInput.value.trim();
    if (!text) return;

    // Mostra messaggio utente
    appendHexabotMessage(text, "user");
    textInput.value = "";
    textInput.disabled = true;

       // Caso: l’utente risponde alla domanda “vuoi migliorare?”
    if (lastGeneratedText && isImprovementRequest(text)) {
        await regenerateImprovedVersion(text);
        textInput.disabled = false;
        return;
    }

    // Mostra indicatore “sta scrivendo…”
    showHexabotTyping();

    const res = await fetch("http://localhost:11434/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            model: "llama3",
            prompt: text,
            stream: true
        })
    });


    // Rimuove indicatore
    removeHexabotTyping();

    // Risposta AI in streaming
    await streamHexabotResponse(res);

    textInput.disabled = false;
    textInput.focus();
}

function appendHexabotMessage(text, type) {
    const msg = document.createElement("div");
    msg.className = type === "user" ? "hexabot-user" : "hexabot-ai";
    msg.innerText = text;

    const container = document.getElementById("hexabot-messages");
    container.appendChild(msg);
    container.scrollTop = container.scrollHeight;
}

function showHexabotTyping() {
    const msg = document.createElement("div");
    msg.id = "hexabot-typing";
    msg.className = "hexabot-ai";
    msg.innerHTML = "Hexabot sta scrivendo<span class='typing-dots'>...</span>";

    document.getElementById("hexabot-messages").appendChild(msg);
    scrollHexabot();
}

function removeHexabotTyping() {
    const el = document.getElementById("hexabot-typing");
    if (el) el.remove();
    scrollHexabot();
}


    function isImprovementRequest(text) {
    const keywords = ["sì", "si", "migliora", "miglioralo", "modifica", "accorcia", "più formale", "più semplice"];
    return keywords.some(k => text.toLowerCase().includes(k));
}


function scrollHexabot() {
    const container = document.getElementById("hexabot-messages");
    container.scrollTop = container.scrollHeight;
}

// TYPING EFFECT in streaming da Ollama
async function streamHexabotResponse(res) {
    const reader = res.body.getReader();
    let aiMsg = document.createElement("div");
    aiMsg.className = "hexabot-ai";
    aiMsg.innerText = "";
    document.getElementById("hexabot-messages").appendChild(aiMsg);

    const decoder = new TextDecoder();

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        const chunk = decoder.decode(value);
        const lines = chunk.split("\n");

        for (const line of lines) {
            if (line.trim().startsWith("{")) {
                try {
                    const json = JSON.parse(line.trim());
                    if (json.response) {
                        aiMsg.innerText += json.response;
                        scrollHexabot();
                    }
                } catch (e) {}
            }
        }
    }
}

async function regenerateImprovedVersion(requestType) {
  try {
    // Mostra overlay durante la rigenerazione
    document.getElementById("loading-overlay").style.display = "block";

    const prompt = `
      Migliora il seguente testo secondo la richiesta dell'utente:
      Richiesta: ${requestType}
      Testo: ${lastGeneratedText}

      Regole:
      - mantieni il significato
      - non inventare contenuti
      - tono chiaro e professionale
      - massimo 2 paragrafi
    `;

    const res = await fetch("http://localhost:11434/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "llama3", prompt, stream: false })
    });

    const data = await res.json();
    const improved = data.response.trim();

    // aggiorno il "testo corrente" su cui lavorare
    lastGeneratedText = improved;

    // conferma nella chat (senza incollare il testo)
    appendHexabotMessage("Fatto! Ho applicato la modifica richiesta e creato un nuovo commento in Word.", "ai");

    // nuovo commento nel documento
    await Word.run(async (context) => {
      const range = context.document.getSelection(); // stesso punto del testo originale
      range.insertComment("[OfficeAI – Aggiornamento]\n" + improved);
      await context.sync();
    });

    // chiede se vuoi continuare a rifinirlo
    askForImprovements();

  } catch (e) {
    appendHexabotMessage("Errore durante la modifica: " + e.message, "ai");
  } finally {
    document.getElementById("loading-overlay").style.display = "none";
  }
}
