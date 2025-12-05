/* global document, Office, Word */

let lastGeneratedText = "";
let lastOperationType = "";

// --------------------------------------------------
// INIT OFFICE
// --------------------------------------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    const sendBtn = document.getElementById("hexabot-send");
    const input = document.getElementById("hexabot-text");

    sendBtn.onclick = sendHexabotMessage;
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendHexabotMessage();
      }
    });
  }
});

// --------------------------------------------------
// OLLAMA CALLER
// --------------------------------------------------
async function callOllama(prompt) {
  const res = await fetch("http://172.30.5.220:11434/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "llama3", prompt, stream: false })
  });

  const data = await res.json();
  return data.response.trim();
}

// --------------------------------------------------
// INTENT DETECTOR (amichevole ma preciso)
// --------------------------------------------------
function detectIntent(text) {
  const t = text.toLowerCase();

  const map = {
    riassunto: ["riassumi", "riepiloga", "riassunto"],
    migliorato: ["migliora", "riscrivi", "riformula"],
    corto: ["accorcia", "più breve", "piu breve"],
    formale: ["più formale", "piu formale"],
    informale: ["più informale", "piu informale"],
    traduzione: ["traduci", "traduzione"],
    spiegazione: ["spiega", "spiegami", "interpretami"],
    correzione: ["correggi", "errori", "correzione"],
    valutazione: ["valuta", "giudica", "dai un voto"]
  };

  for (const intent in map) {
    if (map[intent].some((k) => t.includes(k))) return intent;
  }

  return null;
}

// --------------------------------------------------
// MAIN CHAT ENTRY
// --------------------------------------------------
async function sendHexabotMessage() {
  const input = document.getElementById("hexabot-text");
  const text = input.value.trim();
  if (!text) return;

  appendHexabotMessage(text, "user");
  input.value = "";
  input.disabled = true;

  try {
    const intent = detectIntent(text);

    if (!intent) {
      appendHexabotMessage(
        "Posso aiutarti solo sul testo selezionato. Se vuoi riassumere, migliorare, tradurre o correggere qualcosa, seleziona un testo nel documento e chiedimelo 😊",
        "ai"
      );
      return;
    }

    await runAction(intent);
  } catch (e) {
    appendHexabotMessage("Errore: " + e.message, "ai");
  } finally {
    input.disabled = false;
    input.focus();
  }
}

// --------------------------------------------------
// EXECUTE THE OPERATION
// --------------------------------------------------
async function runAction(intent) {
  const overlay = document.getElementById("loading-overlay");
  overlay.style.display = "block";

  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();

      if (!sel.text || !sel.text.trim()) {
        appendHexabotMessage(
          "Per procedere, seleziona prima un testo nel documento ✏️",
          "ai"
        );
        return;
      }

      const original = sel.text;
      let prompt = "";

      switch (intent) {
        case "riassunto":
          prompt = `
            Riassumi in italiano il seguente testo in modo chiaro, breve e professionale.
            Non aggiungere nulla, non commentare.
            Testo:
            ${original}
          `;
          break;

        case "migliorato":
          prompt = `
            Riscrivi il testo in modo più chiaro, professionale e corretto.
            Niente aggiunte, niente interpretazioni.
            Testo:
            ${original}
          `;
          break;

        case "corto":
          prompt = `
            Accorcia questo testo mantenendo intatto il significato.
            Italiano, stile chiaro e naturale.
            Testo:
            ${original}
          `;
          break;

        case "formale":
          prompt = `
            Riscrivi il testo rendendolo più formale, elegante e professionale.
            Testo:
            ${original}
          `;
          break;

        case "informale":
          prompt = `
            Riscrivi questo testo rendendolo più informale, semplice e colloquiale.
            Testo:
            ${original}
          `;
          break;

        case "traduzione":
          prompt = `
            Traduci il testo nella lingua richiesta dall'utente.
            Non aggiungere commenti.
            Testo:
            ${original}
          `;
          break;

        case "sipegazione":
          prompt = `
            Spiega in modo semplice e chiaro il contenuto del testo.
            Italiano.
            Testo:
            ${original}
          `;
          break;

        case "correzione":
          prompt = `
            Correggi tutti gli errori grammaticali e sintattici, mantenendo stile e significato.
            Testo:
            ${original}
          `;
          break;

        case "valutazione":
          prompt = `
            Valuta la qualità del testo (0-10) dando un giudizio breve e oggettivo in italiano.
            Testo:
            ${original}
          `;
          break;
      }

      const result = await callOllama(prompt);
      lastGeneratedText = result;
      lastOperationType = intent;

      sel.insertComment(`[OfficeAI – ${intent}]\n${result}`);
      await context.sync();

      appendHexabotMessage(
        `Ho creato un nuovo commento.\nVuoi fare altre modifiche?`,
        "ai"
      );
    });
  } finally {
    overlay.style.display = "none";
  }
}

// --------------------------------------------------
// CHAT UI HELPERS
// --------------------------------------------------
function appendHexabotMessage(text, type) {
  const box = document.getElementById("hexabot-messages");
  const msg = document.createElement("div");
  msg.className = type === "user" ? "hexabot-user" : "hexabot-ai";
  msg.innerText = text;
  box.appendChild(msg);
  box.scrollTop = box.scrollHeight;
}
