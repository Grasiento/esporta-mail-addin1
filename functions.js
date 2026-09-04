/*
 * Logica portata da Modulo1.bas (macro VBA "EsportaMailComeMSG").
 * Differenze rispetto alla versione VBA:
 *  - il file esportato e' in formato .eml (MIME standard) e non .msg,
 *    perche' le API disponibili per un add-in restituiscono il contenuto MIME;
 *  - il salvataggio avviene tramite download del browser (cartella Download
 *    predefinita), non su un percorso fisso del disco come nella macro.
 */

Office.onReady();

// Associa la funzione al pulsante definito nel manifest (FunctionName: esportaEmailAction)
Office.actions.associate("esportaEmailAction", esportaEmailAction);

function esportaEmailAction(event) {
  const item = Office.context.mailbox.item;

  try {
    esportaEmail(item, event);
  } catch (e) {
    mostraNotifica(item, "Errore durante l'esportazione: " + e.message);
    event.completed();
  }
}

function esportaEmail(item, event) {
  const dataMail = formattaData(item.dateTimeCreated);

  const mittente = (item.from && item.from.emailAddress) || "";
  const azienda = trovaAziendaDaMail(mittente, item.sender && item.sender.displayName);

  let oggetto = pulisciNomeFile(item.subject || "(nessun oggetto)");
  if (oggetto.length > 70) {
    oggetto = oggetto.substring(0, 70);
  }

  const fileName = dataMail + " - " + azienda + " - " + oggetto + ".eml";

  const restId = Office.context.mailbox.convertToRestId(
    item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      mostraNotifica(item, "Impossibile ottenere il token di accesso.");
      event.completed();
      return;
    }

    const accessToken = result.value;
    const restUrl = Office.context.mailbox.restUrl + "/v2.0/me/messages/" + restId + "/$value";

    fetch(restUrl, {
      headers: { Authorization: "Bearer " + accessToken }
    })
      .then(function (response) {
        if (!response.ok) {
          throw new Error("Errore HTTP " + response.status);
        }
        return response.blob();
      })
      .then(function (blob) {
        scaricaFile(blob, fileName);
        mostraNotifica(item, "Mail esportata: " + fileName);
        event.completed();
      })
      .catch(function (err) {
        mostraNotifica(item, "Errore durante l'esportazione: " + err.message);
        event.completed();
      });
  });
}

// Equivalente di TrovaAziendaDaMail / OttieniSMTP in Modulo1.bas
function trovaAziendaDaMail(emailMittente, nomeMittenteFallback) {
  let azienda;
  const smtp = (emailMittente || "").toLowerCase();

  if (smtp.indexOf("@") > -1) {
    const dominio = smtp.split("@")[1];
    const parti = dominio.split(".");

    if (parti.length >= 2) {
      azienda = parti[parti.length - 2];
    } else {
      azienda = parti[0];
    }
  } else {
    azienda = nomeMittenteFallback || "Sconosciuto";
  }

  switch (azienda.toLowerCase()) {
    case "abb":
      azienda = "ABB";
      break;
    case "unitechpackaging":
    case "unitechcpackaging":
    case "unitech":
      azienda = "Unitech";
      break;
    default:
      azienda = azienda.charAt(0).toUpperCase() + azienda.slice(1);
  }

  return pulisciNomeFile(azienda);
}

// Equivalente di PulisciNomeFile in Modulo1.bas
function pulisciNomeFile(testo) {
  return String(testo)
    .replace(/[\\/:*?"<>|]/g, "")
    .trim();
}

// Equivalente di Format(objMail.ReceivedTime, "yyyy - mm - dd")
function formattaData(date) {
  const d = date instanceof Date ? date : new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const g = String(d.getDate()).padStart(2, "0");
  return y + " - " + m + " - " + g;
}

function scaricaFile(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function mostraNotifica(item, messaggio) {
  item.notificationMessages.replaceAsync("esportaStatus", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: messaggio,
    icon: "Icon.16x16",
    persistent: false
  });
}
