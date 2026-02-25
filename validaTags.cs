function sincronizarUmblerTalkFinal() {
  const API_TOKEN = ''; 
  const ORG_ID = '';
  const ID_TAG_SEM_RESPOSTA = '';
  const ID_TAG_RESPONDEU_CS = '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SECA30OFICIAL_test");
  if (!sheet) return;

  const mapaContatos = {};
  let skip = 0;
  const take = 250; 
  let buscarMais = true;

  try {
    while (buscarMais) {
      
      const url = `https://app-utalk.umbler.com/api/v1/contacts/?organizationId=${ORG_ID}&Skip=${skip}&Take=${take}&Behavior=GetSliceOnly&Tags.Rule=ContainsAny&Tags.Values=${ID_TAG_SEM_RESPOSTA}&Tags.Values=${ID_TAG_RESPONDEU_CS}&State=Active`;
      
      const options = {
        method: 'get',
        headers: { 'Authorization': `Bearer ${API_TOKEN}`, 'Accept': 'application/json' },
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      const items = json.items || [];

      if (items.length === 0) {
        buscarMais = false;
      } else {
        items.forEach(contato => {
          let num = (contato.phoneNumber || "").replace(/\D/g, "");
          if (!num) return;

          const tagsDoContato = (contato.tags || []).map(t => t.id);
          
          if (tagsDoContato.includes(ID_TAG_RESPONDEU_CS)) {
            mapaContatos[num] = false; 
          } else if (tagsDoContato.includes(ID_TAG_SEM_RESPOSTA)) {
            mapaContatos[num] = true;
          }
        });

        console.log(`Página processada: Skip ${skip}. Contatos acumulados: ${Object.keys(mapaContatos).length}`);
        
        skip += take;
        if (items.length < take) buscarMais = false; // Se veio menos que 250, é a última página
        
        // Pausa de segurança para não ser bloqueado por excesso de tráfego
        Utilities.sleep(500); 
      }
    }

    // --- APLICAÇÃO NA PLANILHA ---
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; 
    
    const rangeL = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
    const rangeW = sheet.getRange(2, 23, lastRow - 1, 1);
    const valoresW = rangeW.getValues(); 

    let marcados = 0;
    let desmarcados = 0;

    for (let i = 0; i < rangeL.length; i++) {
      let numP = rangeL[i][0].toString().replace(/\D/g, ''); 
      if (!numP) continue;

      let status = undefined;
      // Checa número puro, com 55 ou removendo 55
      if (mapaContatos.hasOwnProperty(numP)) status = mapaContatos[numP];
      else if (mapaContatos.hasOwnProperty("55" + numP)) status = mapaContatos["55" + numP];
      else if (numP.startsWith("55") && mapaContatos.hasOwnProperty(numP.substring(2))) status = mapaContatos[numP.substring(2)];

      if (status !== undefined) {
        valoresW[i][0] = status;
        status ? marcados++ : desmarcados++;
      }
    }

    rangeW.setValues(valoresW);

    console.log("=== RELATÓRIO FINAL COMPLETO ===");
    console.log("Total de contatos mapeados na Umbler (todas as páginas): " + Object.keys(mapaContatos).length);
    console.log("✅ Checkboxes Marcadas na Planilha: " + marcados);
    console.log("❌ Checkboxes Desmarcadas na Planilha: " + desmarcados);

  } catch (e) {
    console.error("Erro: " + e.toString());
  }
}
