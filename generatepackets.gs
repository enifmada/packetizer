function generateQuizBowlPackets() {
  // Get the active spreadsheet
  var category_id_dict = {"History": "https://docs.google.com/document/d/1YHxCwYBoccp3w9fmiN98DpDd61YiP3KC1tg86dt3jmo/edit",
  "Literature": "https://docs.google.com/document/d/1YHxCwYBoccp3w9fmiN98DpDd61YiP3KC1tg86dt3jmo/edit",
  "Science": "https://docs.google.com/document/d/1AqY1XK6J5JkpCILR7Qcj3BNtP1DGLS15sAyu8BdUNVM/edit",
  "Arts": "https://docs.google.com/document/d/1KSaFsbE10DqmuRJwXL5vuOk1Deg_OzeD0HmZaD6-DXs/edit",
  "Thought": "https://docs.google.com/document/d/1sEkC0pFniPyMrfrzIz9fWgV1ulNJxpKIz8PL_9YQQhI/edit",
  "Other": "https://docs.google.com/document/d/1wMHAKr75PFoqWAyXiRCR13RFFB6faxv1RkgUj6N30fA/edit",}

  var tu_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Packets TU");
  var bo_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Packets Bonus");
  
  // Get the range of data to be used
  var dataRange = tu_sheet.getDataRange();
  var bodataRange = bo_sheet.getDataRange();
  
  // Get the values in the range
  var data = dataRange.getValues();
  var bodata = bodataRange.getValues();
  
  // Get the number of rounds and categories
  var numRounds = data[0].length-1;
  var numCategories = data.length - 2;
  
  // Create an empty array to hold the packets
  var packets = [];
  
  // Loop through each round
  for (var round = 1; round <= numRounds; round++) {
    var packetDoc = DocumentApp.create("CMST II - Round " + round);
    var packetBody = packetDoc.getBody();
    var tournament_header = "2022 CMST II";
    var packet_header = "Packet " + round;
    var editors_header = "Edited by John Marvin and some other folks\n";
    tournament_par = packetBody.appendParagraph(tournament_header);
    tournament_par.setFontFamily("Times New Roman");
    tournament_par.setFontSize(12);
    tournament_par.setBold(true);
    tournament_par.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    packet_par = packetBody.appendParagraph(packet_header);
    packet_par.setBold(false);
    packet_par.setFontSize(10);
    packetBody.appendParagraph(editors_header);

    var allfound = true;
    bo_sheet.getRange(3,round+1, numCategories-3).setBackground("880000");
    tu_sheet.getRange(3,round+1,numCategories-3).setBackground("880000");

    orderedCats = [];
    for (var vi = 1; vi <= numCategories; vi++)
    {
      orderedCats.push(data[vi+1][0].trim());
    }

    var ordered_copy = [...orderedCats];
    tus = generateValidTossups(ordered_copy);
    bos = generateValidBonuses(tus);

    // Loop through each category
    for (var tossup = 0; tossup < numCategories; tossup++) {
      // Get the document for this category
      var indexTU = orderedCats.indexOf(tus[tossup]);
      var indexBO = orderedCats.indexOf(bos[tossup]);
      var categoryNameTU = data[indexTU+2][0];
      var metaCategoryTU = categoryNameTU.split("-")[0].trim();
      var baseAnswerTU = data[indexTU+2][round].split("!")[0].split("{")[0].trim();
      var capsAnswerTU = "";
      var splitanswer = baseAnswerTU.split(" ");
      for (wordi = 0; wordi < splitanswer.length; wordi++){
        capsAnswerTU += splitanswer[wordi].substr(0,1).toUpperCase() + splitanswer[wordi].substr(1) + " ";
      }
      capsAnswerTU = capsAnswerTU.trim();
      var categoryAnswerTU = "ANSWER: " + baseAnswerTU;
      var capsAnswerTU = "ANSWER: " + capsAnswerTU;
      var questionDocTU = DocumentApp.openByUrl(category_id_dict[metaCategoryTU]);
      var docBody = questionDocTU.getBody();
      var exactSearch = docBody.findText(categoryAnswerTU);
      var capsSearch = docBody.findText(capsAnswerTU);
      if (exactSearch !== null || capsSearch !== null){
        if (exactSearch !== null){var searchPar = exactSearch.getElement();}
        else{var searchPar = capsSearch.getElement();}
        if (searchPar.getType() === DocumentApp.ElementType.TEXT){
          searchPar = searchPar.getParent();
        }
        var tubody = searchPar.getPreviousSibling();
        var tuans = tubody.getNextSibling();
        var tuattr = tuans.getNextSibling();
        newtossup = tubody.copy();
        newtext = newtossup.insertText(0, (tossup+1) + ". ");
        newtext.setAttributes(newtossup.getAttributes());
        newend = tuattr.copy();
        newend.appendText("\n");
        newtossup_par = packetBody.appendParagraph(newtossup);
        newtossup_par.setFontSize(10);
        newtossup_par.setFontFamily("Times New Roman");
        packetBody.appendParagraph(tuans.copy());
        packetBody.appendParagraph(newend);
        tu_sheet.getRange(indexTU+2,round+1,1,1).setBackground("008888");
      }
      else{
        var docPars = docBody.getParagraphs();
        for (var i = 0; i < docPars.length; i++) {
          var paragraph = docPars[i];

          // Skip any non-text paragraphs (e.g. headers)
          if (paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) {
            continue;
          }

          if (paragraph.getText().toLowerCase().indexOf("answer:") !== -1){
          var paranswer = extractPrimaryAnswer(paragraph).toLowerCase();
            if (paranswer === categoryAnswerTU.substr(8).toLowerCase() || paranswer === categoryAnswerTU.substr(8).toLowerCase() + "s"){
              newtossup = docPars[i-1].copy();
              newtext = newtossup.insertText(0, (tossup+1) + ". ");
              newtext.setAttributes(newtossup.getAttributes());
              newend = docPars[i+1].copy();
              newend.appendText("\n");
              newtossup_par = packetBody.appendParagraph(newtossup);
              newtossup_par.setFontSize(10);
              newtossup_par.setFontFamily("Times New Roman");
              packetBody.appendParagraph(docPars[i].copy());
              packetBody.appendParagraph(newend);
              tu_sheet.getRange(indexTU+2,round+1,1,1).setBackground("008888");
              break;
            }
          }
          if (i === docPars.length-1){
            packetBody.appendParagraph("TOSSUP " + (tossup+1) + " NOT FOUND: " + data[indexTU+2][round].trim() + " (" + categoryNameTU + ")\n");
            allfound = false;
          }
        }
      }
      var categoryNameBO = bodata[indexBO+2][0];
      var metaCategoryBO = categoryNameBO.split("-")[0].trim();
      var baseAnswerBO = bodata[indexBO+2][round].split("/")[0].split("!")[0].split("{")[0].trim();
      var capsAnswerBO = "";
      var splitanswer = baseAnswerBO.split(" ");
      for (wordi = 0; wordi < splitanswer.length; wordi++){
        capsAnswerBO += splitanswer[wordi].substr(0,1).toUpperCase() + splitanswer[wordi].substr(1) + " ";
      }
      capsAnswerBO = capsAnswerBO.trim();
      var categoryAnswerBO = "ANSWER: " + baseAnswerBO;
      var capsAnswerBO = "ANSWER: " + capsAnswerBO;
      var questionDocBO = DocumentApp.openByUrl(category_id_dict[metaCategoryBO]);
      var docBody = questionDocBO.getBody();
      var exactSearch = docBody.findText(categoryAnswerBO);
      var capsSearch = docBody.findText(capsAnswerBO);
      if (exactSearch !== null || capsSearch !== null){
        if (exactSearch !== null){var searchPar = exactSearch.getElement();}
        else{var searchPar = capsSearch.getElement();}
        if (searchPar.getType() === DocumentApp.ElementType.TEXT){
          searchPar = searchPar.getParent();
        }
        var par = searchPar.getPreviousSibling().getPreviousSibling();
        for (var pi = 0; pi < 7; pi++){
          packetBody.appendParagraph(par.copy());
          par = par.getNextSibling();
        }
        newend = par.copy();
        newend.appendText("\n");
        packetBody.appendParagraph(newend);
        bo_sheet.getRange(indexBO+2,round+1,1,1).setBackground("008888");
      }
      else{
        var docPars = docBody.getParagraphs();
        for (var i = 0; i < docPars.length; i++) {
        var paragraph = docPars[i];

        // Skip any non-text paragraphs (e.g. headers)
        if (paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) {
          continue;
        }
        if (paragraph.getText().toLowerCase().indexOf("answer:") !== -1){
          var paranswer = extractPrimaryAnswer(paragraph).toLowerCase();
          if (paranswer === categoryAnswerBO.substr(8).toLowerCase() || paranswer === categoryAnswerBO.substr(8).toLowerCase() + "s"){
            for (var bi = -2; bi < 5; bi++){
              packetBody.appendParagraph(docPars[i+bi].copy());}
            newend = docPars[i+5].copy();
            newend.appendText("\n")
            packetBody.appendParagraph(newend);
            bo_sheet.getRange(indexBO+2,round+1,1,1).setBackground("008888");
            break;
          }
        }
        if (i === docPars.length-1){
          packetBody.appendParagraph("BONUS " + (tossup+1) + " NOT FOUND: " + bodata[indexBO+2][round].trim() + " (" +           categoryNameBO + ")\n");
          allfound = false;}
        }
      }
    }
    if (!allfound){
      packetDoc.setName(packetDoc.getName() + " INCOMPLETE");
    }
  }
}

function shuffle(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}

function generateValidTossups(catList){
  tuattempt = [...catList];
  while (!validTossups(tuattempt)){
    tuattempt = generateTossups([...catList]);
  }
  return tuattempt;
}

function generateTossups(catList){
  newTUs = [];
  shuffle(catList);
  newTUs.push(catList.pop());
  while (catList.length > 0){
    shuffle(catList);
    if (newTUs[newTUs.length-1].split("-")[0] !== catList[catList.length-1].split("-")[0]){
      newTUs.push(catList.pop());
    }
    else if (newTUs[0].split("-")[0] !== catList[catList.length-1].split("-")[0]){
      newTUs.splice(0,0,catList.pop());
    }
    else{
      for (var inneri = 1; inneri < newTUs.length-1; inneri++){
        if (newTUs[inneri].split("-")[0] !== catList[catList.length-1].split("-")[0] && newTUs[inneri+1].split("-")[0] !== catList[catList.length-1].split("-")[0]){
          newTUs.splice(inneri, 0, catList.pop());
          break;
        }
      }
    }
  }
  return newTUs;
}

function validTossups(catList){
  metaCats = [];
  for (i = 0; i < catList.length; i++){
    metaCats.push(catList[i].split("-")[0]);
  }
  halftime = catList.length/2;
  fhLit = 0;
  fhSci = 0;
  fhHist = 0;
  for (hi = 0; hi < halftime; hi++){
    if (metaCats[hi].indexOf("Science") !== -1){
      fhSci += 1;
    }
    else if (metaCats[hi].indexOf("History") !== -1){
      fhHist += 1;
    }
    else if (metaCats[hi].indexOf("Literature") !== -1){
      fhLit += 1;
    }
  }
  if (fhLit === 2 && fhSci === 2 && fhHist === 2){
    return true;
  }
  return false;
}

function generateValidBonuses(tossups){
  var bonuses = [...tossups];
  while (!validTossups(bonuses) || !validBonuses(tossups, bonuses))
  {
    bonuses = generateTossups([...bonuses]);
  }
  return bonuses;
}

function validBonuses(tossups, bonuses){
  for (var i = 0; i < bonuses.length; i++){
    curMeta = bonuses[i].split("-")[0];
    curTUMeta = tossups[i].split("-")[0];
    if (curMeta === curTUMeta){
      return false;
    }
  }
  return true;
}

function extractPrimaryAnswer(paragraph){
  var res = "";
  var partext = paragraph.editAsText();
  var alt_start = partext.getText().indexOf("[");
  if (alt_start < 0){
    alt_start = partext.getText().length;
  }
  for (var j = 8; j < alt_start; j++) {
    if (partext.isBold(j)) {
      res += partext.getText()[j];
    }
  }
  return res;
}
