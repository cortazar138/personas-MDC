function getPosition(string, subString, index) {
  return string.split(subString, index).join(subString).length;
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function priceAB(ab = "Carry out new pricing") {
  session.findById("wnd[2]/shell/gridRESULT/tbar/txtFINDEXPR").text = ab;
  session.findById("wnd[2]/shell/gridRESULT/tbar/btnFIND").press();
  session.findById("wnd[2]/shell/gridRESULT/tbar/btnCOPY").press();
	if (session.findById("wnd[0]/sbar").messageType == "W") {
    session.findById("wnd[0]").sendVKey(0);
    var tab = findTab("Conditions");
  }
}



//function pressKey(wnd = "0", key=0) {
//	session.findById("wnd[" + wnd + "]").sendVKey(key)
//}

function MDC() {
  
  var prom2 = session.findById("wnd[0]/usr/textEditPersonas_153967360561420").text;
  var prom2 = prom2.trim();
  // var fixed = prompt("fixed");
  var space = prom2.indexOf(" ");
  var space2 = getPosition(prom2, " ", 2);
  var space3 = getPosition(prom2, " ", 3);
  var space4 = getPosition(prom2, " ", 4);
  var space5 = getPosition(prom2, " ", 5);
  var space6 = getPosition(prom2, " ", 6);
  var space7 = getPosition(prom2, " ", 7);
  var space8 = getPosition(prom2, " ", 8);
  OA = prom2.substring(0, space + 1);
  OAline = prom2.substring(space + 1, space2);
  price = prom2.substring(space2 + 1, space3);
  var curr = prom2.substring(space3 + 1, space4);
  var per = prom2.substring(space4 + 1, space5);
  var bom = prom2.substring(space5 + 1, space6);
  var dat = prom2.substring(space6 + 1, space7);
  if (prom2.substring(space7 + 1, space8) == "test") {
    session.utils.put("test", "true");
  } else {
    session.utils.put("test", "false");
  }

  if (curr == "KRW") {
    price = price.substring(0, price.length - 3);
  }
  // IDR exception (no decimal places permited)
  debugger;
  if (curr == "IDR") {
    price = price.replace(",", "");
	price = price.replace(".", "");
    price = (parseInt(price) * 10).toString();
    per = (parseInt(per) * 1000).toString();
  }
  session.startTransaction("ME32K");
  session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").text = OA;
  // check if somebody is processing the OA
  var count = 0;
  do {
    session.findById("wnd[0]").sendVKey(0);
    var txt = session.findById("wnd[0]/sbar").text;
    count += 1;
    if (count > 500) {
      box = txt;
      //SendEmail(txt, OA, "Procurement.Services.EMEA@kemira.com");
      wasError = true;
      return;
    }
  } while (txt.search("already processing") > 0);

  if (session.findById("wnd[0]/sbar").messageType == "E") {
    box = OA + " " + OAline + "error: " + session.findById("wnd[0]/sbar").text;
    wasError = true;
    return;
  }
	
  //check if consigment
  if (session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").getCellValue(0, "RM06E-EPSTP") == "K"){
	throw "Consignment OA found, please update price manually in info record";
  }
	
  //select line
  session.findById("wnd[0]/usr/txtRM06E-EBELP").text = OAline;
  session.findById("wnd[0]").sendVKey(0);
  let row = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").firstVisibleRow.toString();
  session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").selectedRowsAbsolute = row;

  //check if correct line is selected
  let lineCheck = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0,0]").text;
 
  if (lineCheck != OAline) {
    box = OA + " " + OAline + " line error";
    wasError = true;
    return;
  }

  //check if line is blocked
  var chk = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-LOEKZ[13,0]").tooltip;
  if (chk.length > 0) {
    box = OA + " " + OAline + " line blocked";
    wasError = true;
    return;
  }
  // check if current price is 0

  if (session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]").text == "          0,00") {
    box = OA + " " + OAline + " no condition found";
    wasError = true;
    return;
  }

  var vendor = session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text;
  var mat = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").text;
  var plant = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-WERKS[11,0]").text;

  //go to price
  session.findById("wnd[0]").sendVKey(18);
	
  // target value to large
  if (session.findById("wnd[0]/sbar").text == "E: The target value is too large (check your entry)") {
    session.findById("wnd[0]/usr/txtEKPO-KTMNG").text = "99999"
    session.findById("wnd[0]").sendVKey(0);
  }
  // item target smaller than target
  if(session.idExists("wnd[1]/usr/txtMESSTXT1")) {
    session.findById("wnd[1]").sendVKey(0);
  }


  var i = 0;
  do {
    var pos = "position=" + i.toString();
    session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102").executeWebRequest("post", "action", "61", pos, null);
    session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATAB[0,0]").setFocus();
    i = i + 1;
  } while (session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATAB[0,1]").text != "__________");
  session.findById("wnd[1]").sendVKey(8);
  //find PB00 (undeleted)
  i = 0;
  do {
    var CnTy = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0," + i.toString() + "]").text;
    var del = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/chkKONP-LOEVM_KO[6," + i.toString() + "]").selected;
    i = i + 1;

    //var cond = (CnTy !="PB00" && del)
  } while (CnTy != "PB00" || del);

  i = i - 1;
  //scala?
  var isScale = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/chkRV13A-KOSTKZ[7,0]").selected;
  if (isScale) {
    box = OA + " " + OAline + " scale";
    //var msg = OA + " price not updated because of scale";
    //var subject = "MDC - Scale";
    //SendEmail(msg, subject, "Procurement.Services.EMEA@kemira.com");
    wasError = true;
    return;
  }

  session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2," + i.toString() + "]").text = price;
  session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KONWA[3," + i.toString() + "]").text = curr;
  session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KPEIN[4," + i.toString() + "]").text = per;
  session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KMEIN[5," + i.toString() + "]").text = bom;
  session.findById("wnd[0]/usr/ctxtRV13A-DATAB").text = dat;

  // save
  session.findById("wnd[0]").sendVKey(11);
  
  // target value to large
  if (session.findById("wnd[0]/sbar").text == "E: The target value is too large (check your entry)") {
    session.findById("wnd[0]/usr/txtEKPO-KTMNG").text = "99999"
    session.findById("wnd[1]").sendVKey(0); 
  }
   // item target smaller than target
   if(session.idExists("wnd[1]/usr/txtMESSTXT1")) {
    session.findById("wnd[1]").sendVKey(0);
  }

  // overlaping validity periods exception
  if (session.idExists("wnd[1]/usr/tblSAPMV13ATCTRL_D0121")) {
    session.findById("wnd[1]").sendVKey(5);
  }
  if (session.findById("wnd[0]/sbar").messageType == "E") {
    throw "status changed to 10 - to be done by PS";
  }

  session.findById("wnd[1]/usr/btnSPOP-OPTION1").press();

  POPrice(plant, mat, vendor);

  //i = 3;
  //var diff = 1;
  //var dat_now = new Date();
}

function findTab(tabName) {
  let i = 0;
  do {
    i = i + 1;
    if (
      session.idExists(
        "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
          i.toString() +
          ""
      )
    ) {
      var currentTab = session.findById(
        "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
          i.toString() +
          ""
      ).text;
      session
        .findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            i.toString() +
            ""
        )
        .select();
    }
  } while (tabName != currentTab && i < 15);
  return i;
}

function POPrice(plant, mat, vendor) {
  funName = "POprice";
  session.startTransaction("me2l");
  session.findById("wnd[0]/usr/ctxtEL_LIFNR-LOW").text = vendor;
  session.findById("wnd[0]/usr/ctxtEL_EKORG-LOW").text = "";
  session.findById("wnd[0]/usr/ctxtLISTU").text = "ALV";
  session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "WE101";
  session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").text = "";
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = plant;
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = mat;
  session.findById("wnd[0]").sendVKey(8);
  if (session.findById("wnd[0]/sbar").text == "No suitable purchasing documents found") {
    return;
  }
  var max = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount;
  console.log(max)
  for (let i = 0; i < max; i++) {
    var PGr = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "EKGRP");
    console.log(PGr);
    var oRFC = session.createRFC("ZROBO_PRGROUP_EMAIL");
    oRFC.setParameter("IV_GROUP", PGr);
    oRFC.requestResults(["EV_MAIL"]);
    oRFC.send();
    var mail = oRFC.getResultObject("EV_MAIL");
    var PO = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "EBELN");
    var creator = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "ERNAM");
    var lines = [];
    lines, (i = getLines(PO, lines, i));
  
    session.callTransaction("me22n");
    session.findById("wnd[0]/tbar[1]/btn[17]").press();
    session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = PO;
    session.findById("wnd[0]").sendVKey(0);
    //colapse header and item overwiew
    session.findById("wnd[0]").sendVKey(29);
    session.findById("wnd[0]").sendVKey(30);
    // expand item
    session.findById("wnd[0]").sendVKey(28);
    // change mode
    if (!session.idExists("wnd[0]/tbar[1]/btn[39]")) {
      session.findById("wnd[0]").sendVKey(7);
    }
    // check if no critical error occured

    if (session.findById("wnd[0]/sbar").messageType == "E") {
      box = "error: " + PO + " price not updated" + " " + session.findById("wnd[0]/sbar").text;
      //var msg = "error: " + session.findById("wnd[0]/sbar").text;
      //SendEmail(msg, subject, "procurement.services.EMEA@kemira.com");
      session.findById("wnd[0]").sendVKey(3);
      wasError = true;
      continue;
    }

    var msgToProc = "";
    for (let line of lines) {
      var tab = findTab("Conditions");
      selectLine(line);
      var oldPrice = session.findById(
        "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
          tab.toString() +
          "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/txtKOMP-NETWR"
      ).text;
      session
        .findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            tab.toString() +
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,0]"
        )
        .setFocus();
      session
        .findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            tab.toString() +
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY"
        )
        .press();
      console.log(session.findById("wnd[0]/sbar").text);
      debugger;
      if (session.findById("wnd[0]/sbar").text == "New pricing is no longer possible") {
        continue;
      }
        priceAB("Copy manual pricing")
        var newPrice = session.findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            tab.toString() +
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/txtKOMP-NETWR"
        ).text;
        session
        .findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            tab.toString() +
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,0]"
        )
        .setFocus();
      session
        .findById(
          "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
            tab.toString() +
            "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY"
        )
        .press();
        if (newPrice == oldPrice) {
          priceAB("Carry out new pricing")
        }
      
    //  session.findById("wnd[2]/shell/gridRESULT").selectedRowsAbsolute = "2"
	//  session.findById("wnd[1]").sendVKey(0);
      // clicke over yellow errors if occred
      while (session.findById("wnd[0]/sbar").messageType == "W") {
        session.findById("wnd[0]").sendVKey(0);
        var tab = findTab("Conditions");
      }
      var newPrice = session.findById(
        "wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT" +
          tab.toString() +
          "/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/txtKOMP-NETWR"
      ).text;
      if (newPrice != oldPrice && creator != "BGAUTOPO") {
        msgToProc = msgToProc + PO + "/" + line + " created by " + creator + "\n"; //
      }
      // && creator != "BGAUTOPO"
    }

   

    var tab = findTab("Conditions");

    session.findById("wnd[0]").sendVKey(11);
    session.findById("wnd[1]").sendVKey(0);

    // confirming sending new outputs for NA
    if (session.idExists("wnd[1]/usr/btnSPOP-VAROPTION1")) {
      session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press();
    }

    if (session.utils.get("test") == "true") {
      mail = "procurement.services.EMEA@kemira.com";
    }

    //if the below window is still avilable it means the transaction was not saved
    if (session.idExists("wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN")) {
      box = OA + "" + OAline + " PO: " + PO + " could not be saved : " + session.findById("wnd[0]/sbar").text;
      session.findById("wnd[0]").sendVKey(3);
      wasError = false;
      continue;
    }
    if (msgToProc.length > 2) {
      POtab.push([msgToProc, creator]);
      userTab.push(creator);
    }
  }
  //filter of userTab, because nested ifs are not possible in Personas
  var mailSub = "OA " + OA + "/" + OAline + " has been updated";
  userTab = userTab.filter(onlyUnique);
  
  var msgToUser = "";
  for (let crt of userTab) {
    var msgToUser = "";
    for (let ele of POtab) {
      if (crt == ele[1]) {
        msgToUser = msgToUser + ele[0];
      }
    }
    msgToUser = "Hello, \n please be informed that prices in below POs have been updated: \n" + msgToUser + "\n BR, \n PS EMEA team";
    SendEmail(msgToUser, mailSub, crt);
  }
}

function getLines(PO, lines, i) {
  var max = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount;
  var k = 0;
  for (let j = i; j < max; j++) {
    if (PO == session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(j, "EBELN")) {
      lines[k] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(j, "EBELP");
      k++;
      i++;
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = i - 1;
    }
  }
  i = i - 1;
  return lines, i;
}

function selectLine(line) {
  funName = "selectLine";

  var i = 0;
  var item = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST")
    .text;
  var pos = item.indexOf("]") - 1;
  var a = item.substring(2, pos);
  while (item.substring(2, pos) != line) {
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press();
    item = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:*/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").text;
    pos = item.indexOf("]") - 1;
    i = i + 1;
    if (i > 100) {
      box = "could not find PO line";
      return;
    }
  }
}

function SendEmail(msg, subject, user1) {
  funName = "SendEmail";
  session.callTransaction("SBWP");
  session.findById("wnd[0]/tbar[1]/btn[17]").press();
  session.findById("wnd[1]/usr/txtSOS06-S_FOLDES").text = "Inbox";
  session.findById("wnd[1]/usr/radSOS06-S_FOLPRV").select();
  session.findById("wnd[0]").sendVKey(0);
  session.findById("wnd[0]/usr/cntlSINWP_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell/tbar/dropdownCREA_RAW").press();
  session.findById("wnd[0]/usr/txtSOS33-OBJDES").text = subject;
  session.findById("wnd[0]/usr/tabsSO33_TAB1/tabpTAB1/ssubSUB1:SAPLSO33:2100/cntlEDITOR/shellcont/shell").text = msg;
  session.findById("wnd[0]/tbar[1]/btn[20]").press();
  session.findById(
    "wnd[1]/usr/subSENDSCREEN:SAPLSO04:1030/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,0]"
  ).text = user1;
  session
    .findById("wnd[1]/usr/subSENDSCREEN:SAPLSO04:1030/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/chkSOS04-L_SEX[2,0]")
    .select();
  session.findById(
    "wnd[1]/usr/subSENDSCREEN:SAPLSO04:1030/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/ctxtSOS04-L_ADR_NAME[0,1]"
  ).text = "Procurement.Services.EMEA@kemira.com";
  session
    .findById("wnd[1]/usr/subSENDSCREEN:SAPLSO04:1030/subRECLIST:SAPLSO04:0103/tabsTAB_CONTROL/tabpREC/ssubSUB1:SAPLSO04:0150/tblSAPLSO04REC_CONTROL/chkSOS04-L_SKP[3,1]")
    .select();
  session.findById("wnd[1]").sendVKey(0);
  if (session.idExists("wnd[2]")) {
    session.findById("wnd[2]").close();
    session.findById("wnd[1]").close();
    return;
  }
  session.findById("wnd[1]/tbar[0]/btn[20]").press();
  session.findById("wnd[0]").sendVKey(3);
}

try {
  debugger;
  // declaring global variables
  var POtab = [];
  var userTab = [];
  var OA = "";
  var OAline = "";
  var mail = "";
  var wasError = false;
  var funName = "";
  var price;
  // change of width and height so the robot can extract all log data
  session.findById("wnd[0]/usr/textEditPersonas_153967360561420").height = 30;
  session.findById("wnd[0]/usr/textEditPersonas_153967360561420").width = 200;
  var box = session.findById("wnd[0]/usr/textEditPersonas_153967360561420").text;
  MDC();
  debugger;
  session.findById("wnd[0]/tbar[0]/okcd").text = "/n";
  session.findById("wnd[0]").sendVKey(0);
  var endMessage = "";
  wasError ? (endMessage = "Line " + OA + " " + OAline + " not fully updated") : (endMessage = "no errors found during script execution!");
  session.findById("wnd[0]/usr/textEditPersonas_153967360561420").text = box + "\n" + endMessage;
} catch (err) {

  var SAPerr = session.findById("wnd[0]/sbar").text;
  session.findById("wnd[0]/tbar[0]/okcd").text = "/n";
  session.findById("wnd[0]").sendVKey(0);
  session.findById("wnd[0]/usr/textEditPersonas_153967360561420").text = box + "\n" + SAPerr + " / " + err;
  //SendEmail(msg, "unexpected error in MDC Script execution", "B9NT");
}
