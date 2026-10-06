const SHEET_ID = '1Cdie5YhYqGdPUn3J18SpwyQIqsc_5eXYJRB1wmhkcXs'; 

// 🚨 นำ Channel Access Token ของ LINE Bot ที่คุณสร้างมาใส่ตรงนี้ (ในเครื่องหมายคำพูด)
const LINE_ACCESS_TOKEN = '4EQHTk8Hvm5Zrw2dSFa5dbk0KcpQEXjkgOwbhcJy+XGDYSAJt0qJPgnSNoqBy/Gsvo/pL/XrGX1R24LznNxRNu14k1xZP6RhYL5G8CIRG3j7VVgbb9doJG0og8IE9gzX7LWexcjdrPw/rqZJ59icnAdB04t89/1O/w1cDnyilFU=';

function doGet(e) { 
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Project Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no'); 
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function getScriptUrl() { return ScriptApp.getService().getUrl(); }

function logToSheet(who, colCorrect, fromVal, toVal) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID); let logSheet = ss.getSheetByName('LOG');
    if(!logSheet) { logSheet = ss.insertSheet('LOG'); logSheet.appendRow(['date_time', 'who', 'col_correct', 'from', 'to']); }
    let ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    let safeFrom = (fromVal === "" || fromVal == null) ? "-" : String(fromVal); let safeTo = (toVal === "" || toVal == null) ? "-" : String(toVal);
    if (safeFrom !== safeTo) { logSheet.appendRow([ts, who, colCorrect, safeFrom, safeTo]); }
  } catch(e) { console.error("Log Error: " + e.message); }
}

// 🚨 ฟังก์ชันใหม่: ตัวส่งข้อความเข้า LINE Group
// 🚨 ปิดระบบบันทึก LOG ตอนส่ง LINE (ยกเว้นตอน Error ค่อยบันทึก)
function sendLinePush(groupIds, message) {
  if (!LINE_ACCESS_TOKEN || LINE_ACCESS_TOKEN === 'ใส่_TOKEN_ยาวๆ_ตรงนี้') return; 
  if (!groupIds || groupIds.length === 0) return;
  
  let uniqueIds = [...new Set(groupIds.filter(id => id && id.trim() !== ""))];
  if (uniqueIds.length === 0) return;
  
  uniqueIds.forEach(id => {
    const payload = {
      "to": id,
      "messages": [{"type": "text", "text": message}]
    };
    const options = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_ACCESS_TOKEN
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true 
    };
    try { 
        let response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options); 
        let respCode = response.getResponseCode();
        
        // บันทึกเฉพาะตอนที่ส่ง Error จริงๆ (เพื่อให้รู้ถ้าระบบพัง)
        if (respCode !== 200) {
           logToSheet("System", "LINE Error", `ส่งไม่ผ่าน [${id}]`, response.getContentText());
        }
    } catch(e) { 
        logToSheet("System", "LINE Crash", `ระบบส่งข้อมูลพัง`, e.message);
    }
  });
}

// 🚨 ฟังก์ชันใหม่: ดึงข้อมูลชื่อพนักงานคู่กับ LINE Group ID จากคอลัมน์ C และ E
function getLineIdMap() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const empSheet = ss.getSheetByName('Employees');
  const data = empSheet.getDataRange().getValues();
  let map = {};
  // เริ่มอ่านจากแถวที่ 2
  for (let i = 1; i < data.length; i++) {
    let name = data[i][2] ? data[i][2].toString().trim() : ""; // คอลัมน์ C (Index 2)
    let lineId = data[i][4] ? data[i][4].toString().trim() : ""; // คอลัมน์ E (Index 4)
    if (name && lineId) {
       map[name] = lineId;
    }
  }
  return map;
}

function getSettingsDropdown() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const settingsSheet = ss.getSheetByName('Settings'); const empSheet = ss.getSheetByName('Employees');
  let currentUserEmail = Session.getActiveUser().getEmail(); if(!currentUserEmail) currentUserEmail = "";
  
  let result = { scriptUrl: ScriptApp.getService().getUrl(), dropdowns: { brand: [], category: [], priority: [], taskStatus: [], subCategory: [], subTaskStatus: [], approvers: [] }, employeesGrouped: {}, auth: { found: false, email: currentUserEmail, name: "", overview: "NO", personal: "PERSONAL", add: "NO" } };
  
  if(settingsSheet) {
    const set_data = settingsSheet.getDataRange().getValues();
    for(let i = 1; i < set_data.length; i++) {
      if(set_data[i][1] && !result.dropdowns.brand.includes(set_data[i][1].toString())) result.dropdowns.brand.push(set_data[i][1].toString()); 
      if(set_data[i][2] && !result.dropdowns.category.includes(set_data[i][2].toString())) result.dropdowns.category.push(set_data[i][2].toString()); 
      if(set_data[i][3] && !result.dropdowns.priority.includes(set_data[i][3].toString())) result.dropdowns.priority.push(set_data[i][3].toString()); 
      if(set_data[i][4] && !result.dropdowns.taskStatus.includes(set_data[i][4].toString())) result.dropdowns.taskStatus.push(set_data[i][4].toString()); 
      if(set_data[i][6] && !result.dropdowns.subCategory.includes(set_data[i][6].toString())) result.dropdowns.subCategory.push(set_data[i][6].toString()); 
      if(set_data[i][7] && !result.dropdowns.subTaskStatus.includes(set_data[i][7].toString())) result.dropdowns.subTaskStatus.push(set_data[i][7].toString()); 
      if(set_data[i][8] && !result.dropdowns.approvers.includes(set_data[i][8].toString())) result.dropdowns.approvers.push(set_data[i][8].toString()); 
    }
  }
  if(empSheet) {
    const emp_data = empSheet.getDataRange().getValues();
    for(let i = 1; i < emp_data.length; i++) {
      let email = emp_data[i][0] ? emp_data[i][0].toString().trim().toLowerCase() : "";        
      let dept  = emp_data[i][1] ? emp_data[i][1].toString().trim() : "อื่นๆ";       
      let name  = emp_data[i][2] ? emp_data[i][2].toString().trim() : "";        
      let auth_Overview = emp_data[i][6] ? emp_data[i][6].toString().toUpperCase().trim() : "NO";    
      let auth_Personal = emp_data[i][7] ? emp_data[i][7].toString().toUpperCase().trim() : "PERSONAL"; 
      let auth_Add      = emp_data[i][8] ? emp_data[i][8].toString().toUpperCase().trim() : "NO";    
      if(name !== "") { if(!result.employeesGrouped[dept]) result.employeesGrouped[dept] = []; if(!result.employeesGrouped[dept].includes(name)) result.employeesGrouped[dept].push(name); }
      if(email !== "" && currentUserEmail !== "" && email === currentUserEmail.toLowerCase()) { result.auth.found = true; result.auth.name = name; result.auth.overview = auth_Overview; result.auth.personal = auth_Personal; result.auth.add = auth_Add; }
    }
  }
  return result;
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_ID); const mainSheet = ss.getSheetByName('Main_Tasks'); const subSheet = ss.getSheetByName('Sub_Tasks');
  if (!mainSheet || !subSheet) throw new Error("ไม่พบชีต Main_Tasks หรือ Sub_Tasks");
  
  const mainData = mainSheet.getDataRange().getValues(); const subData = subSheet.getDataRange().getValues(); let projects = [];
  const safeStr = (val) => { if (val === null || val === undefined || val === "") return ""; if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"); return val.toString().trim(); };
  
  for(let i = 1; i < mainData.length; i++) {
    let code = safeStr(mainData[i][1]); if(!code) continue;
    let p = { id: safeStr(mainData[i][0]), code: code, name: safeStr(mainData[i][3]), brand: safeStr(mainData[i][2]), category: safeStr(mainData[i][4]), priority: safeStr(mainData[i][5]), creator: safeStr(mainData[i][6]), owner: safeStr(mainData[i][7]), startDate: safeStr(mainData[i][8]), dueDate: safeStr(mainData[i][9]), approvalStatus: safeStr(mainData[i][10]), approvedBy: safeStr(mainData[i][11]), status: safeStr(mainData[i][13]), resultLink: safeStr(mainData[i][14]), remark: safeStr(mainData[i][15]), subTasks: [] }; projects.push(p);
  }
  for(let i = 1; i < subData.length; i++) {
    let mainId = safeStr(subData[i][2]); if(!mainId) continue; 
    let sub = { subId: safeStr(subData[i][0]), subCode: safeStr(subData[i][1]), subCategory: safeStr(subData[i][3]), subName: safeStr(subData[i][4]), assignee: safeStr(subData[i][5]), description: safeStr(subData[i][6]), startDate: safeStr(subData[i][7]), dueDate: safeStr(subData[i][8]), status: safeStr(subData[i][9]), resultLink: safeStr(subData[i][10]), remark: safeStr(subData[i][11]), subApprover: safeStr(subData[i][14]) };
    let proj = projects.find(p => p.id === mainId); if(proj) proj.subTasks.push(sub);
  }
  return projects;
}

function saveProjectData(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID); const mainSheet = ss.getSheetByName('Main_Tasks'); const subSheet = ss.getSheetByName('Sub_Tasks'); const empSheet = ss.getSheetByName('Employees');
  let userCode = "XX"; let empData = empSheet.getDataRange().getValues();
  for(let i=1; i<empData.length; i++) { if(empData[i][2] == payload.main.creator) { userCode = empData[i][5] ? empData[i][5].toString().trim() : "XX"; break; } }
  let random5 = Math.random().toString(36).substring(2, 7).toUpperCase(); let taskId = "T" + userCode + random5; 
  let d = new Date(); let yy = d.getFullYear().toString().slice(-2); let mm = ("0" + (d.getMonth() + 1)).slice(-2); let prefix = userCode + yy + mm; 

  const mainData = mainSheet.getDataRange().getValues(); let existingNums = []; 
  for(let i=1; i<mainData.length; i++) { let codeStr = mainData[i][1] ? mainData[i][1].toString() : ""; if(codeStr.startsWith(prefix)) { let numStr = codeStr.substring(prefix.length, prefix.length + 3); let num = parseInt(numStr); if(!isNaN(num)) { existingNums.push(num); } } }
  let maxSeq = 0; if(existingNums.length > 0) { maxSeq = Math.max(...existingNums); }
  let nextSeq = ("000" + (maxSeq + 1)).slice(-3); let mainCode = prefix + nextSeq; 
  let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  mainSheet.appendRow([taskId, mainCode, payload.main.brand, payload.main.taskName, payload.main.category, payload.main.priority, payload.main.creator, payload.main.assignee, payload.main.startDate, payload.main.dueDate, payload.main.approvalStatus, payload.main.approvedBy, "", payload.main.taskStatus, payload.main.resultLink||"", payload.main.remark, timestamp]);
  logToSheet(payload.who, `สร้างโปรเจกต์ใหม่`, "-", mainCode);
  
  let lineMap = getLineIdMap();
  let creatorLineId = lineMap[payload.main.creator];

  // 1. ส่งแจ้งเตือนเมื่อเปิดโปรเจค
  if(creatorLineId) {
      let msg1 = `【 📢 แจ้งเปิดโปรเจคใหม่ 🆕 】\nชื่อโปรเจค : ${payload.main.taskName}\nความสำคัญ : ${payload.main.priority}\nวันที่เริ่มต้น : ${payload.main.startDate}\nวันที่กำหนดเสร็จ : ${payload.main.dueDate}`;
      sendLinePush([creatorLineId], msg1);
  }
  
  if(payload.subs && payload.subs.length > 0) {
    let subRows = []; let timeSeed = new Date().getTime();
    payload.subs.forEach((sub, index) => { 
      let random3 = Math.random().toString(36).substring(2, 5).toUpperCase(); let subId = "ST-" + timeSeed + "-" + index + random3; let subCode = mainCode + "-" + (index + 1); let approverVal = sub.subApprover ? sub.subApprover : "-"; 
      subRows.push([subId, subCode, taskId, sub.subCategory, sub.subName, sub.subAssignee, sub.description, sub.startDate, sub.dueDate, sub.status, sub.resultLink, sub.remark, "", "", approverVal]); 

      // 2. แจ้งเตือนคนถูกสั่งงานย่อยใหม่
      let assignees = (sub.subAssignee || "").split(',').map(s=>s.trim());
      let assignLineIds = assignees.map(a => lineMap[a]).filter(id => id);
      if(assignLineIds.length > 0) {
          let msg2 = `【 🔔 แจ้งเตือนงานใหม่ 📝 】\nชื่อโปรเจค : ${payload.main.taskName}\nชื่องานย่อย : ${sub.subName}\nความสำคัญ : ${payload.main.priority}\nวันที่เริ่มต้น : ${sub.startDate}\nวันที่กำหนดเสร็จ : ${sub.dueDate}\nรายละเอียดที่ต้องทำ : ${sub.description}\n****** โปรดเข้าไปดูงานใน Web เพื่อดูรายละเอียด *******`;
          sendLinePush(assignLineIds, msg2);
      }
    });
    if (subRows.length > 0) subSheet.getRange(subSheet.getLastRow() + 1, 1, subRows.length, subRows[0].length).setValues(subRows);
  }
  return "บันทึกข้อมูลโปรเจกต์ " + mainCode + " สำเร็จ!";
}

function updateSubTask(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID); const subSheet = ss.getSheetByName('Sub_Tasks'); const data = subSheet.getDataRange().getValues(); const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  for(let i = 1; i < data.length; i++) { 
    if(data[i][0] == payload.subId) { 
        let oldStatus = data[i][9]; let oldLink = data[i][10]; let oldRemark = data[i][11];
        let mainId = data[i][2]; let assigneeStr = data[i][5]; let subDesc = data[i][6]; let subApprover = data[i][14];
        
        subSheet.getRange(i+1, 10).setValue(payload.status); subSheet.getRange(i+1, 11).setValue(payload.resultLink); subSheet.getRange(i+1, 12).setValue(payload.remark); 
        
        if(oldStatus !== payload.status) { 
            logToSheet(payload.who, `สถานะงานย่อย [${data[i][1]}]`, oldStatus, payload.status); 
            if(payload.status.includes('เสร็จสิ้น')) { subSheet.getRange(i+1, 14).setValue(ts); } else { subSheet.getRange(i+1, 14).setValue(""); } 
            
            const mainSheet = ss.getSheetByName('Main_Tasks');
            const mainData = mainSheet.getDataRange().getValues();
            let projName = "ไม่ระบุ";
            for(let j=1; j<mainData.length; j++){ if(mainData[j][0] == mainId) { projName = mainData[j][3]; break; } }
            
            let lineMap = getLineIdMap();
            
            if(payload.status.includes('เสร็จสิ้น') && !payload.status.includes('ตรวจสอบแล้ว')) {
                // 3. แจ้งตรวจงาน
                let approverLineId = lineMap[subApprover];
                if(approverLineId) {
                    let msg3 = `【 🔎 แจ้งตรวจงาน 🧐 】\nชื่อโปรเจค : ${projName}\nชื่องานย่อย : ${payload.subName}\nรายละเอียดที่ต้องทำ : ${subDesc}\nลิงค์ผลงาน : ${payload.resultLink || '-'}\nหมายเหตุ : ${payload.remark || '-'}\n****** โปรดเข้าไปดูงานใน Web เพื่อดูรายละเอียด *******`;
                    sendLinePush([approverLineId], msg3);
                }
            } 
            else if (payload.status === 'ปฎิเสธ / แก้ไข' || payload.status === 'ปฏิเสธ / แก้ไข') {
                // 4. แจ้งแก้ไขงาน 
                let assignees = assigneeStr.split(',').map(s=>s.trim());
                let assignLineIds = assignees.map(a => lineMap[a]).filter(id => id);
                if(assignLineIds.length > 0) {
                    let msg4 = `【 ❌ แจ้งแก้ไขงาน ⚠️ 】\nชื่อโปรเจค : ${projName}\nชื่องานย่อย : ${payload.subName}\nหมายเหตุแก้ไข : ${payload.remark || '-'}\n****** โปรดเข้าไปดูงานใน Web เพื่อดูรายละเอียด *******`;
                    sendLinePush(assignLineIds, msg4);
                }
            } 
            else if (payload.status === 'เสร็จสิ้นและตรวจสอบแล้ว') {
                // 5. แจ้งตรวจสอบงานเสร็จสิ้น 
                let assignees = assigneeStr.split(',').map(s=>s.trim());
                let assignLineIds = assignees.map(a => lineMap[a]).filter(id => id);
                if(assignLineIds.length > 0) {
                    let msg5 = `【 ✅ แจ้งตรวจสอบงานเสร็จสิ้น 🎉 】\nชื่อโปรเจค : ${projName}\nชื่องานย่อย : ${payload.subName}\nสถานะตรวจ : เสร็จสิ้นและตรวจสอบแล้ว\n****** โปรดแจ้งงานต่อไปที่คนที่ต้องทำงานในขั้นตอนต่อไป ******`;
                    sendLinePush(assignLineIds, msg5);
                }
            }
        }
        
        if(oldLink !== payload.resultLink) logToSheet(payload.who, `ลิงก์ผลงานย่อย [${data[i][1]}]`, oldLink, payload.resultLink);
        if(oldRemark !== payload.remark) logToSheet(payload.who, `หมายเหตุงานย่อย [${data[i][1]}]`, oldRemark, payload.remark);
        return "อัปเดตงาน " + payload.subName + " เรียบร้อย!"; 
    } 
  }
  throw new Error("ไม่พบรหัสงานย่อย " + payload.subId);
}

function updateFullProject(payload) {
  const ss = SpreadsheetApp.openById(SHEET_ID); const mainSheet = ss.getSheetByName('Main_Tasks'); const subSheet = ss.getSheetByName('Sub_Tasks'); const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  let mainData = mainSheet.getDataRange().getValues(); let foundMain = false; let taskId = ""; 

  for(let i=1; i<mainData.length; i++) {
     if(mainData[i][1] === payload.mainCode) {
         taskId = mainData[i][0]; let m = payload.main;
         let mappings = [ {col: 3, name: "แบรนด์", newVal: m.brand}, {col: 4, name: "ชื่องานหลัก", newVal: m.name}, {col: 5, name: "หมวดหมู่", newVal: m.category}, {col: 6, name: "ความสำคัญ", newVal: m.priority}, {col: 8, name: "ผู้คุม (Owner)", newVal: m.owner}, {col: 9, name: "วันที่เริ่ม", newVal: m.startDate}, {col: 10, name: "กำหนดส่งหลัก", newVal: m.dueDate}, {col: 14, name: "สถานะหลัก", newVal: m.status}, {col: 15, name: "ลิงก์รวม", newVal: m.resultLink} ];
         mappings.forEach(map => { let oldVal = mainData[i][map.col-1]; if (String(oldVal) !== String(map.newVal)) { logToSheet(payload.who, `${map.name} [${payload.mainCode}]`, oldVal, map.newVal); mainSheet.getRange(i+1, map.col).setValue(map.newVal); } });
         mainSheet.getRange(i+1, 16).setValue(m.remark); foundMain = true; break;
     }
  }
  if(!foundMain) throw new Error("ไม่พบโปรเจกต์ " + payload.mainCode);

  let subData = subSheet.getDataRange().getValues(); let subCountForThisMain = 0;
  for(let i=1; i<subData.length; i++) { if(subData[i][2] === taskId) { subCountForThisMain++; } }
  let payloadSubIds = [];
  
  let lineMap = getLineIdMap();

  payload.subs.forEach(sub => {
     let approverVal = sub.subApprover ? sub.subApprover : "-";
     if(sub.subId && sub.subId.trim() !== "") { 
         payloadSubIds.push(sub.subId);
         for(let i=1; i<subData.length; i++) {
             if(subData[i][0] === sub.subId) {
                 let oldDesc = subData[i][6]; let oldDue = subData[i][8]; let oldStatus = subData[i][9]; let oldInlineLog = subData[i][12] || ""; let newInlineLog = oldInlineLog; let changedLog = false;
                 if (String(oldDesc) !== String(sub.description)) { logToSheet(payload.who, `รายละเอียดงานย่อย [${subData[i][1]}]`, oldDesc, sub.description); newInlineLog += `[${ts} | ${payload.who}] แก้ไขรายละเอียด\n`; changedLog = true; }
                 if (String(oldDue) !== String(sub.dueDate)) { logToSheet(payload.who, `กำหนดส่งงานย่อย [${subData[i][1]}]`, oldDue, sub.dueDate); newInlineLog += `[${ts} | ${payload.who}] เลื่อนกำหนดเป็น ${sub.dueDate}\n`; changedLog = true; }
                 if (String(oldStatus) !== String(sub.status)) { logToSheet(payload.who, `สถานะงานย่อย [${subData[i][1]}]`, oldStatus, sub.status); if (sub.status.includes('เสร็จสิ้น')) { subSheet.getRange(i+1, 14).setValue(ts); } else if (!sub.status.includes('เสร็จสิ้น') && String(oldStatus).includes('เสร็จสิ้น')) { subSheet.getRange(i+1, 14).setValue(""); } }
                 subSheet.getRange(i+1, 4).setValue(sub.category); subSheet.getRange(i+1, 5).setValue(sub.name); subSheet.getRange(i+1, 6).setValue(sub.assignee); subSheet.getRange(i+1, 7).setValue(sub.description); subSheet.getRange(i+1, 8).setValue(sub.startDate); subSheet.getRange(i+1, 9).setValue(sub.dueDate); subSheet.getRange(i+1, 10).setValue(sub.status); 
                 if(changedLog) subSheet.getRange(i+1, 13).setValue(newInlineLog); subSheet.getRange(i+1, 15).setValue(approverVal);
                 break;
             }
         }
     }
  });

  for(let i = subData.length - 1; i > 0; i--) {
     if(subData[i][2] === taskId) { let existingSubId = subData[i][0].toString(); if(existingSubId !== "" && !payloadSubIds.includes(existingSubId)) { logToSheet(payload.who, `ลบงานย่อย [${subData[i][1]}]`, subData[i][4], "(ถูกลบทิ้ง)"); subSheet.deleteRow(i + 1); } }
  }

  let newSubRows = []; let timeSeed = new Date().getTime();
  payload.subs.forEach((sub, index) => {
     let approverVal = sub.subApprover ? sub.subApprover : "-";
     if(!sub.subId || sub.subId.trim() === "") { 
         subCountForThisMain++; let random3 = Math.random().toString(36).substring(2, 5).toUpperCase(); let newSubId = "ST-" + timeSeed + "-" + index + random3; let newSubCode = payload.mainCode + "-" + subCountForThisMain; 
         logToSheet(payload.who, `เพิ่มงานย่อยใหม่ [${newSubCode}]`, "-", sub.name); 
         newSubRows.push([newSubId, newSubCode, taskId, sub.category, sub.name, sub.assignee, sub.description, sub.startDate, sub.dueDate, sub.status, "", "", "", "", approverVal]); 

         let assignees = (sub.assignee || "").split(',').map(s=>s.trim());
         let assignLineIds = assignees.map(a => lineMap[a]).filter(id => id);
         if(assignLineIds.length > 0) {
             let msg2 = `【 🔔 แจ้งเตือนงานใหม่ 📝 】\nชื่อโปรเจค : ${payload.main.name}\nชื่องานย่อย : ${sub.name}\nความสำคัญ : ${payload.main.priority}\nวันที่เริ่มต้น : ${sub.startDate}\nวันที่กำหนดเสร็จ : ${sub.dueDate}\nรายละเอียดที่ต้องทำ : ${sub.description}\n****** โปรดเข้าไปดูงานใน Web เพื่อดูรายละเอียด *******`;
             sendLinePush(assignLineIds, msg2);
         }
     }
  });
  if(newSubRows.length > 0) { subSheet.getRange(subSheet.getLastRow() + 1, 1, newSubRows.length, newSubRows[0].length).setValues(newSubRows); }
  return "แก้ไขและอัปเดตข้อมูลเสร็จสมบูรณ์!";
}
function deleteProjectData(mainCode, who) {
  const ss = SpreadsheetApp.openById(SHEET_ID); const mainSheet = ss.getSheetByName('Main_Tasks'); const subSheet = ss.getSheetByName('Sub_Tasks');
  let mainDeleted = false; let taskId = ""; const mainData = mainSheet.getDataRange().getValues();
  for(let i = mainData.length - 1; i > 0; i--) { if(mainData[i][1] === mainCode) { taskId = mainData[i][0]; logToSheet(who || "System", `ลบโปรเจกต์หลัก`, mainCode, "(ถูกลบทั้งหมด)"); mainSheet.deleteRow(i + 1); mainDeleted = true; break; } }
  if(!mainDeleted) throw new Error("ไม่พบโปรเจกต์ " + mainCode + " ในระบบ");
  const subData = subSheet.getDataRange().getValues();
  for(let i = subData.length - 1; i > 0; i--) { if(subData[i][2] === taskId) { subSheet.deleteRow(i + 1); } }
  return "ลบโปรเจกต์ [" + mainCode + "] และงานย่อยที่เกี่ยวข้องทั้งหมดเรียบร้อยแล้ว!";
}
// 🟢 ฟังก์ชันสำหรับใช้เทสระบบ LINE ว่าติดขัดตรงไหน
function testLineNotify() {
  // 🚨 1. เอา Group ID หรือ User ID จากคอลัมน์ E มาใส่ตรงนี้ 1 อันครับ
  var testGroupId = "ใส่_Group_ID_ตรงนี้"; 
  
  var payload = {
    "to": testGroupId,
    "messages": [{"type": "text", "text": "🟢 ทดสอบระบบแจ้งเตือนจาก Google Apps Script! ถ้าระบบสมบูรณ์ ข้อความนี้จะต้องเด้งเข้า LINE ครับ"}]
  };
  
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + LINE_ACCESS_TOKEN // ใช้ Token ที่คุณใส่ไว้ด้านบน
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    Logger.log("✅ ผลลัพธ์จาก LINE API: " + response.getContentText());
  } catch(e) {
    Logger.log("❌ ระบบ Google พังตอนส่ง: " + e.message);
  }
}
// =========================================================
// ⏰ ระบบแจ้งเตือนตามเวลา (Daily Reminder: เช้า, บ่าย, เย็น)
// =========================================================

function processTaskReminders(timePeriod) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const mainSheet = ss.getSheetByName('Main_Tasks');
  const subSheet = ss.getSheetByName('Sub_Tasks');
  
  if (!mainSheet || !subSheet) return;
  
  const mainData = mainSheet.getDataRange().getValues();
  const subData = subSheet.getDataRange().getValues();
  const lineMap = getLineIdMap();
  
  let userTasks = {};
  const now = new Date();
  const todayZero = new Date(); todayZero.setHours(0,0,0,0);
  const tomorrowZero = new Date(todayZero); tomorrowZero.setDate(todayZero.getDate() + 1);

  // วนลูปอ่านงานย่อยทั้งหมด
  for (let i = 1; i < subData.length; i++) {
      let mainId = subData[i][2];
      let subName = subData[i][4];
      let assigneesStr = subData[i][5] ? subData[i][5].toString() : "";
      let dueDateStr = subData[i][8] ? subData[i][8].toString() : "";
      let status = subData[i][9] ? subData[i][9].toString() : "";
      let completedDateStr = subData[i][13] ? subData[i][13].toString() : ""; // คอลัมน์ N ที่เก็บเวลาเสร็จ
      
      let isCompleted = status.includes('เสร็จสิ้น') && !status.includes('รอตรวจ'); 
      let isWaitingForApproval = status.includes('เสร็จสิ้น (รอตรวจ)');
      
      if (status.includes('Cancel') || status === 'ยังไม่ต้องดำเนินการ') continue;
      
      let projName = "ไม่ระบุ";
      let priority = "";
      for (let j = 1; j < mainData.length; j++) {
          if (mainData[j][0] == mainId) {
              projName = mainData[j][3];
              priority = mainData[j][5] ? mainData[j][5].toString() : "";
              let mainStatus = mainData[j][13] ? mainData[j][13].toString() : "";
              if (mainStatus.includes('เสร็จ') || mainStatus.toLowerCase().includes('done')) {
                  projName = "FINISHED";
              }
              break;
          }
      }
      
      if (projName === "FINISHED") continue;
      
      // คำนวณวันที่กำหนดส่ง
      let dueDateObj = new Date(9999, 0, 1);
      if (dueDateStr) {
          try {
             if (dueDateStr instanceof Date) { dueDateObj = dueDateStr; }
             else {
               let str = dueDateStr.trim();
               let parts = str.split(/[ \T]/);
               let dmy = parts[0].split(/[\/-]/);
               if(dmy.length === 3) {
                 let date = parseInt(dmy[0]); let m = parseInt(dmy[1]) - 1; let y = parseInt(dmy[2]);
                 if(y > 2500) y -= 543;
                 let time = parts[1] ? parts[1].split(':') : [0,0];
                 dueDateObj = new Date(y, m, date, parseInt(time[0])||0, parseInt(time[1])||0);
               }
             }
          } catch(e) {}
      }
      
      let dueDateZero = new Date(dueDateObj); dueDateZero.setHours(0,0,0,0);
      let isDueToday = dueDateZero.getTime() === todayZero.getTime();
      let isDueTomorrow = dueDateZero.getTime() === tomorrowZero.getTime();
      let isOverdue = (!isCompleted && !isWaitingForApproval) && dueDateObj < now;

      // เช็กว่าเพิ่งเสร็จวันนี้หรือไม่ (ดูจากคอลัมน์ N)
      let finishedToday = false;
      if ((isCompleted || isWaitingForApproval) && completedDateStr !== "") {
          let cStr = completedDateStr.split(' ')[0]; // เอาแค่ dd/MM/yyyy
          let tStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
          if (cStr === tStr) finishedToday = true;
      }

      let assignees = assigneesStr.split(',').map(s => s.trim()).filter(s => s !== "");
      
      assignees.forEach(empName => {
          if (!userTasks[empName]) {
              userTasks[empName] = { pendingTasks: [], completedToday: 0, dueTomorrowTasks: [] };
          }
          
          if (finishedToday) {
              userTasks[empName].completedToday++;
          } else if (!isCompleted && !isWaitingForApproval) {
              userTasks[empName].pendingTasks.push({
                 projName: projName,
                 subName: subName,
                 dueStr: dueDateStr,
                 isOverdue: isOverdue,
                 isDueToday: isDueToday,
                 status: status,
                 priority: priority
              });
          }

          if (!isCompleted && !isWaitingForApproval && isDueTomorrow) {
              userTasks[empName].dueTomorrowTasks.push({ subName: subName });
          }
      });
  }
  
  // สร้างและส่งข้อความตามช่วงเวลา
  for (let empName in userTasks) {
      let data = userTasks[empName];
      let lineId = lineMap[empName];
      
      if (lineId) {
          let msg = "";
          let activeTasks = data.pendingTasks;
          let totalPending = activeTasks.length;

          // จัดเรียง: ล่าช้าขึ้นก่อน ตามด้วยส่งวันนี้ ตามด้วยความสำคัญ
          activeTasks.sort((a, b) => {
              if (a.isOverdue !== b.isOverdue) return a.isOverdue ? -1 : 1;
              if (a.isDueToday !== b.isDueToday) return a.isDueToday ? -1 : 1;
              return b.priority.localeCompare(a.priority);
          });

          // 🌅 ข้อความตอนเช้า
          if (timePeriod === "MORNING") {
              if (totalPending === 0) {
                  msg = `🌅 อรุณสวัสดิ์! วันนี้คุณ ${empName}\n🎉 ไม่มีงานค้างในระบบเลยครับ ยอดเยี่ยมมาก!`;
              } else {
                  msg = `🌅 อรุณสวัสดิ์! วันนี้คุณ ${empName} มีงานที่ต้องทำตามนี้ โปรดทำงานอย่างแข็งขันด้วยล่ะ 💪\n\n`;
                  msg += `╭━━━━━━━━━━━━━╮\n`;
                  msg += `     📝 รายการงาน (${totalPending})\n`;
                  msg += `╰━━━━━━━━━━━━━╯\n`;
                  let limit = Math.min(totalPending, 7);
                  for (let i = 0; i < limit; i++) {
                      let t = activeTasks[i];
                      let alert = t.isOverdue ? "🚨[ล่าช้า]" : (t.isDueToday ? "⏰[ส่งวันนี้]" : "⏳");
                      msg += `${alert} ${t.subName}\n   └ 📌 ${t.status}\n`;
                  }
                  if (totalPending > 7) msg += `...และงานอื่นๆ อีก ${totalPending - 7} งาน\n`;
              }
          } 
          // ☀️ ข้อความตอนบ่าย
          else if (timePeriod === "AFTERNOON") {
              if (totalPending === 0) {
                  msg = `☀ ช่วงบ่ายแล้ว!\n\nไม่มีงานค้างแล้วนะ เก่งมาก 👍🎉`;
              } else {
                  msg = `☀️ ช่วงบ่ายแล้ว!\nทำงานเสร็จไปแล้ว ${data.completedToday} งาน ✅\nแต่วันนี้ยังเหลืองานที่ต้องทำส่งอยู่จำนวน ${totalPending} งาน 📝\n\n`;
                  msg += `╭━━━━━━━━━━━━━╮\n`;
                  msg += `     📋 งานที่ยังค้างอยู่\n`;
                  msg += `╰━━━━━━━━━━━━━╯\n`;
                  let limit = Math.min(totalPending, 5);
                  for (let i = 0; i < limit; i++) {
                      let t = activeTasks[i];
                      let alert = t.isOverdue ? "🚨" : (t.isDueToday ? "⏰" : "📌");
                      msg += `${alert} ${t.subName}\n`;
                  }
                  if (totalPending > 5) msg += `...และอีก ${totalPending - 5} งาน\n`;
              }
          }
          // 🌆 ข้อความตอนเย็น
          else if (timePeriod === "EVENING") {
              msg = `🌆 สรุปผลงานประจำวัน (End of Day)\n\n`;
              msg += `✅ วันนี้ทำงานเสร็จสิ้นไป: ${data.completedToday} งาน\n`;
              msg += `📝 มีงานที่ยังค้างอยู่: ${totalPending} งาน\n`;
              msg += `📅 พรุ่งนี้มีงานรอส่ง: ${data.dueTomorrowTasks.length} งาน\n\n`;
              
              if (totalPending > 0) {
                  msg += `😭 ยังมีงานค้างอยู่นะ ทำเสร็จหรือยังน้า อย่าลืมส่งด้วยล่ะ!`;
              } else {
                  msg += `🎉 ไม่มีงานค้างแล้วนะ เก่งมาก 👍 กลับบ้านพักผ่อนให้สบายใจได้เลย 🏡`;
              }
          }

          if (msg !== "") {
              sendLinePush([lineId], msg); // ใช้ฟังก์ชัน sendLinePush เดิมที่มีอยู่แล้ว
          }
      }
  }
}

// 🟢 ฟังก์ชันสำหรับให้ Trigger เรียกใช้ตามเวลา
function triggerMorningReminder() { processTaskReminders("MORNING"); }
function triggerAfternoonReminder() { processTaskReminders("AFTERNOON"); }
function triggerEveningReminder() { processTaskReminders("EVENING"); }
