/**
 * 3D Printing Cost Calculator and Inventory System (Modern ES6+ Version)
 * Version: 2.1 (Multi-color support up to 4 filaments & Web App updated)
 */

/**
 * ฟังก์ชันสำหรับ Web App
 */
const doGet = () => {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('3D Print Cost Calculator (Multi-color)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
};

/**
 * ดึงข้อมูลเครื่องพิมพ์และเส้นสำหรับส่งให้หน้าเว็บ
 */
const getDropdownData = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const inventorySheet = ss.getSheetByName('Inventory');
  
  const printers = settingsSheet.getRange(7, 1, Math.max(1, settingsSheet.getLastRow() - 6), 1).getValues().flat().filter(String);
  const filaments = inventorySheet.getRange(2, 1, Math.max(1, inventorySheet.getLastRow() - 1), 1).getValues().flat().filter(String);
  
  return { printers, filaments };
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('3D Print Manager')
      .addItem('1. สร้างโครงสร้างตาราง (Setup)', 'setupSheets')
      .addItem('2. อัปเดตรายชื่อเครื่อง/เส้น (Sync)', 'syncSettings')
      .addSeparator()
      .addItem('คำนวณต้นทุน (Calculate)', 'calculateCostFromUI')
      .addItem('ตัดสต็อกสินค้า (Complete Order)', 'completeOrder')
      .addToUi();
}

/**
 * ฟังก์ชันสำหรับสร้างหน้า Sheet และหัวข้อตารางให้อัตโนมัติ (Updated for 4 colors)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Settings Sheet
  const settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  settings.clear();
  settings.getRange('A1:B3').setValues([
    ['Electricity Rate (per kWh)', 5.0],
    ['Labor Rate (per hour)', 50.0],
    ['Failure Buffer (%)', 10.0]
  ]);
  settings.getRange('A5:C6').setValues([
    ['--- Printer Profiles ---', '', ''],
    ['Printer Name', 'Power (Watts)', 'Wear & Tear (per hr)']
  ]);
  settings.getRange('A7:C7').setValues([['A1 Combo', 120, 3.0]]);
  settings.getRange('A1:A6').setFontWeight('bold');

  // 2. Inventory Sheet
  const inv = ss.getSheetByName('Inventory') || ss.insertSheet('Inventory');
  inv.clear();
  inv.getRange('A1:I1').setValues([
    ['Spool ID', 'Brand', 'Material', 'Color', 'Total Weight (g)', 'Cost per Spool', 'Cost per Gram', 'Remaining Weight (g)', 'Status']
  ]).setFontWeight('bold');
  
  // Formulas
  inv.getRange('G2:G100').setFormula('=IF(E2>0, F2/E2, 0)');
  inv.getRange('H2:H100').setFormula('=IF(E2>0, E2, 0)'); // เริ่มต้น Remaining = Total
  inv.getRange('I2:I100').setFormula('=IF(H2<=0, "Out of Stock", IF(H2<200, "Low Stock", "Available"))');

  // 3. Calculator Sheet (Layout for 4 colors + Wipe Tower)
  const calc = ss.getSheetByName('Calculator') || ss.insertSheet('Calculator');
  calc.clear();
  calc.getRange('A1:C1').setValues([['Input Details', 'Selection', 'Weight (g)']]).setFontWeight('bold');
  calc.getRange('A2:A8').setValues([
    ['Printer Name'],
    ['Filament 1'],
    ['Filament 2'],
    ['Filament 3'],
    ['Filament 4'],
    ['Wipe Tower / Waste'],
    ['Total Print Time (Hrs)']
  ]);
  calc.getRange('A9:A11').setValues([
    ['Setup/Post Time (Min)'],
    ['Profit Margin (%)'],
    ['']
  ]);
  calc.getRange('D2:D7').setValues([
    ['Filament Cost'],
    ['Electricity Cost'],
    ['Wear & Tear Cost'],
    ['Labor Cost'],
    ['Total Cost (+Buffer)'],
    ['Suggested Price']
  ]);
  calc.getRange('A2:A10').setFontWeight('bold');
  calc.getRange('D2:D7').setFontWeight('bold');
  
  // Default values
  calc.getRange('B10').setValue(30);
  calc.getRange('B9').setValue(15);

  // 4. Orders Sheet
  const orders = ss.getSheetByName('Orders') || ss.insertSheet('Orders');
  orders.clear();
  orders.getRange('A1:I1').setValues([
    ['Order ID', 'Date', 'Project Name', 'Printer', 'Filament IDs', 'Total Weight (g)', 'Total Cost', 'Suggested Price', 'Status']
  ]).setFontWeight('bold');

  SpreadsheetApp.getUi().alert('สร้างโครงสร้างตารางใหม่ (รองรับ 4 สี) เรียบร้อยแล้ว!');
}

/**
 * ดึงข้อมูลจาก UI มาคำนวณ (รองรับ 4 สี)
 */
function calculateCostFromUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName('Calculator');
  
  const printerName = calcSheet.getRange('B2').getValue();
  const printTimeHours = calcSheet.getRange('B8').getValue();
  const setupTimeMinutes = calcSheet.getRange('B9').getValue();
  const marginPercent = calcSheet.getRange('B10').getValue();
  const wasteWeight = calcSheet.getRange('C7').getValue() || 0;
  
  const filaments = [];
  // Loop through 4 color slots (B3:C6)
  for (let i = 3; i <= 6; i++) {
    const id = calcSheet.getRange(i, 2).getValue();
    const weight = calcSheet.getRange(i, 3).getValue() || 0;
    if (id) {
      filaments.push({ id, weight });
    }
  }
  
  if (!printerName || filaments.length === 0) {
    SpreadsheetApp.getUi().alert('กรุณาเลือกเครื่องพิมพ์และเส้นอย่างน้อย 1 ชนิด');
    return;
  }
  
  const results = calculateCost(printerName, filaments, wasteWeight, printTimeHours, setupTimeMinutes, marginPercent);
  
  calcSheet.getRange('E2').setValue(results.filamentCost);
  calcSheet.getRange('E3').setValue(results.electricityCost);
  calcSheet.getRange('E4').setValue(results.wearAndTearCost);
  calcSheet.getRange('E5').setValue(results.laborCost);
  calcSheet.getRange('E6').setValue(results.totalCost);
  calcSheet.getRange('E7').setValue(results.suggestedPrice);
  
  SpreadsheetApp.getUi().alert('คำนวณต้นทุน 4 สีเรียบร้อย!');
}

/**
 * ลอจิกการคำนวณหลัก
 * @param {string} printerName
 * @param {Array} filaments - [{id, weight}]
 * @param {number} wasteWeight - น้ำหนัก Wipe Tower
 */
function calculateCost(printerName, filaments, wasteWeight, printTimeHours, setupTimeMinutes, marginPercent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const inventorySheet = ss.getSheetByName('Inventory');
  const settingsData = settingsSheet.getDataRange().getValues();
  
  const electricityRate = settingsData[0][1]; 
  const laborRate = settingsData[1][1];
  const failureBuffer = settingsData[2][1] / 100;
  
  let printerPower = 0;
  let wearAndTear = 0;
  for (let i = 6; i < settingsData.length; i++) {
    if (settingsData[i][0] === printerName) {
      printerPower = settingsData[i][1];
      wearAndTear = settingsData[i][2];
      break;
    }
  }
  
  const invData = inventorySheet.getDataRange().getValues();
  let totalFilamentCost = 0;
  let mainFilamentCostPerGram = 0;

  filaments.forEach((f, index) => {
    for (let j = 1; j < invData.length; j++) {
      if (invData[j][0] === f.id) {
        const costPerGram = invData[j][6]; 
        totalFilamentCost += (f.weight * costPerGram);
        if (index === 0) mainFilamentCostPerGram = costPerGram; // ใช้ราคาเส้นแรกเป็นเกณฑ์สำหรับ Waste
        break;
      }
    }
  });
  
  // เพิ่มต้นทุน Wipe Tower (คำนวณจากราคาเฉลี่ยเส้นที่ใช้ หรือเส้นแรก)
  const wasteCost = wasteWeight * mainFilamentCostPerGram;
  totalFilamentCost += wasteCost;
  
  const electricityCost = (printerPower / 1000) * printTimeHours * electricityRate;
  const wearAndTearCost = printTimeHours * wearAndTear;
  const laborCost = ((printTimeHours + (setupTimeMinutes / 60)) * laborRate);
  
  const subtotal = totalFilamentCost + electricityCost + wearAndTearCost + laborCost;
  const totalWithBuffer = subtotal * (1 + failureBuffer);
  const suggestedPrice = totalWithBuffer * (1 + (marginPercent / 100));
  
  return {
    filamentCost: totalFilamentCost,
    electricityCost,
    wearAndTearCost,
    laborCost,
    totalCost: totalWithBuffer,
    suggestedPrice
  };
}

/**
 * ซิงค์ Dropdown (B3:B6)
 */
function syncSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const inventorySheet = ss.getSheetByName('Inventory');
  const calcSheet = ss.getSheetByName('Calculator');
  
  const printers = settingsSheet.getRange(7, 1, Math.max(1, settingsSheet.getLastRow() - 6), 1).getValues();
  const filaments = inventorySheet.getRange(2, 1, Math.max(1, inventorySheet.getLastRow() - 1), 1).getValues();
  
  const printerRule = SpreadsheetApp.newDataValidation().requireValueInList(printers).build();
  calcSheet.getRange('B2').setDataValidation(printerRule);
  
  const filamentRule = SpreadsheetApp.newDataValidation().requireValueInList(filaments).build();
  calcSheet.getRange('B3:B6').setDataValidation(filamentRule);
  
  SpreadsheetApp.getUi().alert('อัปเดต Dropdown (4 สี) เรียบร้อย!');
}

/**
 * หักสต็อกเมื่อออเดอร์เสร็จสิ้น (Version 3.1)
 */
function deductStock(projectName, printer, filaments, waste, printTime, totalCost, suggestedPrice) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const invSheet = ss.getSheetByName('Inventory');
    const orderSheet = ss.getSheetByName('Orders');
    const invData = invSheet.getDataRange().getValues();
    
    let warnings = [];
    let totalWeightUsed = 0;
    
    // 1. คำนวณน้ำหนักรวมของเส้นที่ใช้ (ไม่รวม Waste)
    const netWeight = filaments.reduce((sum, f) => sum + f.weight, 0);
    
    // 2. หักสต็อก
    filaments.forEach(f => {
      // ค้นหาแถวของเส้นลวด
      for (let i = 1; i < invData.length; i++) {
        if (invData[i][0] === f.id) {
          let currentWeight = invData[i][7]; // Column H: Remaining Weight (g)
          
          // ถ้า Remaining Weight เป็นค่าว่างหรือ 0 ให้ใช้ Total Weight (Index 4) เป็นค่าเริ่มต้น
          if (currentWeight === "" || currentWeight === 0) {
            currentWeight = invData[i][4];
          }
          
          // คำนวณส่วนแบ่ง Waste (ถ้ามี)
          const wasteShare = netWeight > 0 ? (f.weight / netWeight) * waste : 0;
          const totalUsed = f.weight + wasteShare;
          
          const newWeight = currentWeight - totalUsed;
          invSheet.getRange(i + 1, 8).setValue(newWeight);
          
          // เช็คแจ้งเตือนของใกล้หมด (Threshold: 200g)
          if (newWeight < 200) {
            warnings.push(`${f.id} เหลือเพียง ${newWeight.toFixed(1)}g`);
          }
          
          totalWeightUsed += totalUsed;
          break;
        }
      }
    });
    
    // 3. บันทึกข้อมูลลงหน้า Orders
    const filamentIds = filaments.map(f => f.id).join(', ');
    orderSheet.appendRow([
      "ORD-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmm"),
      new Date(),
      projectName,
      printer,
      filamentIds,
      totalWeightUsed.toFixed(1),
      totalCost,
      suggestedPrice,
      "Completed"
    ]);
    
    return { success: true, warnings: warnings };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
