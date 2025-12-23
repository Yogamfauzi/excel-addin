var isEditMode = false;
var currentID = "";
var masterData = {
  jadwal: [],
  leader: [],
  supervisor: [],
  machine: { data: [], map: {} },
  reject: { data: [], map: {} },
  targetOEE: { data: [], header: [] },
  sku: [],
  categories: [],
  pics: [],
  statuses: [],
  downtimeMap: {},
  downtimeMachines: [],
  categoryMapping: [],
};
var MAX_DOWNTIME_ROWS = 30;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    setupEventListeners();
    loadMasterData().then(function () {
      initializeForm();
      makeAllFieldsEditable();
      setupTimeInputAutoFormat();
    });
  }
});

function makeAllFieldsEditable() {
  var ids = ["TotalJam", "TotalPcs", "TotalWaste", "TotalActual", "TotalStandart"];

  if (masterData.categoryMapping.length > 0) {
    masterData.categoryMapping.forEach(function (cat) {
      ids.push("Total_" + cat.short);
    });
  } else {
    var defaults = ["UTND", "CIMOH", "NPT", "PM", "PS", "PCO", "BM", "OLPS", "EQFB", "LOG", "PRL", "QUAL"];
    defaults.forEach(function (d) {
      ids.push("Total_" + d);
    });
  }

  ids.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) {
      el.removeAttribute("readonly");
      el.classList.remove("bg-red", "bg-yellow", "bg-blue", "bg-grey", "bg-green");
      el.style.backgroundColor = "#ffffff";
    }
  });

  var allInputs = document.querySelectorAll("input, select, textarea");
  allInputs.forEach(function (input) {
    if (input.id && input.id.indexOf("Total_") === 0) {
      input.setAttribute("readonly", "true");
      input["style"].backgroundColor = "#f0f0f0";
    } else {
      input.removeAttribute("readonly");
      input.removeAttribute("disabled");
    }
  });
}

function handleDatalistMouseDown(el) {

}

function handleDatalistBlur(el, defaultText) {
  if (el.value.trim() === "") {
    el.placeholder = defaultText;
  }
}

function showNotification(message, type) {
  console.log(message);
  var toast = document.createElement("div");
  toast.style.cssText =
    "position:fixed; top:20px; right:20px; padding:15px; background:" +
    (type === "success" ? "#dff6dd" : "#fde7e9") +
    "; border:1px solid " +
    (type === "success" ? "#107c10" : "#a4262c") +
    "; z-index:10000; border-radius:4px; font-family: 'Segoe UI', sans-serif; font-size: 13px; font-weight: 600;";
  toast.innerText = message;
  document.body.appendChild(toast);
  setTimeout(function () {
    toast.remove();
  }, 3000);
}

async function loadMasterData() {
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItemOrNullObject("Item Master");
      await context.sync();
      if (sheet.isNullObject) return;

      var tJadwal = sheet.tables.getItemOrNullObject("JadwalTable");
      var tLeader = sheet.tables.getItemOrNullObject("LeaderTable");
      var tSpv = sheet.tables.getItemOrNullObject("SupervisorTable");
      var tMachine = sheet.tables.getItemOrNullObject("MachineTable");
      var tReject = sheet.tables.getItemOrNullObject("RejectTable");
      var tTarget = sheet.tables.getItemOrNullObject("TargetOEETable");
      var tSKU = sheet.tables.getItemOrNullObject("SKUTable");
      var tCat = sheet.tables.getItemOrNullObject("CategoryTable");
      var tPic = sheet.tables.getItemOrNullObject("PicTable");
      var tStatus = sheet.tables.getItemOrNullObject("StatusTable");

      var tables = [tJadwal, tLeader, tSpv, tMachine, tReject, tTarget, tSKU, tCat, tPic, tStatus];
      tables.forEach(function (t) {
        t.load("isNullObject");
      });
      await context.sync();

      var ranges = [];
      if (!tJadwal.isNullObject) ranges.push({ key: "jadwal", rng: tJadwal.getDataBodyRange().load("values") });
      if (!tLeader.isNullObject) ranges.push({ key: "leader", rng: tLeader.getDataBodyRange().load("values") });
      if (!tSpv.isNullObject) ranges.push({ key: "spv", rng: tSpv.getDataBodyRange().load("values") });
      if (!tSKU.isNullObject) ranges.push({ key: "sku", rng: tSKU.getDataBodyRange().load("values") });
      if (!tCat.isNullObject) ranges.push({ key: "cat", rng: tCat.getDataBodyRange().load("values") });
      if (!tPic.isNullObject) ranges.push({ key: "pic", rng: tPic.getDataBodyRange().load("values") });
      if (!tStatus.isNullObject) ranges.push({ key: "status", rng: tStatus.getDataBodyRange().load("values") });

      if (!tMachine.isNullObject) {
        ranges.push({ key: "machineBody", rng: tMachine.getDataBodyRange().load("values") });
        ranges.push({ key: "machineHead", rng: tMachine.getHeaderRowRange().load("values") });
      }
      if (!tReject.isNullObject) {
        ranges.push({ key: "rejectBody", rng: tReject.getDataBodyRange().load("values") });
        ranges.push({ key: "rejectHead", rng: tReject.getHeaderRowRange().load("values") });
      }
      if (!tTarget.isNullObject) {
        ranges.push({ key: "targetBody", rng: tTarget.getDataBodyRange().load("values") });
        ranges.push({ key: "targetHead", rng: tTarget.getHeaderRowRange().load("values") });
      }

      if (ranges.length > 0) await context.sync();

      ranges.forEach(function (item) {
        if (item.key === "jadwal") masterData.jadwal = item.rng.values;
        if (item.key === "leader") masterData.leader = item.rng.values;
        if (item.key === "spv") masterData.supervisor = item.rng.values;
        if (item.key === "sku") masterData.sku = item.rng.values;

        if (item.key === "cat") {
          masterData.categories = item.rng.values.map((r) => r[0]);
          masterData.categoryMapping = item.rng.values.map(function (row) {
            return { full: row[0], short: row[1] };
          });
        }

        if (item.key === "pic") masterData.pics = item.rng.values.map((r) => r[0]);
        if (item.key === "status") masterData.statuses = item.rng.values.map((r) => r[0]);

        if (item.key === "machineBody") masterData.machine.data = item.rng.values;
        if (item.key === "machineHead") {
          var h = item.rng.values[0];
          for (var i = 0; i < h.length; i++) masterData.machine.map[h[i]] = i;
        }
        if (item.key === "rejectBody") masterData.reject.data = item.rng.values;
        if (item.key === "rejectHead") {
          var h = item.rng.values[0];
          for (var i = 0; i < h.length; i++) masterData.reject.map[h[i]] = i;
        }
        if (item.key === "targetBody") masterData.targetOEE.data = item.rng.values;
        if (item.key === "targetHead") masterData.targetOEE.header = item.rng.values[0];
      });

      populateDatalist("SupervisorList", masterData.supervisor, 0);
      populateDatalist("LeaderList", masterData.leader, 0);
      populateDatalist("TeamList", masterData.leader, 0);
      populateDatalist("ShiftList", masterData.jadwal, 0);

      makeAllFieldsEditable();
    });
  } catch (e) {
    console.error(e);
  }
}

async function loadMachineByLine(line) {
  if (!line) return;
  masterData.downtimeMap = {};

  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItemOrNullObject("Item Master");
      await context.sync();
      if (sheet.isNullObject) return;

      var tablesToLoad = [];
      var lineClean = line.trim().toUpperCase();
      if (lineClean === "A") { tablesToLoad.push("DTASoyTable", "DTASyrupTable"); }
      else { tablesToLoad.push("DT" + lineClean + "Table"); }

      var loadedItems = [];
      tablesToLoad.forEach(function (tblName) {
        var tbl = sheet.tables.getItemOrNullObject(tblName);
        loadedItems.push({ name: tblName, obj: tbl });
        tbl.load("isNullObject");
      });
      await context.sync();

      var ranges = [];
      for (var item of loadedItems) {
        if (!item.obj.isNullObject) {
          ranges.push({
            body: item.obj.getDataBodyRange().load("values"),
            header: item.obj.getHeaderRowRange().load("values"),
          });
        }
      }
      if (ranges.length > 0) await context.sync();

      ranges.forEach(function (item) {
        var headers = item.header.values[0];
        var rows = item.body.values;
        var idxMach = -1, idxDesc = -1;

        for (var h = 0; h < headers.length; h++) {
          var head = String(headers[h]).trim().toUpperCase();
          if (head === "MACHINE") idxMach = h;
          if (head === "DESCRIPTION") idxDesc = h;
        }

        if (idxMach > -1 && idxDesc > -1) {
          rows.forEach(function (row) {
            var mKey = String(row[idxMach] || "").trim();
            var dVal = String(row[idxDesc] || "").trim();
            if (mKey && dVal) {
              if (!masterData.downtimeMap[mKey]) masterData.downtimeMap[mKey] = [];
              if (masterData.downtimeMap[mKey].indexOf(dVal) === -1) masterData.downtimeMap[mKey].push(dVal);
            }
          });
        }
      });

      refreshDowntimeDropdowns();
    });
  } catch (e) { console.error(e); }
}

function refreshDowntimeDropdowns() {
  var machineOpts = '<option value="">--Pilih Mesin--</option>';
  var addedMachines = [];

  for (var i = 1; i <= 13; i++) {
    var mName = getValue("TUTMachine" + i).trim();
    if (mName) {
      if (mName.indexOf("/") > -1) {
        var parts = mName.split("/");
        parts.forEach(function (part) {
          var cleanPart = part.trim();
          if (addedMachines.indexOf(cleanPart) === -1) {
            machineOpts += '<option value="' + cleanPart + '">' + cleanPart + "</option>";
            addedMachines.push(cleanPart);
          }
        });
      } else {
        if (addedMachines.indexOf(mName) === -1) {
          machineOpts += '<option value="' + mName + '">' + mName + "</option>";
          addedMachines.push(mName);
        }
      }
    }
  }

  var catOpts = '<option value="">--Pilih Category--</option>';
  masterData.categories.forEach(function (c) { catOpts += '<option value="' + c + '">' + c + "</option>"; });

  var picOpts = '<option value="">--Pilih PIC--</option>';
  masterData.pics.forEach(function (p) { picOpts += '<option value="' + p + '">' + p + "</option>"; });

  var statusOpts = '<option value="">--Pilih Status--</option>';
  masterData.statuses.forEach(function (s) { statusOpts += '<option value="' + s + '">' + s + "</option>"; });

  for (var j = 1; j <= MAX_DOWNTIME_ROWS; j++) {
    var dlMach = document.getElementById("machineList" + j);
    if (dlMach) dlMach.innerHTML = machineOpts;

    var dlCat = document.getElementById("CategoryList" + j);
    if (dlCat) dlCat.innerHTML = catOpts;

    var dlPic = document.getElementById("PicList" + j);
    if (dlPic) dlPic.innerHTML = picOpts;

    var dlStatus = document.getElementById("StatusList" + j);
    if (dlStatus) dlStatus.innerHTML = statusOpts;
  }
}

function updateDescriptionOptions(rowIdx) {
  var selectedMachine = getValue("Machine" + rowIdx).trim().toUpperCase();
  var descDatalist = document.getElementById("descList" + rowIdx);

  if (!descDatalist) return;

  var descOpts = "";
  if (!selectedMachine) {
    descDatalist.innerHTML = "";
    return;
  }

  for (var dbMachineName in masterData.downtimeMap) {
    var cleanDbName = dbMachineName.trim().toUpperCase();
    var isMatch = false;

    if (cleanDbName.indexOf(selectedMachine) > -1 || selectedMachine.indexOf(cleanDbName) > -1) {
      isMatch = true;
    }

    if (isMatch) {
      masterData.downtimeMap[dbMachineName].forEach(function (desc) {
        if (descOpts.indexOf('value="' + desc + '"') === -1) {
          descOpts += '<option value="' + desc + '">' + desc + "</option>";
        }
      });
    }
  }
  descDatalist.innerHTML = descOpts;
}

function setupEventListeners() {
  var editBtn = document.getElementById("EditBtn");
  if (editBtn) editBtn.onclick = handleEdit;
  var hapusBtn = document.getElementById("HapusBtn");
  if (hapusBtn) hapusBtn.onclick = handleDelete;
  var submitBtn = document.getElementById("SubmitBtn");
  if (submitBtn) submitBtn.onclick = handleSubmit;
  var addBtn = document.getElementById("AddDowntimeBtn");
  if (addBtn)
    addBtn.onclick = function () {
      addDowntimeRow(null);
    };

  document.getElementById("Line").addEventListener("change", onLineChange);
  document.getElementById("Leader").addEventListener("change", onLeaderChange);
  var dateEl = document.getElementById("Tanggal");
  if (dateEl) dateEl.addEventListener("change", onTanggalChange);
  document.getElementById("Shift").addEventListener("change", onShiftChange);
  document.getElementById("SKUCode").addEventListener("change", updateSKUName);

  var prodInputs = ["SpeedJam", "Isi1Dus", "TotalJam"];
  prodInputs.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) {
      el.addEventListener("input", function () {
        hitungTotalPcs();
        if (this.id === "SpeedJam") updateStandartFromSpeed();
      });
    }
  });

  var timeInputs = ["StartProduction", "EndProduction"];
  timeInputs.forEach(function (id) {
    var el = document.getElementById(id);
    if (el) {
      el.addEventListener("change", function () {
        hitungTotalJam();
        hitungIntervalHour();
      });
    }
  });

  for (var i = 1; i <= 15; i++) {
    var wasteEl = document.getElementById("Waste" + i);
    if (wasteEl) wasteEl.addEventListener("input", hitungTotalWaste);
  }

  for (var j = 1; j <= 20; j++) {
    var actEl = document.getElementById("Actual" + j);
    var stdEl = document.getElementById("Standart" + j);
    if (actEl) actEl.addEventListener("input", hitungDowntimePerJam);
    if (stdEl) stdEl.addEventListener("input", hitungDowntimePerJam);
  }

  setupKeyboardNavigation();
}

function updateStandartFromSpeed() {
  var speedVal = parseFloat(getValue("SpeedJam")) || 0;
  for (var i = 1; i <= 20; i++) {
    var hourVal = getValue("Hour" + i);
    if (hourVal && hourVal.indexOf("-") > -1) {
      var parts = hourVal.split("-");
      var start = parseTime(parts[0].trim());
      var end = parseTime(parts[1].trim());
      var durationMinutes = end - start;
      if (durationMinutes < 0) durationMinutes += 24 * 60;
      if (speedVal > 0) {
        var adjustedStd = (speedVal * durationMinutes) / 60;
        setValue("Standart" + i, Math.round(adjustedStd));
      } else {
        setValue("Standart" + i, "");
      }
    } else {
      setValue("Standart" + i, "NONE");
    }
    var eventMock = { target: { id: "Standart" + i } };
    hitungDowntimePerJam(eventMock);
  }
}

function convertDateInputToExcel(ymdString) {
  if (!ymdString) return "";
  var parts = ymdString.split("-");
  if (parts.length !== 3) return ymdString;
  var y = parts[0];
  var mInt = parseInt(parts[1]);
  var d = parts[2];
  var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  return d + "-" + months[mInt - 1] + "-" + y;
}

function convertExcelDateToInput(excelStr) {
  if (!excelStr) return "";
  if (typeof excelStr === "number") {
    var date = new Date((excelStr - (25567 + 2)) * 86400 * 1000);
    return date.toISOString().split("T")[0];
  }
  var parts = excelStr.split(/[-/]/);
  if (parts.length < 3) return "";
  var d = parts[0];
  var m = parts[1];
  var y = parts[2];
  var mOut = m;
  if (isNaN(m)) {
    var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
    var idx = months.findIndex((item) => item.toLowerCase() === m.toLowerCase());
    if (idx > -1) mOut = ("0" + (idx + 1)).slice(-2);
  } else {
    mOut = ("0" + m).slice(-2);
  }
  return y + "-" + mOut + "-" + ("0" + d).slice(-2);
}

function initializeForm() {
  var d = new Date();
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear();
  setValue("Tanggal", year + "-" + month + "-" + day);
  onTanggalChange();

  if (getValue("Line")) {
    onLineChange();
  }
}

function populateDatalist(datalistId, data, colIdx) {
  var datalist = document.getElementById(datalistId);
  if (!datalist) return;

  var opts = "";
  if (data) {
    data.forEach(function (row) {
      if (row[colIdx]) {
        opts += '<option value="' + row[colIdx] + '">' + row[colIdx] + "</option>";
      }
    });
  }
  datalist.innerHTML = opts;
}

function onLineChange() {
  var line = getValue("Line");
  hitungTargetOEE();
  loadHiddenMachineMaps();
  loadRejectMaps();
  loadMachineByLine(line);
}

function onLeaderChange() {
  var leader = getValue("Leader");
  setValue("Team", leader);
  loadHiddenMachineMaps();
  loadRejectMaps();
}

function onTanggalChange() {
  var tglStr = getValue("Tanggal");
  if (!tglStr) return;
  var date = new Date(tglStr);
  var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  setValue("Hari", days[date.getDay()]);
  hitungTargetOEE();
}

function onShiftChange() {
  var shiftVal = getValue("Shift");
  if (!shiftVal) return;
  var row = null;
  for (var i = 0; i < masterData.jadwal.length; i++) {
    if (masterData.jadwal[i][0] == shiftVal) {
      row = masterData.jadwal[i];
      break;
    }
  }
  if (row) {
    var startStr = formatExcelTime(row[1]);
    var endStr = formatExcelTime(row[2]);
    setValue("StartProduction", startStr);
    setValue("EndProduction", endStr);
    hitungTotalJam();
    hitungIntervalHour();
  }
}

function hitungTargetOEE() {
  var line = getValue("Line");
  var tglStr = getValue("Tanggal");
  if (!line || !tglStr) return;
  var date = new Date(tglStr);
  var months = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  var curMonth = months[date.getMonth()];
  var colIdx = -1;
  if (masterData.targetOEE.header) {
    for (var i = 0; i < masterData.targetOEE.header.length; i++) {
      if (String(masterData.targetOEE.header[i]).indexOf(curMonth) > -1) {
        colIdx = i;
        break;
      }
    }
  }
  if (colIdx > -1) {
    var row = null;
    for (var j = 0; j < masterData.targetOEE.data.length; j++) {
      if (masterData.targetOEE.data[j][0] === line) {
        row = masterData.targetOEE.data[j];
        break;
      }
    }
    if (row) {
      var val = row[colIdx];
      setValue("TargetOEE", (parseFloat(val) * 100).toFixed(2) + "%");
      return;
    }
  }
  setValue("TargetOEE", "0%");
}

function updateSKUName() {
  var code = getValue("SKUCode");
  var row = null;
  for (var i = 0; i < masterData.sku.length; i++) {
    if (masterData.sku[i][0] == code) {
      row = masterData.sku[i];
      break;
    }
  }
  if (row) setValue("SKUName", row[1]);
  else setValue("SKUName", "");
}

function hitungTotalJam() {
  var s = getValue("StartProduction");
  var e = getValue("EndProduction");
  if (!s || !e) return;
  var start = parseTime(s);
  var end = parseTime(e);
  if (end < start) end += 24 * 60;
  var diffMin = end - start;
  var hours = diffMin / 60;
  setValue("TotalJam", hours.toFixed(2));
  hitungTotalPcs();
}

function hitungIntervalHour() {
  var startStr = getValue("StartProduction");
  var endStr = getValue("EndProduction");
  if (!startStr || !endStr) return;
  var current = parseTime(startStr);
  var end = parseTime(endStr);
  if (end < current) end = end + 24 * 60;

  for (var i = 1; i <= 10; i++) {
    var hourLabel = "";
    if (current < end) {
      var next = current + 60;
      if (current % 60 === 0) next = current + 60;
      else next = (Math.floor(current / 60) + 1) * 60;
      if (next > end) next = end;
      if (next <= current) next = end;
      var sLabel = formatMinToTime(current);
      var eLabel = formatMinToTime(next);
      hourLabel = sLabel + " - " + eLabel;
      current = next;
    } else {
      if (i === 9) hourLabel = "OT1";
      else if (i === 10) hourLabel = "OT2";
      else hourLabel = "NONE";
    }
    setValue("Hour" + i, hourLabel);

    if (hourLabel.indexOf("-") === -1) {
      setValue("Standart" + i, "NONE");
      setValue("Actual" + i, "NONE");
      setValue("Quality" + i, "NONE");
      setValue("Safety" + i, "NONE");
    } else {
      if (getValue("Standart" + i) === "NONE") setValue("Standart" + i, "");
      if (getValue("Actual" + i) === "NONE") setValue("Actual" + i, "");
      var curQ = getValue("Quality" + i);
      var curS = getValue("Safety" + i);
      if (curQ === "NONE" || curQ === "") setValue("Quality" + i, "√");
      if (curS === "NONE" || curS === "") setValue("Safety" + i, "√");
    }
  }

  for (var k = 11; k <= 20; k++) {
    var sourceIdx = k - 10;
    var mirrorVal = getValue("Hour" + sourceIdx);
    setValue("Hour" + k, mirrorVal);
  }

  updateStandartFromSpeed();
}

function hitungDowntimePerJam(event) {
  var id = event.target.id;
  var idx = id.replace("Actual", "").replace("Standart", "");
  var hourStr = getValue("Hour" + idx);
  var stdVal = parseFloat(getValue("Standart" + idx)) || 0;
  var actVal = parseFloat(getValue("Actual" + idx)) || 0;

  if (!hourStr || hourStr.indexOf("-") === -1) {
    setValue("DT" + idx, 0);
  } else {
    var parts = hourStr.split("-");
    var start = parseTime(parts[0].trim());
    var end = parseTime(parts[1].trim());
    var durationMinutes = end - start;
    if (durationMinutes < 0) durationMinutes += 24 * 60;

    if (durationMinutes > 0 && stdVal > 0 && actVal < stdVal) {
      var speedPerMinute = stdVal / durationMinutes;
      var lossQty = stdVal - actVal;
      var result = lossQty / speedPerMinute;
      setValue("DT" + idx, result.toFixed(2));
    } else {
      setValue("DT" + idx, 0);
    }
  }

  var totalAct = 0;
  var totalStd = 0;
  for (var i = 1; i <= 20; i++) {
    var a = getValue("Actual" + i);
    var s = getValue("Standart" + i);
    if (a !== "NONE" && a !== "") totalAct += parseFloat(a) || 0;
    if (s !== "NONE" && s !== "") totalStd += parseFloat(s) || 0;
  }
  setValue("TotalActual", totalAct);
  setValue("TotalStandart", totalStd);
}

function hitungTotalPcs() {
  var speed = parseFloat(getValue("SpeedJam")) || 0;
  var isi = parseFloat(getValue("Isi1Dus")) || 0;
  var jam = parseFloat(getValue("TotalJam")) || 0;
  var total = speed * isi * jam;
  setValue("TotalPcs", total.toLocaleString("en-US"));
}

function hitungTotalWaste() {
  var sum = 0;
  for (var i = 1; i <= 10; i++) {
    sum += parseFloat(getValue("Waste" + i)) || 0;
  }
  setValue("TotalWaste", sum);
}

function loadHiddenMachineMaps() {
  var lineInput = getValue("Line");
  for (var x = 1; x <= 13; x++) {
    setValue("TUTMachine" + x, "");
    setValue("PDTMachine" + x, "");
    setValue("UPDTMachine" + x, "");
  }

  if (!lineInput || masterData.machine.data.length === 0) return;

  var colMap = masterData.machine.map;
  var idxLine = colMap["Line"];
  var cleanLineInput = String(lineInput).trim().toUpperCase();
  var rows = masterData.machine.data;
  var foundRow = null;

  for (var i = 0; i < rows.length; i++) {
    var mLineRaw = String(rows[i][idxLine] || "").trim().toUpperCase();
    if (mLineRaw.indexOf(cleanLineInput) > -1 || cleanLineInput.indexOf(mLineRaw) > -1) {
      foundRow = rows[i];
      break;
    }
  }

  if (foundRow) {
    for (var k = 1; k <= 13; k++) {
      var idxMach = colMap["Machine" + k];
      if (idxMach !== undefined) {
        var machName = foundRow[idxMach];
        if (machName) {
          setValue("TUTMachine" + k, machName);
          setValue("PDTMachine" + k, machName);
          setValue("UPDTMachine" + k, machName);
        }
      }
    }
  }
}

function loadRejectMaps() {
  var lineInput = getValue("Line");
  var leaderInput = getValue("Leader");
  for (var x = 1; x <= 12; x++) {
    setValue("MachineReject" + x, "");
  }
  if (!lineInput || !leaderInput || masterData.reject.data.length === 0) return;
  var colMap = masterData.reject.map;
  var idxLine = colMap["Line"];
  var idxLeader = colMap["Nama Leader"];
  if (idxLine === undefined || idxLeader === undefined) return;
  var cleanLineInput = String(lineInput).trim().toUpperCase();
  var cleanLeaderInput = String(leaderInput).trim().toUpperCase();
  var foundRow = null;
  var rows = masterData.reject.data;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var mLineRaw = String(row[idxLine] || "")
      .trim()
      .toUpperCase();
    var mLeaderRaw = String(row[idxLeader] || "")
      .trim()
      .toUpperCase();
    var mLineClean = mLineRaw.replace("LINE", "").trim();
    var isLineMatch = mLineClean === cleanLineInput || mLineClean.indexOf(cleanLineInput) > -1;
    var isLeaderMatch = false;
    if (mLeaderRaw !== "") {
      if (cleanLeaderInput.indexOf(mLeaderRaw) > -1 || mLeaderRaw.indexOf(cleanLeaderInput) > -1) {
        isLeaderMatch = true;
      }
    }
    if (isLineMatch && isLeaderMatch) {
      foundRow = row;
      break;
    }
  }
  if (foundRow) {
    for (var k = 1; k <= 12; k++) {
      var colName = "Reject" + k;
      var idxReject = colMap[colName];
      if (idxReject !== undefined) {
        var rejName = foundRow[idxReject];
        if (rejName) setValue("MachineReject" + k, rejName);
      }
    }
  }
}
function addDowntimeRow(data) {
  var idx = -1;
  for (var i = 1; i <= MAX_DOWNTIME_ROWS; i++) {
    if (!document.getElementById("DowntimeRow_" + i)) {
      idx = i;
      break;
    }
  }
  if (idx === -1) {
    showNotification("Maksimal baris downtime tercapai!", "error");
    return;
  }

  var container = document.getElementById("dynamic-downtime-container");
  var rowDiv = document.createElement("div");
  rowDiv.id = "DowntimeRow_" + idx;
  rowDiv.style.cssText = "border:1px solid #ccc; padding:10px; margin-bottom:10px; background:#fff;";

  rowDiv.innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
      <strong>Downtime #${idx}</strong>
      <div>
        <button type="button" class="btn-primary" style="padding:2px 8px; margin-right:5px; background-color:#0078d4; color:white; border:none; cursor:pointer;" onclick="copyDowntimeRow(${idx})" title="Salin Baris Ini">Salin</button>
        <button type="button" class="btn-danger" style="padding:2px 8px; background-color:#a4262c; color:white; border:none; cursor:pointer;" onclick="removeDowntimeRow(${idx})" title="Hapus Baris Ini">X</button>
      </div>
    </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
      <div><label>Category:</label>
        <input list="CategoryList${idx}" id="Category${idx}" class="d-input" style="width:100%" placeholder="--Pilih Category--" onmousedown="handleDatalistMouseDown(this)" onblur="handleDatalistBlur(this, '--Pilih Category--')">
        <datalist id="CategoryList${idx}"></datalist>
      </div>
      <div><label>Machine:</label>
        <input list="machineList${idx}" id="Machine${idx}" class="d-input" style="width:100%" placeholder="--Pilih Mesin--" onmousedown="handleDatalistMouseDown(this)" onblur="handleDatalistBlur(this, '--Pilih Mesin--')" onchange="updateDescriptionOptions(${idx})">
        <datalist id="machineList${idx}"></datalist>
      </div>
      <div style="grid-column: span 2;"><label>Description:</label>
        <input list="descList${idx}" id="Description${idx}" class="d-input" style="width:100%" placeholder="--Pilih/Ketik Deskripsi--" onmousedown="handleDatalistMouseDown(this)" onblur="handleDatalistBlur(this, '--Pilih/Ketik Deskripsi--')">
        <datalist id="descList${idx}"></datalist>
      </div>
      <div style="grid-column: span 2;"><label>Action:</label>
        <input id="Action${idx}" type="text" class="d-input" style="width:100%" placeholder="Isi Action...">
      </div>
      <div><label>Start (Jam):</label>
        <input id="JamStart${idx}" type="text" maxlength="5" placeholder="HH:MM" class="d-input" style="width:100%" data-time-input="true">
      </div>
      <div><label>Durasi (Menit):</label>
        <input id="Durasi${idx}" type="number" min="0" class="d-input" style="width:100%">
      </div>
      <div><label>PIC:</label>
        <input list="PicList${idx}" id="Pic${idx}" class="d-input" style="width:100%" placeholder="--Pilih PIC--" onmousedown="handleDatalistMouseDown(this)" onblur="handleDatalistBlur(this, '--Pilih PIC--')">
        <datalist id="PicList${idx}"></datalist>
      </div>
      <div><label>Status:</label>
        <input list="StatusList${idx}" id="Status${idx}" class="d-input" style="width:100%" placeholder="--Pilih Status--" onmousedown="handleDatalistMouseDown(this)" onblur="handleDatalistBlur(this, '--Pilih Status--')">
        <datalist id="StatusList${idx}"></datalist>
      </div>
    </div>
  `;

  container.appendChild(rowDiv);

  var inputs = rowDiv.querySelectorAll(".d-input");
  inputs.forEach(input => {
    input.addEventListener("change", hitungAgregasi);
    input.addEventListener("input", hitungAgregasi);
  });

  if (data) {
    setValue("Category" + idx, data.Category || "");
    setValue("Machine" + idx, data.Machine || "");
    updateDescriptionOptions(idx);
    setValue("Description" + idx, data.Description || "");
    setValue("Durasi" + idx, data.Durasi || "");
    setValue("Action" + idx, data.Action || "");
    setValue("Pic" + idx, data.Pic || "");
    setValue("Status" + idx, data.Status || "");
    setValue("JamStart" + idx, normalizeTimeInput(data.JamStart || ""));
  }

  refreshDowntimeDropdowns();
  setupDynamicTimeInputs();
  setupKeyboardNavigation();

  if (data) hitungAgregasi();
}

function removeDowntimeRow(idx) {
  var row = document.getElementById("DowntimeRow_" + idx);
  if (row) {
    row.remove();
    hitungAgregasi();
  }
}

function copyDowntimeRow(sourceIdx) {
  var data = {
    Category: getValue("Category" + sourceIdx),
    Machine: getValue("Machine" + sourceIdx),
    Description: getValue("Description" + sourceIdx),
    Action: getValue("Action" + sourceIdx),
    Durasi: getValue("Durasi" + sourceIdx),
    JamStart: getValue("JamStart" + sourceIdx),
    Pic: getValue("Pic" + sourceIdx),
    Status: getValue("Status" + sourceIdx)
  };

  addDowntimeRow(data);
  showNotification("Baris #" + sourceIdx + " berhasil disalin!", "success");
}

function getCategoryPrefix(catInput) {
  if (!catInput) return "";
  var cleanCat = catInput.trim().toUpperCase();
  var found = masterData.categoryMapping.find(function (item) {
    return cleanCat === item.full.toUpperCase();
  });
  if (found) return found.short;
  return "";
}
function hitungAgregasi() {
  for (var m = 1; m <= 13; m++) {
    masterData.categoryMapping.forEach(function (c) { setValue(c.short + m, 0); });
  }

  var aggregatedTotals = {};
  masterData.categoryMapping.forEach(function (c) { aggregatedTotals[c.short] = 0; });

  for (var i = 1; i <= MAX_DOWNTIME_ROWS; i++) {
    if (!document.getElementById("DowntimeRow_" + i)) continue;

    var catVal = getValue("Category" + i).trim().toUpperCase();
    var machVal = getValue("Machine" + i).trim().toUpperCase();
    var durVal = parseFloat(getValue("Durasi" + i)) || 0;

    if (catVal && machVal && durVal > 0) {
      var prefix = getCategoryPrefix(catVal);
      if (prefix !== "") {
        var targetIndex = -1;

        for (var k = 1; k <= 13; k++) {
          var masterMachine = getValue("TUTMachine" + k).trim().toUpperCase();

          if (machVal === masterMachine) {
            targetIndex = k;
            break;
          }

          if (masterMachine.indexOf("/") > -1) {
            var parts = masterMachine.split("/");
            var isMatch = false;
            for (var p = 0; p < parts.length; p++) {
              if (parts[p].trim() === machVal) {
                isMatch = true;
                break;
              }
            }
            if (isMatch) {
              targetIndex = k;
              break;
            }
          }
        }

        if (targetIndex !== -1) {
          var targetID = prefix + targetIndex;
          var currentVal = parseFloat(getValue(targetID)) || 0;
          setValue(targetID, currentVal + durVal);
          if (aggregatedTotals[prefix] !== undefined) aggregatedTotals[prefix] += durVal;
        }
      }
    }
  }

  for (var key in aggregatedTotals) {
    setValue("Total_" + key, aggregatedTotals[key]);
  }
}

function generateCatatanLogic() {
  var noteResult = "";
  var downtimeGroups = {};
  var totalDowntimeAll = 0;

  for (var i = 1; i <= MAX_DOWNTIME_ROWS; i++) {
    var row = document.getElementById("DowntimeRow_" + i);
    if (row) {
      var cat = getValue("Category" + i)
        .trim()
        .toUpperCase();
      var mach = getValue("Machine" + i).trim();
      var desc = getValueOrNone("Description" + i);
      if (desc === "NONE") desc = "";
      var dur = parseFloat(getValue("Durasi" + i)) || 0;

      if (mach && dur > 0) {
        if (!downtimeGroups[cat]) {
          downtimeGroups[cat] = { total: 0, items: [] };
        }
        downtimeGroups[cat].total += dur;
        downtimeGroups[cat].items.push(mach + ": " + desc);
        totalDowntimeAll += dur;
      }
    }
  }

  var priorityKeys = ["TEMUAN ABNORMALITY", "ISSUE SAFETY", "ISSUE QUALITY"];

  function formatBlock(catName, group) {
    var blk = catName + " (Total " + group.total + " Menit)\n";
    group.items.forEach(function (item, idx) {
      blk += "    " + (idx + 1) + ". " + item + "\n";
    });
    blk += "\n";
    return blk;
  }

  priorityKeys.forEach(function (key) {
    if (downtimeGroups[key]) {
      noteResult += formatBlock(key, downtimeGroups[key]);
      delete downtimeGroups[key];
    }
  });

  for (var key in downtimeGroups) {
    noteResult += formatBlock(key, downtimeGroups[key]);
  }

  var totalStd = parseFloat(getValue("TotalStandart")) || 0;
  var totalAct = parseFloat(getValue("TotalActual")) || 0;
  var targetOee = parseFloat(getValue("TargetOEE").replace("%", "")) || 85;

  var oeePct = 0;
  if (totalStd > 0) oeePct = (totalAct / totalStd) * 100;

  var statusOEE = oeePct >= targetOee ? "Tercapai" : "Tidak Tercapai";
  var selisih = Math.abs(oeePct - targetOee).toFixed(2);
  var ketSelisih = oeePct >= targetOee ? "Lebih" : "Kurang";

  var footer = "\n---\n";
  footer += "Pencapaian OEE : " + oeePct.toFixed(2) + "%\n";
  footer += "Target OEE : " + targetOee + "% " + ketSelisih + " " + selisih + "% " + statusOEE;

  return noteResult + footer;
}

async function handleEdit() {
  var id = getValue("KolomEdit");
  if (!id) {
    showNotification("Masukkan ID untuk diedit!", "error");
    return;
  }
  await Excel.run(async function (context) {
    var sheet = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
    await context.sync();
    if (sheet.isNullObject) {
      showNotification("Sheet Input Shiftly tidak ditemukan", "error");
      return;
    }
    var tbl = sheet.tables.getItemOrNullObject("TableLaporanAkhir");
    await context.sync();
    if (tbl.isNullObject) {
      showNotification("Tabel Laporan Akhir tidak ditemukan", "error");
      return;
    }

    var headerRange = tbl.getHeaderRowRange().load("values");
    var bodyRange = tbl.getDataBodyRange().load("values");
    await context.sync();

    var sourceColIdx = -1;
    var headers = headerRange.values[0];
    for (var h = 0; h < headers.length; h++) {
      if (String(headers[h]).trim().toLowerCase() === "source") {
        sourceColIdx = h;
        break;
      }
    }

    if (sourceColIdx === -1) {
      showNotification("Kolom Source tidak ditemukan!", "error");
      return;
    }

    var found = false;
    var rowData = null;

    for (var i = 0; i < bodyRange.values.length; i++) {
      if (bodyRange.values[i][sourceColIdx] === id) {
        rowData = bodyRange.values[i];
        found = true;
        break;
      }
    }

    if (!found) {
      showNotification("Data tidak ditemukan!", "error");
      return;
    }

    isEditMode = true;
    currentID = id;

    var map = {};
    for (var k = 0; k < headers.length; k++) {
      map[headers[k]] = k;
    }

    function getData(colName) {
      if (map[colName] !== undefined) {
        var val = rowData[map[colName]];
        return val !== null && val !== undefined ? val : "";
      }
      return "";
    }

    function getFuzzyData(name) {
      var cleanName = name.toUpperCase().replace(/[^A-Z0-9]/g, "");
      for (var k = 0; k < headers.length; k++) {
        var headerClean = String(headers[k])
          .toUpperCase()
          .replace(/[^A-Z0-9]/g, "");
        if (headerClean === cleanName) {
          var val = rowData[k];
          if (val !== null && val !== undefined && val !== "NONE" && val !== "") {
            return val;
          }
        }
      }
      return null;
    }

    function getExactData(name) {
      if (map[name] !== undefined) {
        var val = rowData[map[name]];
        return val;
      }
      return null;
    }

    function isValid(val) {
      return val !== null && val !== undefined && val !== "" && val !== "NONE";
    }

    setValue("Tanggal", convertExcelDateToInput(getData("Date")));

    var date = new Date(getValue("Tanggal"));
    var days = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    setValue("Hari", days[date.getDay()]);

    setValue("Shift", getData("Shift"));
    setValue("Line", getData("Line"));

    await loadMachineByLine(getData("Line"));
    hitungTargetOEE();

    setValue("Leader", getData("Leader"));
    setValue("SKUCode", getData("SKU Code"));
    updateSKUName();

    setValue("Planning", getData("Plan"));
    setValue("TanggalMasak", getData("Tanggal Masak"));
    setValue("EXPCode", getData("EXP Code"));
    setValue("NoTangki", getData("No Tangki"));
    setValue("Supervisor", getData("SPV"));
    setValue("Team", getData("Team"));
    setValue("SpeedJam", getData("Speed / Jam"));
    setValue("Isi1Dus", getData("Isi 1 Dus"));

    var soVal = getData("No SO");
    if (soVal) {
      soVal = String(soVal).replace("No SO ", "").replace("NONE", "");
      setValue("NoSO", soVal);
    }

    setValue("StartProduction", formatExcelTime(getData("Start")));
    setValue("EndProduction", formatExcelTime(getData("Finish")));

    hitungTotalJam();
    hitungIntervalHour();

    setValue("FreqBreakdown", getData("Freq. Breakdown") || 0);
    setValue("FreqOpsFailure", getData("Freq. Ops Failure") || 0);
    setValue("TotalHead", getData("Total Head") || 0);
    setValue("TotalSafety", getData("Total Safety") || 0);
    setValue("TotalQuality", getData("Total Quality") || 0);
    setValue("CederaKecil", getData("Cedera Kecil") || 0);
    setValue("CederaTercatat", getData("Cedera Tercatat") || 0);
    setValue("HampirCelaka", getData("Hampir Celaka") || 0);
    setValue("FSQHabit", getData("FSQ Habit") || 0);
    setValue("InsidenKualitas", getData("Insiden Kualitas") || 0);
    setValue("CIL", getData("CIL") || 0);
    setValue("Centerline1", getData("Centerline (Persentase)") || 0);
    setValue("Centerline2", getData("Centerline (Jumlah)") || 0);
    setValue("Skor5S", getData("Skor 5S") || 0);
    setValue("TemuanAbnormality", getData("Temuan Abnormality") || 0);
    setValue("PerbaikanAbnormality", getData("Perbaikan Abnormality") || 0);

    if (masterData.categoryMapping.length > 0) {
      masterData.categoryMapping.forEach(function (cat) {
        var val = null;
        val = getExactData(cat.full);
        if (!isValid(val)) val = getExactData(cat.short);
        if (!isValid(val)) val = getFuzzyData(cat.full);
        if (!isValid(val)) val = getFuzzyData(cat.short);
        if (!isValid(val)) val = 0;
        setValue("Total_" + cat.short, val);
      });
    }

    for (var j = 1; j <= 20; j++) {
      setValue("Actual" + j, getData("Actual(" + j + ")"));
      setValue("Quality" + j, getData("Quality(" + j + ")"));
      setValue("Safety" + j, getData("Safety(" + j + ")"));
      setValue("Standart" + j, getData("Standart(" + j + ")"));
      var eventMock = { target: { id: "Actual" + j } };
      hitungDowntimePerJam(eventMock);
    }

    for (var w = 1; w <= 15; w++) {
      setValue("Waste" + w, getData("Waste(" + w + ")"));
    }
    hitungTotalWaste();

    setValue("TotalPcs", getData("Total Pcs"));

    await loadChildData(context, id);

    loadHiddenMachineMaps();
    loadRejectMaps();

    makeAllFieldsEditable();

    showNotification("Data berhasil dimuat untuk diedit.", "success");
    document.getElementById("KolomEdit").scrollIntoView();
  });
}
async function loadChildData(context, id) {
  var sheetDown = context.workbook.worksheets.getItemOrNullObject("Input Downtime");
  await context.sync();
  if (sheetDown.isNullObject) return;
  var tblDown = sheetDown.tables.getItemOrNullObject("DowntimeTable");
  var tblDetail = sheetDown.tables.getItemOrNullObject("DetailDowntimeTable");
  var tblReject = sheetDown.tables.getItemOrNullObject("IsiRejectTable");
  await context.sync();
  document.getElementById("dynamic-downtime-container").innerHTML = "";
  if (!tblDown.isNullObject) {
    var headerRange = tblDown.getHeaderRowRange().load("values");
    var bodyRange = tblDown.getDataBodyRange().load("values");
    await context.sync();
    var headers = headerRange.values[0];
    var map = {};
    for (var h = 0; h < headers.length; h++) map[headers[h]] = h;

    if (map["Source"] !== undefined) {
      var rows = bodyRange.values;
      rows.forEach(function (row) {
        if (row[map["Source"]] === id) {
          addDowntimeRow({
            Category: row[map["Category"]],
            Machine: row[map["Machine"]],
            Description: row[map["Description"]],
            Durasi: row[map["Durasi"]],
            Action: row[map["Action"]],
            Pic: row[map["PIC"]],
            Status: row[map["Status"]],
            JamStart: formatExcelTime(row[map["Start"]]),
          });
        }
      });
    }
  }
  if (!tblDetail.isNullObject) {
    var detailHeaderRange = tblDetail.getHeaderRowRange().load("values");
    var detailBodyRange = tblDetail.getDataBodyRange().load("values");
    await context.sync();
    var detailHeaders = detailHeaderRange.values[0];
    var detailMap = {};
    for (var h = 0; h < detailHeaders.length; h++) detailMap[detailHeaders[h]] = h;

    if (detailMap["Source"] !== undefined) {
      var detailRows = detailBodyRange.values;
      for (var i = 0; i < detailRows.length; i++) {
        if (detailRows[i][detailMap["Source"]] === id) {
          var detailRow = detailRows[i];

          if (masterData.categoryMapping.length > 0) {
            masterData.categoryMapping.forEach(function (cat) {
              for (var v = 1; v <= 13; v++) {
                var colName = cat.short + v;
                if (detailMap[colName] !== undefined) {
                  setValue(colName, detailRow[detailMap[colName]]);
                }
              }
            });
          }
          break;
        }
      }
    }
  }
  if (!tblReject.isNullObject) {
    var rejectHeaderRange = tblReject.getHeaderRowRange().load("values");
    var rejectBodyRange = tblReject.getDataBodyRange().load("values");
    await context.sync();
    var rejectHeaders = rejectHeaderRange.values[0];
    var rejectMap = {};
    for (var h = 0; h < rejectHeaders.length; h++) rejectMap[rejectHeaders[h]] = h;

    if (rejectMap["Source"] !== undefined) {
      var rejectRows = rejectBodyRange.values;
      for (var i = 0; i < rejectRows.length; i++) {
        if (rejectRows[i][rejectMap["Source"]] === id) {
          var rejectRow = rejectRows[i];

          for (var r = 1; r <= 12; r++) {
            if (rejectMap["Isi" + r] !== undefined) {
              setValue("Reject" + r, rejectRow[rejectMap["Isi" + r]]);
            }
          }
          break;
        }
      }
    }
  }
}
async function handleSubmit() {
  var valLine = getValue("Line");
  var valDate = getValue("Tanggal");
  var sProd = getValue("StartProduction");
  var eProd = getValue("EndProduction");

  if (!valLine || !valDate) {
    showNotification("Line dan Tanggal harus diisi!", "error");
    return;
  }

  if (sProd.length !== 5 || eProd.length !== 5) {
    showNotification("Format Jam Start/Finish salah! Harus HH:MM (Contoh: 07:00)", "error");
    return;
  }

  for (var m = 1; m <= MAX_DOWNTIME_ROWS; m++) {
    var row = document.getElementById("DowntimeRow_" + m);
    if (row) {
      var dur = parseFloat(getValue("Durasi" + m)) || 0;
      var cat = getValue("Category" + m);
      var mac = getValue("Machine" + m);
      if (dur > 0 && (cat === "" || mac === "")) {
        showNotification("Downtime #" + m + " memiliki durasi tapi Kategori/Mesin kosong. Harap lengkapi!", "error");
        return;
      }
    }
  }

  var btn = document.getElementById("SubmitBtn");
  if (btn) {
    btn["innerText"] = "⚙️ Menyimpan & Menghitung OEE...";
    btn["disabled"] = true;
  }

  try {
    await Excel.run(async function (context) {
      var saveID = isEditMode ? currentID : generateUniqueID();
      var rawDate = getValue("Tanggal");
      var strDate = convertDateInputToExcel(rawDate);

      var rawParts = rawDate.split("-");
      var justDay = rawParts[2];

      var rawDateMasak = getValue("TanggalMasak");
      var dVal = rawParts[2];
      var mVal = rawParts[1];
      var yVal = rawParts[0];

      var targetOEEVal = parseFloat(getValue("TargetOEE").replace("%", "")) || 0;
      if (targetOEEVal > 1) targetOEEVal = targetOEEVal / 100;

      var noSOVal = getValue("NoSO");
      if (noSOVal && noSOVal.toUpperCase().indexOf("NO SO") === -1) noSOVal = "No SO " + noSOVal;
      else if (!noSOVal) noSOVal = "NONE";

      var dataMain = {
        Source: saveID,
        Date: strDate,
        "Date(0)": justDay,
        Shift: getValueOrNone("Shift"),
        Line: getValueOrNone("Line"),
        Leader: getValueOrNone("Leader"),
        "SKU Code": getValueOrNone("SKUCode"),
        Plan: getValue("Planning"),
        "Tanggal Masak": rawDateMasak,
        "EXP Code": getValueOrNone("EXPCode"),
        "No Tangki": getValueOrNone("NoTangki"),
        SPV: getValueOrNone("Supervisor"),
        Team: getValueOrNone("Team"),
        "Speed / Jam": getValue("SpeedJam"),
        "Isi 1 Dus": getValue("Isi1Dus"),
        "No SO": noSOVal,
        Start: getValue("StartProduction"),
        Finish: getValue("EndProduction"),
        "Freq. Breakdown": getValue("FreqBreakdown") || 0,
        "Freq. Ops Failure": getValue("FreqOpsFailure") || 0,
        "Total Head": getValue("TotalHead") || 0,
        "Total Safety": getValue("TotalSafety") || 0,
        "Total Quality": getValue("TotalQuality") || 0,
        "Cedera Kecil": getValue("CederaKecil") || 0,
        "Cedera Tercatat": getValue("CederaTercatat") || 0,
        "Hampir Celaka": getValue("HampirCelaka") || 0,
        "FSQ Habit": getValue("FSQHabit") || 0,
        "Insiden Kualitas": getValue("InsidenKualitas") || 0,
        CIL: getValue("CIL") || 0,
        "Centerline (Persentase)": getValue("Centerline1") || 0,
        "Centerline (Jumlah)": getValue("Centerline2") || 0,
        "Skor 5S": getValue("Skor5S") || 0,
        "Temuan Abnormality": getValue("TemuanAbnormality") || 0,
        "Perbaikan Abnormality": getValue("PerbaikanAbnormality") || 0,
        "Target OEE": targetOEEVal,
        Actual: getValue("TotalActual") || 0,
        Waste: getValue("TotalWaste") || 0,
        "Total Pcs": getValue("TotalPcs").replace(/,/g, "") || 0,
        Catatan: "",
      };

      if (masterData.categoryMapping.length > 0) {
        masterData.categoryMapping.forEach(function (cat) {
          var val = parseFloat(getValue("Total_" + cat.short)) || 0;
          dataMain[cat.full] = val;
        });
      }

      for (var h = 1; h <= 20; h++) {
        dataMain["Hour(" + h + ")"] = getValue("Hour" + h);
        dataMain["Actual(" + h + ")"] = getValue("Actual" + h);
        dataMain["Quality(" + h + ")"] = getValueOrNone("Quality" + h);
        dataMain["Safety(" + h + ")"] = getValueOrNone("Safety" + h);
        dataMain["Standart(" + h + ")"] = getValue("Standart" + h);
      }
      for (var w = 1; w <= 15; w++) {
        dataMain["Waste(" + w + ")"] = parseFloat(getValue("Waste" + w)) || 0;
      }

      var sheetMain = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      var sheetDown = context.workbook.worksheets.getItemOrNullObject("Input Downtime");
      await context.sync();

      if (sheetMain.isNullObject) throw new Error("Sheet 'Input Shiftly' tidak ditemukan!");
      if (sheetDown.isNullObject) throw new Error("Sheet 'Input Downtime' tidak ditemukan!");

      var tblMain = sheetMain.tables.getItemOrNullObject("TableLaporanAkhir");
      var tblDowntime = sheetDown.tables.getItemOrNullObject("DowntimeTable");
      var tblDetailDT = sheetDown.tables.getItemOrNullObject("DetailDowntimeTable");
      var tblReject = sheetDown.tables.getItemOrNullObject("IsiRejectTable");

      await context.sync();
      if (tblMain.isNullObject) throw new Error("Tabel 'TableLaporanAkhir' tidak ditemukan!");

      var colMain = tblMain.columns.load("items/name");
      var colDT = tblDowntime.columns.load("items/name");
      var colDetail = tblDetailDT.columns.load("items/name");
      var colReject = tblReject.columns.load("items/name");
      await context.sync();

      var rowMainArray = mapDataToRow(colMain, dataMain);
      if (isEditMode) await deleteRowByID(context, tblMain, saveID);
      tblMain.rows.add(null, [rowMainArray]);

      await context.sync();

      var idxOEE = -1, idxCatatan = -1, idxSource = -1;
      for (var c = 0; c < colMain.items.length; c++) {
        var cName = colMain.items[c].name.toUpperCase();
        if (cName === "OEE") idxOEE = c;
        if (cName === "CATATAN") idxCatatan = c;
        if (cName === "SOURCE") idxSource = c;
      }

      if (idxOEE !== -1 && idxCatatan !== -1 && idxSource !== -1) {
        var bodyRange = tblMain.getDataBodyRange().load("values");
        await context.sync();
        var rowIndex = -1;
        var oeeResult = 0;
        for (var r = 0; r < bodyRange.values.length; r++) {
          if (bodyRange.values[r][idxSource] === saveID) {
            rowIndex = r;
            oeeResult = bodyRange.values[r][idxOEE];
            break;
          }
        }

        if (rowIndex !== -1) {
          var currentNote = generateCatatanLogic();
          if (currentNote.length > 32000) currentNote = currentNote.substring(0, 32000) + "...";

          var cellCatatan = tblMain.getDataBodyRange().getCell(rowIndex, idxCatatan);
          cellCatatan.values = [[currentNote]];
          cellCatatan.format.wrapText = false;
        }
      }

      if (isEditMode) await deleteRowByID(context, tblDowntime, saveID);
      var rowsDTArray = [];
      for (var m = 1; m <= MAX_DOWNTIME_ROWS; m++) {
        if (document.getElementById("DowntimeRow_" + m)) {
          var machineName = getValue("Machine" + m);
          if (machineName && machineName.trim() !== "") {
            var dataDT = {
              Source: saveID,
              "Date(1)": dVal,
              "Month(1)": mVal,
              "Year(1)": yVal,
              "Line(1)": getValue("Line"),
              Category: getValue("Category" + m),
              Machine: machineName,
              Description: getValueOrNone("Description" + m),
              Action: getValue("Action" + m),
              PIC: getValue("Pic" + m),
              Status: getValue("Status" + m),
              Start: getValue("JamStart" + m),
              Durasi: getValue("Durasi" + m),
            };
            rowsDTArray.push(mapDataToRow(colDT, dataDT));
          }
        }
      }
      if (rowsDTArray.length > 0) tblDowntime.rows.add(null, rowsDTArray);

      var dataDetail = { Source: saveID };
      for (var machIdx = 1; machIdx <= 13; machIdx++) {
        dataDetail["Machine" + machIdx] = getValue("TUTMachine" + machIdx);
      }
      if (masterData.categoryMapping.length > 0) {
        masterData.categoryMapping.forEach(function (cat) {
          for (var v = 1; v <= 13; v++) {
            var valDetail = getValue(cat.short + v);
            dataDetail[cat.short + v] = valDetail === "" || valDetail === null ? 0 : valDetail;
          }
        });
      }
      var rowDetailArray = mapDataToRow(colDetail, dataDetail);
      if (isEditMode) await deleteRowByID(context, tblDetailDT, saveID);
      tblDetailDT.rows.add(null, [rowDetailArray]);

      var dataReject = { Source: saveID };
      for (var r = 1; r <= 12; r++) {
        dataReject["Reject" + r] = getValueOrNone("MachineReject" + r);
        dataReject["Isi" + r] = parseFloat(getValue("Reject" + r)) || 0;
      }
      var rowRejectArray = mapDataToRow(colReject, dataReject);
      if (isEditMode) await deleteRowByID(context, tblReject, saveID);
      tblReject.rows.add(null, [rowRejectArray]);

      await context.sync();

      showNotification(isEditMode ? "Data berhasil diperbarui!" : "Data berhasil disimpan!", "success");
      isEditMode = false;
      currentID = "";
      setTimeout(function () { location.reload(); }, 2000);

    });
  } catch (error) {
    console.error(error);
    var msg = error.message;
    if (error.code === "InvalidArgument") {
      msg = "ERROR KOLOM EXCEL! Jumlah kolom di coding tidak sama dengan di Excel. Cek apakah ada kolom baru/hilang di Tabel Excel.";
    } else if (msg.indexOf("Sheet") > -1) {
      msg = "ERROR SHEET: " + msg;
    } else {
      msg = "Gagal Menyimpan: " + msg;
    }

    showNotification(msg, "error");
  } finally {
    if (btn) {
      btn["innerText"] = "💾 SIMPAN DATA";
      btn["disabled"] = false;
    }
  }
}

function mapDataToRow(excelColumns, dataObject) {
  var rowArray = [];
  var colCount = excelColumns.items.length;

  for (var j = 0; j < colCount; j++) {
    var colName = excelColumns.items[j].name;

    if (dataObject.hasOwnProperty(colName)) {
      var value = dataObject[colName];
      rowArray.push(value === undefined ? null : value);
    } else {
      rowArray.push(null);
    }
  }
  return rowArray;
}

var pendingDeleteID = "";

async function handleDelete() {
  console.log("=== FUNGSI HAPUS DIPANGGIL ===");

  var id = getValue("KolomEdit");
  console.log("ID yang diambil:", id);

  if (!id) {
    showNotification("Masukkan ID untuk dihapus!", "error");
    return;
  }

  pendingDeleteID = id;
  document.getElementById("modalIDDisplay").innerText = id;
  document.getElementById("deleteModal").style.display = "flex";
}

function closeDeleteModal() {
  document.getElementById("deleteModal").style.display = "none";
  pendingDeleteID = "";
  showNotification("Penghapusan dibatalkan", "error");
}

async function confirmDelete() {
  document.getElementById("deleteModal").style.display = "none";

  var id = pendingDeleteID;
  pendingDeleteID = "";

  if (!id) {
    showNotification("ID tidak valid!", "error");
    return;
  }

  console.log("Konfirmasi diterima, melanjutkan hapus untuk ID:", id);
  showNotification("🔄 Memproses penghapusan...", "success");

  try {
    await Excel.run(async function (context) {
      console.log("Masuk ke Excel.run");

      var sheetMain = context.workbook.worksheets.getItemOrNullObject("Input Shiftly");
      var sheetDown = context.workbook.worksheets.getItemOrNullObject("Input Downtime");
      await context.sync();

      if (sheetMain.isNullObject || sheetDown.isNullObject) {
        showNotification("Sheet tidak ditemukan", "error");
        return;
      }

      var tblMain = sheetMain.tables.getItemOrNullObject("TableLaporanAkhir");
      var tblDown = sheetDown.tables.getItemOrNullObject("DowntimeTable");
      var tblDetail = sheetDown.tables.getItemOrNullObject("DetailDowntimeTable");
      var tblReject = sheetDown.tables.getItemOrNullObject("IsiRejectTable");
      await context.sync();

      var deletedCount = 0;

      console.log("Mulai hapus dari tabel...");

      if (!tblMain.isNullObject) {
        console.log("Hapus dari TableLaporanAkhir...");
        var count = await deleteRowByID(context, tblMain, id);
        console.log("Terhapus dari TableLaporanAkhir:", count);
        deletedCount += count;
      }
      if (!tblDown.isNullObject) {
        console.log("Hapus dari DowntimeTable...");
        var count = await deleteRowByID(context, tblDown, id);
        console.log("Terhapus dari DowntimeTable:", count);
        deletedCount += count;
      }
      if (!tblDetail.isNullObject) {
        console.log("Hapus dari DetailDowntimeTable...");
        var count = await deleteRowByID(context, tblDetail, id);
        console.log("Terhapus dari DetailDowntimeTable:", count);
        deletedCount += count;
      }
      if (!tblReject.isNullObject) {
        console.log("Hapus dari IsiRejectTable...");
        var count = await deleteRowByID(context, tblReject, id);
        console.log("Terhapus dari IsiRejectTable:", count);
        deletedCount += count;
      }

      await context.sync();

      console.log("Total baris terhapus:", deletedCount);

      if (deletedCount > 0) {
        showNotification("✅ BERHASIL! Terhapus " + deletedCount + " baris dengan ID: " + id, "success");
        setTimeout(function () {
          location.reload();
        }, 2000);
      } else {
        showNotification("❌ ID '" + id + "' tidak ditemukan di database!", "error");
      }
    });
  } catch (e) {
    console.error("Error saat hapus:", e);
    showNotification("Gagal hapus data: " + e.message, "error");
  }
}
function getValue(id) {
  var el = document.getElementById(id);
  if (el) return el["value"] || "";
  return "";
}
function getValueOrNone(id) {
  var val = getValue(id);
  if (val === "" || val === null) return "NONE";
  return val;
}
function setValue(id, val) {
  var el = document.getElementById(id);
  if (el) el["value"] = val !== null && val !== undefined ? val : "";
}
function clearHiddenMachineFields() {
  for (var i = 1; i <= 13; i++) {
    setValue("TUTMachine" + i, "");
    setValue("PDTMachine" + i, "");
    setValue("UPDTMachine" + i, "");
  }
}
function generateUniqueID() {
  var d = new Date();
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear().toString().substr(-2);
  var hour = ("0" + d.getHours()).slice(-2);
  var minute = ("0" + d.getMinutes()).slice(-2);
  var datePart = "" + day + month + year + "-" + hour + minute;
  var leaderVal = getValue("Leader") || "NONAME";
  var leader = leaderVal.split(" ")[0].toUpperCase();
  var rand = Math.floor(Math.random() * 900) + 100;
  return datePart + "-" + leader + "-" + rand;
}
function parseTime(timeStr) {
  if (!timeStr) return 0;
  var parts = timeStr.split(":");
  var h = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  return h * 60 + m;
}
function formatMinToTime(totalMin) {
  var h = Math.floor(totalMin / 60) % 24;
  var m = totalMin % 60;
  var hStr = ("0" + h).slice(-2);
  var mStr = ("0" + m).slice(-2);
  return hStr + ":" + mStr;
}
function getMonthNameFromDDMMYYYY(dateStr) {
  if (!dateStr) return "";
  var parts = dateStr.split("-");
  if (parts.length < 2) return "";
  return parts[1] || "";
}
function getYearNumFromDDMMYYYY(dateStr) {
  if (!dateStr) return "";
  var parts = dateStr.split("-");
  if (parts.length < 3) return "";
  return parts[2];
}
function formatExcelTime(excelVal) {
  if (typeof excelVal === "number") {
    var totalSeconds = Math.round(excelVal * 86400);
    var h = Math.floor(totalSeconds / 3600);
    var m = Math.floor((totalSeconds % 3600) / 60);
    var hStr = ("0" + h).slice(-2);
    var mStr = ("0" + m).slice(-2);
    return hStr + ":" + mStr;
  }
  return excelVal;
}
async function deleteRowByID(context, table, id) {
  var bodyRange = table.getDataBodyRange().load("values");
  var headerRange = table.getHeaderRowRange().load("values");
  await context.sync();

  if (!headerRange.values || headerRange.values.length === 0) {
    return 0;
  }

  var headers = headerRange.values[0];
  var sourceColIdx = -1;

  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h]).trim().toLowerCase() === "source") {
      sourceColIdx = h;
      break;
    }
  }

  if (sourceColIdx === -1) {
    console.log("Kolom Source tidak ditemukan di tabel");
    return 0;
  }

  var rowsToDelete = [];
  for (var i = 0; i < bodyRange.values.length; i++) {
    var cellValue = String(bodyRange.values[i][sourceColIdx]).trim();
    var searchId = String(id).trim();

    if (cellValue === searchId) {
      rowsToDelete.push(i);
      console.log("Ditemukan data untuk dihapus di baris index: " + i);
    }
  }

  if (rowsToDelete.length === 0) {
    console.log("Tidak ada data dengan ID: " + id);
    return 0;
  }

  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    var rowIndex = rowsToDelete[i];
    console.log("Menghapus baris index: " + rowIndex);
    table.rows.getItemAt(rowIndex).delete();
  }

  await context.sync();
  return rowsToDelete.length;
}

function setupKeyboardNavigation() {
  var inputs = document.querySelectorAll("input:not([type='hidden']), select, textarea");
  inputs.forEach(function (input) {
    input.removeEventListener("keydown", handleKeyNavigation);
    input.addEventListener("keydown", handleKeyNavigation);
  });
}

function handleKeyNavigation(e) {
  var target = e.target;

  if (e.key === "Enter" || e.key === "ArrowDown") {
    var success = attemptVerticalNavigation(target, 1);

    if (!success) {
      focusNextInput(target);
    }
    e.preventDefault();
  } else if (e.key === "ArrowUp") {
    var success = attemptVerticalNavigation(target, -1);

    if (!success) {
      focusPreviousInput(target);
    }
    e.preventDefault();
  }
}

function attemptVerticalNavigation(currentElement, direction) {
  if (!currentElement.id) return false;

  var match = currentElement.id.match(/^([a-zA-Z_]+)(\d+)$/);

  if (match) {
    var prefix = match[1];
    var currentNum = parseInt(match[2]);
    var nextNum = currentNum + direction;

    if (nextNum < 1) return false;

    var nextId = prefix + nextNum;
    var nextEl = document.getElementById(nextId);

    if (nextEl) {
      if (typeof nextEl["focus"] === "function") {
        nextEl["focus"]();
      }

      if (nextEl.tagName === "INPUT" || nextEl.tagName === "TEXTAREA") {
        setTimeout(function () {
          if (typeof nextEl["select"] === "function") {
            nextEl["select"]();
          }
        }, 10);
      }
      return true;
    }
  }
  return false;
}

function focusNextInput(currentElement) {
  var allNodeList = document.querySelectorAll(
    "input:not([type='hidden']):not([disabled]):not([readonly]), select:not([disabled]), textarea:not([disabled])",
  );
  var allInputs = [];

  for (var i = 0; i < allNodeList.length; i++) {
    allInputs.push(allNodeList[i]);
  }

  allInputs = allInputs.filter(function (el) {
    return el.getBoundingClientRect().width > 0 || el.getBoundingClientRect().height > 0;
  });

  var currentIndex = allInputs.indexOf(currentElement);
  var nextIndex = currentIndex + 1;

  if (nextIndex < allInputs.length) {
    var nextEl = allInputs[nextIndex];

    if (typeof nextEl["focus"] === "function") {
      nextEl["focus"]();
    }

    if (nextEl.tagName === "INPUT" || nextEl.tagName === "TEXTAREA") {
      setTimeout(function () {
        if (typeof nextEl["select"] === "function") {
          nextEl["select"]();
        }
      }, 10);
    }
  }
}

function focusPreviousInput(currentElement) {
  var allNodeList = document.querySelectorAll(
    "input:not([type='hidden']):not([disabled]):not([readonly]), select:not([disabled]), textarea:not([disabled])",
  );
  var allInputs = [];

  for (var i = 0; i < allNodeList.length; i++) {
    allInputs.push(allNodeList[i]);
  }

  allInputs = allInputs.filter(function (el) {
    return el.getBoundingClientRect().width > 0 || el.getBoundingClientRect().height > 0;
  });

  var currentIndex = allInputs.indexOf(currentElement);
  var prevIndex = currentIndex - 1;

  if (prevIndex >= 0) {
    var prevEl = allInputs[prevIndex];

    if (typeof prevEl["focus"] === "function") {
      prevEl["focus"]();
    }

    if (prevEl.tagName === "INPUT" || prevEl.tagName === "TEXTAREA") {
      setTimeout(function () {
        if (typeof prevEl["select"] === "function") {
          prevEl["select"]();
        }
      }, 10);
    }
  }
}

function setupTimeInputAutoFormat() {
  var timeInputIds = ["StartProduction", "EndProduction"];

  timeInputIds.forEach(function (id) {
    var el = document.getElementById(id);

    if (el) {
      el["type"] = "text";
      el["maxLength"] = 5;
      el["placeholder"] = "HH:MM";
      el.setAttribute("data-time-input", "true");

      el.addEventListener("input", handleTimeInput);
      el.addEventListener("keydown", handleTimeKeydown);
      el.addEventListener("blur", handleTimeBlur);
      el.addEventListener("paste", handleTimePaste);
    }
  });

  setupDynamicTimeInputs();
}

function setupDynamicTimeInputs() {
  for (var i = 1; i <= MAX_DOWNTIME_ROWS; i++) {
    var el = document.getElementById("JamStart" + i);

    if (el && !el.getAttribute("data-time-input")) {
      el["type"] = "text";
      el["maxLength"] = 5;
      el["placeholder"] = "HH:MM";
      el.setAttribute("data-time-input", "true");

      el.addEventListener("input", handleTimeInput);
      el.addEventListener("keydown", handleTimeKeydown);
      el.addEventListener("blur", handleTimeBlur);
      el.addEventListener("paste", handleTimePaste);
    }
  }
}

function handleTimeInput(e) {
  var input = e.target;
  var value = input.value;

  var numbers = value.replace(/\D/g, "");

  if (numbers.length > 4) {
    numbers = numbers.substring(0, 4);
  }

  var formatted = "";

  if (numbers.length > 0) {
    var hours = numbers.substring(0, 2);

    if (parseInt(hours) > 23) {
      hours = "23";
    }

    formatted = hours;

    if (numbers.length > 2) {
      var minutes = numbers.substring(2, 4);

      if (parseInt(minutes) > 59) {
        minutes = "59";
      }

      formatted += ":" + minutes;
    }
  }

  input.value = formatted;

  if (formatted.length === 5) {
    var event = new Event("change", { bubbles: true });
    input.dispatchEvent(event);
  }
}

function handleTimeKeydown(e) {
  var input = e.target;
  var key = e.key;

  if (
    ["Backspace", "Delete", "Tab", "Escape", "Enter", "ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"].indexOf(
      key,
    ) !== -1
  ) {
    return;
  }

  if ((e.ctrlKey || e.metaKey) && ["a", "c", "v", "x"].indexOf(key.toLowerCase()) !== -1) {
    return;
  }

  if (key < "0" || key > "9") {
    e.preventDefault();
    return;
  }

  var value = input.value.replace(/\D/g, "");
  if (value.length >= 4) {
    e.preventDefault();
  }
}

function handleTimeBlur(e) {
  var input = e.target;
  var value = input.value;

  if (!value) return;

  var numbers = value.replace(/\D/g, "");

  if (numbers.length === 1) {
    numbers = "0" + numbers + "00";
  } else if (numbers.length === 2) {
    numbers = numbers + "00";
  } else if (numbers.length === 3) {
    numbers = "0" + numbers;
  }

  if (numbers.length >= 4) {
    var hours = numbers.substring(0, 2);
    var minutes = numbers.substring(2, 4);

    if (parseInt(hours) > 23) hours = "23";
    if (parseInt(minutes) > 59) minutes = "59";

    input.value = hours + ":" + minutes;

    var event = new Event("change", { bubbles: true });
    input.dispatchEvent(event);
  }
}

function handleTimePaste(e) {
  e.preventDefault();

  var input = e.target;

  var clipboardData = e["clipboardData"] || window["clipboardData"];
  var pastedText = clipboardData ? clipboardData.getData("text") : "";

  var numbers = pastedText.replace(/\D/g, "");

  if (numbers.length >= 3) {
    var hours = numbers.substring(0, 2);
    var minutes = numbers.substring(2, 4);

    if (parseInt(hours) > 23) hours = "23";
    if (parseInt(minutes) > 59) minutes = "59";

    input["value"] = hours + ":" + minutes;

    var event = new Event("change", { bubbles: true });
    input.dispatchEvent(event);
  } else {
    input["value"] = numbers;
  }
}

function isValidTimeFormat(timeStr) {
  if (!timeStr) return false;

  var pattern = /^([0-1][0-9]|2[0-3]):([0-5][0-9])$/;
  return pattern.test(timeStr);
}

function normalizeTimeInput(input) {
  if (!input) return "";

  var numbers = String(input).replace(/\D/g, "");

  if (numbers.length === 0) return "";
  if (numbers.length === 1) numbers = "0" + numbers + "00";
  if (numbers.length === 2) numbers = numbers + "00";
  if (numbers.length === 3) numbers = "0" + numbers;

  var hours = numbers.substring(0, 2);
  var minutes = numbers.substring(2, 4);

  if (parseInt(hours) > 23) hours = "23";
  if (parseInt(minutes) > 59) minutes = "59";

  return hours + ":" + minutes;
}