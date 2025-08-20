let tabCount = 0;
let currentTab = 0;
const formTemplates = [];
const formData = [];

const siteId = "your-site-id";
const listId = "your-list-id";
const driveId = "your-drive-id";
const folderName = "PART145Uploads";

// Replace with MSAL or other auth method
async function getAccessToken() {
  return "your-access-token";
}

function isFormValid(data) {
  const requiredIndexes = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 16, 17, 18];
  return requiredIndexes.every(i => data[i] && data[i].trim() !== "");
}

async function loadFormTemplate() {
  const response = await fetch('fill_form.html');
  return await response.text();
}

async function switchTab(index) {
  const tabs = document.querySelectorAll('.tab');
  tabs.forEach((tab, i) => {
    tab.classList.toggle('active', i === index);
  });

  const formHTML = formTemplates[index].replace('{{index}}', index + 1);
  document.getElementById('tab-content').innerHTML = formHTML;
  currentTab = index;

  const inputs = document.querySelectorAll('.form-instance input, .form-instance textarea');
  inputs.forEach((input, i) => {
    input.value = formData[index][i] || '';
    input.oninput = () => {
      formData[index][i] = input.value;
      localStorage.setItem('formData', JSON.stringify(formData));
    };
  });
}

async function addTab() {
  const tabsContainer = document.getElementById('tabs');
  tabCount++;
  const tabIndex = tabCount - 1;

  const tabWrapper = document.createElement('div');
  tabWrapper.className = 'tab-wrapper';

  const tabLabel = `Form ${tabCount}`;

  const newTab = document.createElement('button');
  newTab.className = 'tab';
  newTab.textContent = tabLabel;
  newTab.onclick = () => switchTab(tabIndex);

  const delBtn = document.createElement('span');
  delBtn.className = 'del-icon';
  delBtn.textContent = '✖';
  delBtn.onclick = (e) => {
    e.stopPropagation();
    tabsContainer.removeChild(tabWrapper);
    formTemplates.splice(tabIndex, 1);
    formData.splice(tabIndex, 1);
    tabCount--;

    localStorage.setItem('formData', JSON.stringify(formData));

    const remainingTabs = document.querySelectorAll('.tab');
    if (remainingTabs.length > 0) {
      switchTab(0);
    } else {
      document.getElementById('tab-content').innerHTML = '';
    }
  };

  tabWrapper.appendChild(newTab);
  tabWrapper.appendChild(delBtn);
  tabsContainer.appendChild(tabWrapper);

  const formHTML = await loadFormTemplate();
  formTemplates.push(formHTML);
  formData.push([]);
  switchTab(tabIndex);
}

window.onload = async () => {
  const savedData = localStorage.getItem('formData');
  if (savedData) {
    const parsedData = JSON.parse(savedData);
    for (let i = 0; i < parsedData.length; i++) {
      await addTab();
    }
    formData.splice(0, formData.length, ...parsedData);
    switchTab(0);
  } else {
    addTab();
  }
};

function clearCurrentForm() {
  const confirmed = confirm("Are you sure you want to clear all fields in this form?");
  if (!confirmed) return;

  const inputs = document.querySelectorAll('.form-instance input, .form-instance textarea');
  inputs.forEach((input, i) => {
    input.value = '';
    formData[currentTab][i] = '';
  });

  localStorage.setItem('formData', JSON.stringify(formData));
}

async function submitToSharePoint(dataArray, headers) {
  const token = await getAccessToken();
  for (const data of dataArray) {
    const fields = {};
    headers.forEach((header, i) => {
      fields[header.replace(/[^a-zA-Z0-9]/g, "")] = data[i];
    });

    await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ fields })
    });
  }
}

async function uploadExcelToFolder(fileBlob, filename) {
  const token = await getAccessToken();
  await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${folderName}/${filename}:/content`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: fileBlob
  });
}

document.getElementById('submitDownloadButton').addEventListener('click', async () => {
  const validate = document.getElementById('validateBeforeExport').checked;

  if (validate) {
    const invalidForms = formData.filter(data => !isFormValid(data));
    if (invalidForms.length > 0) {
      alert("Some forms are missing required fields. Please complete them before submitting.");
      return;
    }
  }

  const headers = [
    "Viasat Tech", "Date", "Station", "Customer", "Type of Dispatch",
    "A/C Tail Number/Nose ID", "A/C Type", "Airline Log Page", "WO Number",
    "Time On", "Time Off", "MEL/NEF/NA",
    "PN OFF", "SN OFF", "PN ON", "SN ON",
    "Discrepancy Reported", "Work Performed"
  ];

  const rows = formData.map(data => headers.map((_, i) => data[i] || ""));
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Report");

  const firstForm = formData[0] || [];
  const station = firstForm[2] || "Station";
  const date = firstForm[1] || "Date";
  const airline = firstForm[3] || "Airline";
  const safeName = `${station}_Turnover_${date}_${airline}`.replace(/[^a-zA-Z0-9-_]/g, "_");
  const filename = `${safeName}.xlsx`;

  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const fileBlob = new Blob([wbout], { type: "application/octet-stream" });

  await submitToSharePoint(rows, headers);
  await uploadExcelToFolder(fileBlob, filename);

  XLSX.writeFile(workbook, filename);

  formTemplates.length = 0;
  formData.length = 0;
  tabCount = 0;
  currentTab = 0;
  document.getElementById('tabs').innerHTML = '';
  document.getElementById('tab-content').innerHTML = '';
  localStorage.removeItem('formData');
  await addTab();
});

document.getElementById('excelUpload').addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async function (e) {
    const workbook = XLSX.read(e.target.result, { type: "binary" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataArray = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers = dataArray[0];
    const rows = dataArray.slice(1);

    formTemplates.length = 0;
    formData.length = 0;
    tabCount = 0;
    currentTab = 0;
    document.getElementById('tabs').innerHTML = '';
    document.getElementById('tab-content').innerHTML = '';

    for (const row of rows) {
      await addTab();
      formData[formData.length - 1] = row;
    }

    localStorage.setItem('formData', JSON.stringify(formData));
  };
  reader.readAsBinaryString(file);
});

const msalConfig = {
  auth: {
    clientId: "YOUR-CLIENT-ID", // from Azure portal
    authority: "https://login.microsoftonline.com/YOUR-TENANT-ID",
    redirectUri: window.location.href
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function getAccessToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    await msalInstance.loginPopup({ scopes: ["Sites.ReadWrite.All", "Files.ReadWrite.All"] });
  }

  const result = await msalInstance.acquireTokenSilent({
    scopes: ["Sites.ReadWrite.All", "Files.ReadWrite.All"],
    account: msalInstance.getAllAccounts()[0]
  });

  return result.accessToken;
}
