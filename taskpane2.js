
/* global Office, fetch, console, document */
let templatesData = [];

const msalConfig = {
  auth: {
    clientId: "CLIENT ID", // 🔁 Insert your Azure App (Client) ID
    authority: "https://login.microsoftonline.com/TENAND_ID", // 🔁 Insert your Tenant ID
    redirectUri: "https://harpal804.github.io/mail-template-pub/taskpane.html" // 🔁 Match with your manifest
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const loginRequest = {
  scopes: ["https://graph.microsoft.com/Sites.Read.All"]
};

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insertTemplateBtn").onclick = insertTemplate;
    loadTemplates(); // Load SharePoint list templates
  }
});

async function loadTemplates() {
  try {
    const accounts = msalInstance.getAllAccounts();
    let account = accounts[0];

    if (!account) {
      const loginResponse = await msalInstance.loginPopup(loginRequest);
      account = loginResponse.account;
    }

    msalInstance.setActiveAccount(account);

    let tokenResponse;
    try {
      tokenResponse = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    } catch (silentError) {
      tokenResponse = await msalInstance.acquireTokenPopup(loginRequest);
    }

    const accessToken = tokenResponse.accessToken;

    const siteUrl = "https://COMPANY.sharepoint.com/sites/OperationWiki";
    const listName = "EmailTemplates";
    const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

    const response = await fetch(apiUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json;odata=verbose"
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    templatesData = data.d.results.map(item => ({
      category: item.Category,
      title: item.Title,
      body: item.Body
    }));

    console.log("✅ Templates fetched from SharePoint:", templatesData);
    populateCategoryDropdown();
  } catch (error) {
    console.error("❌ Error fetching templates from SharePoint:", error);
    showNotification("Error", "Failed to load templates from SharePoint.");
  }
}

function populateCategoryDropdown() {
  const categorySelect = document.getElementById("categorySelect");
  const categories = [...new Set(templatesData.map(t => t.category))];

  categorySelect.innerHTML = '<option value="">Select Category</option>';
  categories.forEach(category => {
    let option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    categorySelect.appendChild(option);
  });

  categorySelect.addEventListener("change", populateTitleListBox);
}

function populateTitleListBox() {
  const titleListBox = document.getElementById("titleListBox");
  const titleSearch = document.getElementById("titleSearch");
  const clearSearchBtn = document.getElementById("clearSearchBtn");
  const category = document.getElementById("categorySelect").value;

  if (!titleListBox || !titleSearch || !clearSearchBtn || !category) return;

  titleSearch.style.display = "block";
  clearSearchBtn.style.display = "block";
  titleSearch.value = "";

  let filteredTemplates = templatesData.filter(t => t.category === category);

  function updateTitleListBox(searchText = "") {
    titleListBox.innerHTML = "";
    filteredTemplates
      .filter(template => template.title.toLowerCase().startsWith(searchText.toLowerCase()))
      .forEach(template => {
        let option = document.createElement("option");
        option.value = template.body;
        option.textContent = template.title;
        titleListBox.appendChild(option);
      });
  }

  updateTitleListBox();

  titleSearch.addEventListener("input", () => {
    updateTitleListBox(titleSearch.value);
  });

  clearSearchBtn.addEventListener("click", () => {
    titleSearch.value = "";
    updateTitleListBox();
  });

  titleListBox.addEventListener("change", () => {
    document.getElementById("insertTemplateBtn").removeAttribute("disabled");
  });
}

function insertTemplate() {
  const categorySelect = document.getElementById("categorySelect");
  const titleListBox = document.getElementById("titleListBox");
  const titleSearch = document.getElementById("titleSearch");
  const clearSearchBtn = document.getElementById("clearSearchBtn");
  const insertButton = document.getElementById("insertTemplateBtn");

  if (!titleListBox || !titleListBox.value) {
    showNotification("⚠️ Error", "Please select a template before inserting it into the email.");
    return;
  }

  const templateText = titleListBox.value;

  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    showNotification("❌ Error", "Office Add-in API not available. Please restart Outlook.");
    return;
  }

  Office.context.mailbox.item.body.setSelectedDataAsync(
    templateText,
    { coercionType: Office.CoercionType.Text },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("❌ Failed to insert template:", result.error.message);
        showNotification("❌ Error", "Unable to insert the template.");
      } else {
        console.log("✅ Template inserted successfully.");
        showNotification("✅ Success", "Template inserted into the email.");

        categorySelect.selectedIndex = 0;
        titleListBox.innerHTML = "";
        titleSearch.value = "";
        titleSearch.style.display = "none";
        clearSearchBtn.style.display = "none";
        insertButton.setAttribute("disabled", "true");
      }
    }
  );
}

function showNotification(title, message) {
  if (!Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    alert(title + ": " + message);
    return;
  }

  Office.context.mailbox.item.notificationMessages.addAsync("notification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "icon-16",
    persistent: false
  });
}
