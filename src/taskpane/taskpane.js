/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT license.
 */

/* global document, Office, fetch */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("insertTemplateBtn").onclick = insertTemplate;
    loadTemplates();
  }
});

let templatesData = [];

async function loadTemplates() {
  try {
    const response = await fetch("/templates.json");
    const data = await response.json();
    templatesData = data.templates;

    populateCategoryDropdown();
  } catch (error) {
    console.error("Error loading templates:", error);
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
  const category = (document.getElementById("categorySelect")).value;

  if (!titleListBox || !titleSearch || !clearSearchBtn || !category) {
    console.error("Title list box, search input, or clear button not found.");
    return;
  }

  // Show the search input and clear button when category is selected
  titleSearch.style.display = "block";
  clearSearchBtn.style.display = "block";
  titleSearch.value = ""; // Reset search when category changes

  // Get the filtered titles based on category
  let filteredTemplates = templatesData.filter(t => t.category === category);

  // Function to update the title list box based on search text
  function updateTitleListBox(searchText = "") {
    titleListBox.innerHTML = ''; // Clear existing items

    filteredTemplates
      .filter(template => template.title.toLowerCase().startsWith(searchText.toLowerCase()))
      .forEach(template => {
        let option = document.createElement("option");
        option.value = template.body; // Store template body in value
        option.textContent = template.title;
        titleListBox.appendChild(option);
      });
  }

  // Initial population of the list box
  updateTitleListBox();  
  
  // Listen for user input in search box
  titleSearch.addEventListener("input", () => {
    updateTitleListBox(titleSearch.value);
  });

  // Clear button functionality
  clearSearchBtn.addEventListener("click", () => {
    titleSearch.value = "";
    updateTitleListBox();
  });
  
  // Ensure selecting an item enables the insert button
  titleListBox.addEventListener("change", () => {
    const insertButton = document.getElementById("insertTemplateBtn");
    if (insertButton) {
      insertButton.removeAttribute("disabled");
    }
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

  const templateText = titleListBox.value; // Get selected template body

  console.log("✅ Inserting template:", templateText);

  // Ensure Office API is initialized before inserting text
  if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
    showNotification("❌ Error", "Office Add-in API not available. Please restart Outlook.");
    console.error("Office.context.mailbox.item is undefined.");
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
      
      // ✅ Reset selections after insertion
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

/**
 * Show an Office notification message (instead of alert()).
 */
function showNotification(title, message) {
  Office.context.mailbox.item.notificationMessages.addAsync("notification", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: "icon-16", // Optional
    persistent: false
  });
}
