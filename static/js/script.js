const dom = {};

// Initialize theme based on saved preference or system default
function initTheme() {
  const toggleSwitch = document.querySelector("#checkbox");
  dom.toggleSwitch = toggleSwitch;
  const currentTheme =
    localStorage.getItem("theme") ||
    (window.matchMedia?.("(prefers-color-scheme: dark)").matches
      ? "dark"
      : "light");

  document.documentElement.setAttribute("data-theme", currentTheme);
  toggleSwitch.checked = currentTheme === "dark";

  if (
    window.matchMedia?.("(prefers-color-scheme: dark)").matches &&
    !localStorage.getItem("theme")
  ) {
    localStorage.setItem("theme", "dark");
  }

  toggleSwitch.addEventListener("change", (e) => {
    const theme = e.target.checked ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("theme", theme);
  });
}

// Function to match spans based on data-diff-id attribute - REVISED
function matchSpans(addedSpans, removedSpans, isTemp = false) {
  const hoverClass = "bright-highlight";
  const prefix = isTemp ? "_temp" : "_"; // Prefix for storing event listeners to allow removal

  // Clear existing event listeners first
  const clearListeners = (spans, prefix) => {
    spans.forEach((span) => {
      if (span[`${prefix}HandleMouseOver`]) {
        span.removeEventListener("mouseover", span[`${prefix}HandleMouseOver`]);
        span[`${prefix}HandleMouseOver`] = null; // Clear stored function
      }
      if (span[`${prefix}HandleMouseOut`]) {
        span.removeEventListener("mouseout", span[`${prefix}HandleMouseOut`]);
        span[`${prefix}HandleMouseOut`] = null; // Clear stored function
      }
    });
  };

  clearListeners(addedSpans, prefix);
  clearListeners(removedSpans, prefix);

  const removedSpansMap = new Map();
  removedSpans.forEach((span) => {
    const diffId = span.getAttribute("data-diff-id");
    if (diffId) {
      removedSpansMap.set(diffId, span);
    }
  });

  addedSpans.forEach((addedSpan) => {
    const diffId = addedSpan.getAttribute("data-diff-id");
    if (diffId && removedSpansMap.has(diffId)) {
      const matchedRemovedSpan = removedSpansMap.get(diffId);

      // Store listeners on the elements to manage them
      const handleAddedOver = () =>
        matchedRemovedSpan.classList.add(hoverClass);
      const handleAddedOut = () =>
        matchedRemovedSpan.classList.remove(hoverClass);
      addedSpan[`${prefix}HandleMouseOver`] = handleAddedOver;
      addedSpan[`${prefix}HandleMouseOut`] = handleAddedOut;
      addedSpan.addEventListener("mouseover", handleAddedOver);
      addedSpan.addEventListener("mouseout", handleAddedOut);

      const handleRemovedOver = () => addedSpan.classList.add(hoverClass);
      const handleRemovedOut = () => addedSpan.classList.remove(hoverClass);
      matchedRemovedSpan[`${prefix}HandleMouseOver`] = handleRemovedOver;
      matchedRemovedSpan[`${prefix}HandleMouseOut`] = handleRemovedOut;
      matchedRemovedSpan.addEventListener("mouseover", handleRemovedOver);
      matchedRemovedSpan.addEventListener("mouseout", handleRemovedOut);

      // Optional: remove from map if IDs are truly unique per pair
      // removedSpansMap.delete(diffId);
    }
  });
}

// Helper function for diff highlighting - REVISED (calls the modified matchSpans)
function setupDiffHighlighting(row) {
  if (!row) return;
  const colA = row.querySelector("td:nth-child(2)");
  const colB = row.querySelector("td:nth-child(3)");

  if (!colA || !colB) return;

  const colAContent = colA.querySelector(".cell-content");
  const colBContent = colB.querySelector(".cell-content");

  if (!colAContent || !colBContent) return;

  // Ensure we select spans directly within the content divs
  const addedSpans = Array.from(
    colBContent.querySelectorAll(".added[data-diff-id]")
  );
  const removedSpans = Array.from(
    colAContent.querySelectorAll(".removed[data-diff-id]")
  );

  matchSpans(addedSpans, removedSpans);
}

// --- Scroll Syncing Logic ---
let isSyncingScroll = false; // Flag to prevent infinite loops

// The handler function that synchronizes scroll positions
function handleSyncedScroll(event, elementsToSync) {
  if (isSyncingScroll) return; // Prevent recursion
  isSyncingScroll = true;

  const scrollTop = event.target.scrollTop;
  const scrollLeft = event.target.scrollLeft; // Also sync horizontal scroll if needed

  elementsToSync.forEach((el) => {
    // Only sync if the element is different from the source and exists
    if (el && el !== event.target) {
      el.scrollTop = scrollTop;
      el.scrollLeft = scrollLeft;
    }
  });

  // Use requestAnimationFrame to reset the flag after the browser has painted
  requestAnimationFrame(() => {
    isSyncingScroll = false;
  });
}

// Function to add scroll listeners
function addScrollSyncListeners(elements) {
  elements.forEach((el) => {
    if (el) {
      // Store the handler reference on the element for easy removal
      el._scrollSyncHandler = (e) => handleSyncedScroll(e, elements);
      el.addEventListener("scroll", el._scrollSyncHandler);
    }
  });
}

// Function to remove scroll listeners
function removeScrollSyncListeners(elements) {
  elements.forEach((el) => {
    if (el && el._scrollSyncHandler) {
      el.removeEventListener("scroll", el._scrollSyncHandler);
      delete el._scrollSyncHandler; // Clean up the stored reference
    }
  });
}
// --- End Scroll Syncing Logic ---

// Modify makeEditable function
function makeEditable(element) {
  const container = element.parentNode;
  const textArea = container.querySelector(".editable");
  const saveBtn = container.querySelector(".save-btn");
  const actionBtns = container.querySelector(".action-buttons");
  const row = element.closest("tr");
  const colATd = row.querySelector("td:nth-child(2)"); // Get the TD for Col A
  const colAContent = colATd ? colATd.querySelector(".cell-content") : null; // Find content within Col A TD

  // Add scrolling class to Col A content if it exists
  if (colAContent) {
    colAContent.classList.add("scrolling-active");
  }

  textArea.setAttribute(
    "data-original-col-a",
    colAContent ? colAContent.textContent.trim() : ""
  );
  textArea.setAttribute(
    "data-original-col-a-html",
    colAContent ? colAContent.innerHTML : ""
  );

  // UI updates
  element.style.display = "none";
  textArea.style.display = "block";
  saveBtn.style.display = "inline-block";

  Object.assign(actionBtns.style, {
    zIndex: "10",
    bottom: "0px",
    opacity: "1",
  });

  // Scroll into view if needed
  const rect = textArea.getBoundingClientRect();
  const viewHeight = window.innerHeight;
  const top = rect.top + window.scrollY;
  const bottom = rect.bottom + window.scrollY;

  if (bottom > window.scrollY + viewHeight || top < window.scrollY) {
    window.scrollTo({
      top: top - viewHeight / 2 + rect.height / 2,
      behavior: "smooth",
    });
  }

  // Create/setup diff previews (ensure colAContainer is correctly identified)
  const previewB =
    container.querySelector(".diff-preview") ||
    (() => {
      const div = document.createElement("div");
      div.className = "diff-preview";
      container.insertBefore(div, textArea);
      return div;
    })();
  previewB.style.display = "block"; // Show preview B

  const colAContainer = colAContent ? colAContent.parentNode : null; // Parent of col A content

  const previewA = colAContainer
    ? colAContainer.querySelector(".col-a-diff-preview") ||
      (() => {
        const div = document.createElement("div");
        div.className = "col-a-diff-preview";
        div.style.display = "none"; // Start hidden
        colAContainer.insertBefore(div, colAContent.nextSibling);
        return div;
      })()
    : null; // Only create preview A if col A exists

  // --- Scroll Sync Setup ---
  const elementsToSync = [colAContent, textArea, previewB, previewA].filter(
    (el) => el
  );
  textArea._syncElements = elementsToSync;
  addScrollSyncListeners(elementsToSync);
  // --- End Scroll Sync Setup ---

  // Update diff preview function (keep existing fetch logic)
  const updatePreview = () => {
    const colAText = colAContent ? colAContent.textContent.trim() : ""; // Handle null colAContent
    const colBText = textArea.value.trimEnd();

    fetch("/preview_diff", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: `text1=${encodeURIComponent(colAText)}&text2=${encodeURIComponent(
        colBText
      )}`,
    })
      .then((response) => response.json())
      .then((data) => {
        if (data.status === "error") {
          console.error("Preview error:", data.message);
          previewB.innerHTML = `<div>${colBText.replace(/\n/g, "<br>")}</div>`;
          if (colAContent) colAContent.style.display = "block"; // Show original A if preview fails
          if (previewA) previewA.style.display = "none"; // Hide A preview
          if (previewA) previewA.innerHTML = ""; // Clear A preview
          return;
        }
        // Update previews
        previewB.innerHTML = `<div>${data.highlighted_b || ""}</div>`;
        if (previewA)
          previewA.innerHTML = `<div>${data.highlighted_a || ""}</div>`;

        if (colAContent) colAContent.style.display = "none"; // Hide original A content
        if (previewA) previewA.style.display = "block"; // Show A diff preview

        // Setup temp highlighting for the preview elements
        const tempAddedSpans = Array.from(
          previewB.querySelectorAll(".added[data-diff-id]")
        );
        const tempRemovedSpans = previewA
          ? Array.from(previewA.querySelectorAll(".removed[data-diff-id]"))
          : [];
        matchSpans(tempAddedSpans, tempRemovedSpans, true);
      })
      .catch((error) => {
        console.error("Error fetching diff preview:", error);
        previewB.innerHTML = `<div>${colBText.replace(/\n/g, "<br>")}</div>`;
        if (colAContent) colAContent.style.display = "block";
        if (previewA) previewA.style.display = "none";
        if (previewA) previewA.innerHTML = "";
      });
  };

  updatePreview(); // Initial call

  // Event listeners (keep existing escape/ctrl+enter)
  textArea.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      cancelEditing(element);
    } else if (e.key === "Enter" && e.ctrlKey) {
      saveCell(saveBtn);
    }
  });
  textArea.addEventListener("input", function () {
    updatePreview();
  }); // Removed height adjustment

  // Focus the textarea after UI has updated and preview is generated
  setTimeout(() => {
    textArea.focus();
    textArea.selectionStart = textArea.selectionEnd = textArea.value.length;
  }, 100);

  // Handle clicks outside editing area
  const handleOutsideClick = (e) => {
    if (!container.contains(e.target)) {
      cancelEditing(element);
      // No need to remove the listener here due to { once: true }
    }
  };
  setTimeout(
    () =>
      document.addEventListener("click", handleOutsideClick, {
        capture: true,
        once: true,
      }),
    300
  ); // Use capture and once for reliability
}

// Modify cancelEditing
function cancelEditing(contentElement) {
  const container = contentElement.parentNode;
  const textArea = container.querySelector(".editable");
  const saveBtn = container.querySelector(".save-btn");
  const previewDiv = container.querySelector(".diff-preview");
  const actionBtns = container.querySelector(".action-buttons");

  // Get Column A elements
  const row = container.closest("tr");
  const colATd = row.querySelector("td:nth-child(2)");
  const colAContent = colATd ? colATd.querySelector(".cell-content") : null;
  const colAPreview = colATd
    ? colATd.querySelector(".col-a-diff-preview")
    : null;

  // --- Remove Scroll Sync Listeners ---
  if (textArea._syncElements) {
    removeScrollSyncListeners(textArea._syncElements);
    delete textArea._syncElements; // Clean up
  }
  // --- End Remove Scroll Sync Listeners ---

  // Remove scrolling class from Col A content if it exists
  if (colAContent) {
    colAContent.classList.remove("scrolling-active");
  }

  // Restore displays
  contentElement.style.display = "block";
  textArea.style.display = "none";
  saveBtn.style.display = "none";
  if (previewDiv) previewDiv.style.display = "none";
  if (colAContent) colAContent.style.display = "block"; // Restore original Col A display
  if (colAPreview) colAPreview.style.display = "none"; // Hide Col A preview

  // Restore action button styles
  if (actionBtns) {
    Object.assign(actionBtns.style, {
      zIndex: "", // Revert to default or CSS defined
      bottom: "", // Revert to default or CSS defined
      opacity: "", // Revert to default or CSS defined
    });
  }

  // Remove the temporary event listener for outside clicks if it hasn't fired yet
  // Note: This part is tricky; using {once: true} in makeEditable is generally better.
  // If not using {once: true}, you'd need a reference to handleOutsideClick stored somewhere.
}

// Modify saveCell
function saveCell(saveBtn) {
  const container = saveBtn.closest(".cell-container");
  const textArea = container.querySelector(".editable");
  const contentDiv = container.querySelector(".cell-content");
  const previewDiv = container.querySelector(".diff-preview");
  const rowIdx = textArea.getAttribute("data-row");
  const page = textArea.getAttribute("data-page"); // Ensure these are set if needed by backend/logic
  const rowsPerPage = textArea.getAttribute("data-rows"); // Ensure these are set if needed

  // Get Column A elements
  const row = container.closest("tr");
  const colATd = row.querySelector("td:nth-child(2)");
  const colAContent = colATd ? colATd.querySelector(".cell-content") : null;
  const colAPreview = colATd
    ? colATd.querySelector(".col-a-diff-preview")
    : null;

  // Remove Scroll Sync Listeners if they exist
  if (textArea._syncElements) {
    removeScrollSyncListeners(textArea._syncElements);
    delete textArea._syncElements; // Clean up
  }

  // Remove scrolling class from Col A content if it exists
  if (colAContent) {
    colAContent.classList.remove("scrolling-active");
  }

  const text = textArea.value.replace(/\\r\\n/g, "\\n").replace(/\\r/g, "\\n");
  const formData = new FormData();
  formData.append("row_idx", rowIdx);
  formData.append("text", text);
  // Add page/rowsPerPage if your backend uses them during save
  // formData.append('page', page);
  // formData.append('rows_per_page', rowsPerPage);

  console.log(
    "Saving cell - Row:",
    rowIdx,
    "Text:",
    text.substring(0, 50) + "..."
  ); // Debug log

  fetch("/edit", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      console.log("Save response status:", response.status); // Debug log
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then((data) => {
      console.log("Save response data:", data); // Debug log
      if (data.status === "success") {
        console.log("Save successful, updating UI..."); // Debug log

        // Update Col B content
        contentDiv.innerHTML =
          data.highlighted_html || text.replace(/\\n/g, "<br>"); // Fallback if html missing

        // Update Col A content if provided and element exists
        if (data.highlighted_a_html && colAContent) {
          colAContent.innerHTML = data.highlighted_a_html;
        }

        // Update cell status class
        const cell = contentDiv.closest("td");
        if (cell) {
          // Check if cell exists
          cell.classList.remove("same", "different");
          cell.classList.add(data.diff_status || "different"); // Add fallback
        }

        // Restore UI states
        contentDiv.style.display = "block";
        textArea.style.display = "none";
        saveBtn.style.display = "none";
        if (previewDiv) previewDiv.style.display = "none";
        if (colAContent) colAContent.style.display = "block"; // Ensure Col A content is visible
        if (colAPreview) colAPreview.style.display = "none"; // Hide Col A preview

        // Show notification
        showNotification("Changes saved directly to Excel!"); // Ensure this function works

        // Re-setup highlighting
        setTimeout(() => {
          if (row) setupDiffHighlighting(row); // Check if row exists
        }, 0);
      } else {
        console.error("Save failed:", data.message); // Debug log
        showNotification("Error saving: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error saving cell:", error); // Debug log
      showNotification("Error saving cell: " + error.message, "error");
      // Optionally, restore editing state or inform user more clearly
      // For example, don't hide the textarea if save failed
    });
}

// Notification system - simplified
function showNotification(message, type = "success") {
  if (!dom.notification) {
    dom.notification = document.getElementById("notification");
  }

  const backgroundColor =
    {
      error: "#f44336",
      info: "#2196F3",
      success: "#4CAF50",
    }[type] || "#4CAF50";

  dom.notification.textContent = message;
  dom.notification.style.backgroundColor = backgroundColor;
  dom.notification.classList.add("show");

  setTimeout(() => dom.notification.classList.remove("show"), 3000);
}

// Handle form submissions - optimized
function setupFormHandlers() {
  // Factory for form submission handling
  const createHandler = (action, successCb) =>
    function (e) {
      e.preventDefault();
      const formData = new FormData(this);
      const button = this.querySelector("button");
      button.classList.add("tick-animation");

      fetch(action, {
        method: "POST",
        body: formData,
      })
        .then((response) => response.json())
        .then((data) => {
          if (data.status === "success") {
            successCb(button, formData, data);
          } else {
            showNotification("Error: " + data.message, "error");
          }
        })
        .catch((error) => {
          console.error("Error:", error);
          showNotification(`Error with ${action}`, "error");
        })
        .finally(() => {
          setTimeout(() => button.classList.remove("tick-animation"), 500);
        });
    };

  // Approval forms
  document.querySelectorAll('form[action="/approve"]').forEach((form) => {
    form.addEventListener(
      "submit",
      createHandler("/approve", (button, formData) => {
        const cell = button.closest("td");
        const type = formData.get("approval_type");
        const classes = {
          green: "approved",
          yellow: "yellow-approved",
          red: "red-approved",
        };

        cell.classList.remove("approved", "yellow-approved", "red-approved");
        cell.classList.add(classes[type]);
        showNotification(`Cell approved with ${type} tick!`);
      })
    );
  });

  // Reset cell forms
  document.querySelectorAll('form[action="/reset_cell"]').forEach((form) => {
    form.addEventListener(
      "submit",
      createHandler("/reset_cell", (button) => {
        button
          .closest("td")
          .classList.remove("approved", "yellow-approved", "red-approved");
        showNotification("Color Reset!");
      })
    );
  });
}

// Text selection functionality - optimized
function setupTextSelection() {
  let selectedText = "";
  const selBtn = document.getElementById("selectionButton");
  dom.selectionButton = selBtn;

  document.addEventListener("mouseup", (e) => {
    setTimeout(() => {
      const selection = window.getSelection();
      selectedText = selection.toString().trim();

      // Check if selection is inside a cell-content element
      let isInsideCellContent = false;

      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const container = range.commonAncestorContainer;

        // Check if the container or its parent is a cell-content element
        if (
          container.nodeType === 1 &&
          container.classList.contains("cell-content")
        ) {
          isInsideCellContent = true;
        } else if (
          container.parentElement &&
          container.parentElement.closest(".cell-content")
        ) {
          isInsideCellContent = true;
        }
      }

      if (
        selectedText &&
        !document.querySelector(".editable:focus") &&
        isInsideCellContent
      ) {
        const range = selection.getRangeAt(0);
        const rect = range.getBoundingClientRect();

        selBtn.style.display = "flex";
        selBtn.style.left = `${
          rect.left + rect.width / 2 - selBtn.offsetWidth / 2 + window.scrollX
        }px`;
        selBtn.style.top = `${rect.bottom + window.scrollY + 10}px`;

        // Keep on screen
        const btnRect = selBtn.getBoundingClientRect();
        if (btnRect.right > window.innerWidth) {
          selBtn.style.left = `${
            window.innerWidth - btnRect.width - 10 + window.scrollX
          }px`;
        }
        if (btnRect.left < 0) {
          selBtn.style.left = `${10 + window.scrollX}px`;
        }
      } else if (!e.target.closest("#selectionButton")) {
        selBtn.style.display = "none";
      }
    }, 10);
  });

  selBtn.addEventListener("click", () => {
    if (selectedText) {
      const formData = new FormData();
      formData.append("selected_text", selectedText);

      showNotification("Saving selection...", "info");

      fetch("/save_selection", {
        method: "POST",
        body: formData,
      })
        .then((response) => response.json())
        .then((data) => {
          showNotification(
            data.status === "success" ? "Saved!" : "Error: " + data.message,
            data.status === "success" ? "success" : "error"
          );
        })
        .catch((error) => {
          console.error("Error:", error);
          showNotification("Error saving selection", "error");
        });

      selBtn.style.display = "none";
    }
  });

  // Button styling events with arrow functions
  selBtn.addEventListener("mouseover", () => {
    selBtn.style.backgroundColor = "#45a049";
    selBtn.style.transform = "scale(1.05)";
  });

  selBtn.addEventListener("mouseout", () => {
    selBtn.style.backgroundColor = "#4CAF50";
    selBtn.style.transform = "scale(1)";
  });

  // Hide when clicking elsewhere
  document.addEventListener("click", (e) => {
    if (!e.target.closest("#selectionButton")) {
      selBtn.style.display = "none";
    }
  });
}

// Cell regeneration functionality - optimized
function regenerateCell(button) {
  const rowIdx = button.getAttribute("data-row");
  const page = button.getAttribute("data-page");
  const rowsPerPage = button.getAttribute("data-rows");

  const originalContent = button.innerHTML;
  button.innerHTML = '<span class="loading-spinner"></span>';
  button.disabled = true;

  showNotification("Regenerating...", "info");

  const container = button.closest(".cell-container");
  const contentDiv = container.querySelector(".cell-content");
  const textArea = container.querySelector(".editable");

  const formData = new FormData();
  formData.append("row_idx", rowIdx);
  formData.append("page", page);
  formData.append("rows_per_page", rowsPerPage);

  fetch("/regenerate_cell", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      return response.json();
    })
    .then((data) => {
      if (data.status === "success") {
        const row = container.closest("tr");
        const colAContent = row.querySelector("td:nth-child(2) .cell-content");
        const cell = contentDiv.closest("td");

        // 1. Update Col B content
        contentDiv.innerHTML = data.highlighted_html;
        textArea.value = data.new_text;

        // 2. Update Col A content (if provided)
        if (data.highlighted_a_html && colAContent) {
          colAContent.innerHTML = data.highlighted_a_html;
        }

        // 3. Update diff status
        cell.classList.remove("same", "different");
        cell.classList.add(data.diff_status);

        // 4. Update color approval status
        cell.classList.remove("approved", "yellow-approved", "red-approved");
        if (data.col_b_approved) {
          const classMap = {
            green: "approved",
            yellow: "yellow-approved",
            red: "red-approved",
          };
          const approvalClass = classMap[data.col_b_type] || "approved"; // Default to green if type is invalid
          cell.classList.add(approvalClass);
        }

        showNotification("Regenerated!");

        // 5. Re-setup highlighting
        setTimeout(() => {
          setupDiffHighlighting(row);
        }, 0);
      } else {
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error regenerating cell:", error);
      showNotification("Error regenerating cell: " + error.message, "error");
    })
    .finally(() => {
      button.innerHTML = originalContent;
      button.disabled = false;
    });
}

function regenerateWithPrompt1(button) {
  const rowIdx = button.getAttribute("data-row");
  const page = button.getAttribute("data-page");
  const rowsPerPage = button.getAttribute("data-rows");

  const originalContent = button.innerHTML;
  button.innerHTML = '<span class="loading-spinner"></span>';
  button.disabled = true;

  showNotification("Regenerating with prompt 1...", "info");

  const container = button.closest(".cell-container");
  const contentDiv = container.querySelector(".cell-content");
  const textArea = container.querySelector(".editable");

  const formData = new FormData();
  formData.append("row_idx", rowIdx);
  formData.append("page", page);
  formData.append("rows_per_page", rowsPerPage);

  fetch("/regenerate_with_prompt_1", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      return response.json();
    })
    .then((data) => {
      if (data.status === "success") {
        const row = container.closest("tr");
        const colAContent = row.querySelector("td:nth-child(2) .cell-content");
        const cell = contentDiv.closest("td");

        // 1. Update Col B content
        contentDiv.innerHTML = data.highlighted_html;
        textArea.value = data.new_text;

        // 2. Update Col A content (if provided)
        if (data.highlighted_a_html && colAContent) {
          colAContent.innerHTML = data.highlighted_a_html;
        }

        // 3. Update diff status
        cell.classList.remove("same", "different");
        cell.classList.add(data.diff_status);

        // 4. Update color approval status
        cell.classList.remove("approved", "yellow-approved", "red-approved");
        if (data.col_b_approved) {
          const classMap = {
            green: "approved",
            yellow: "yellow-approved",
            red: "red-approved",
          };
          const approvalClass = classMap[data.col_b_type] || "approved"; // Default to green if type is invalid
          cell.classList.add(approvalClass);
        }

        showNotification("Regenerated with prompt 1!");

        // 5. Re-setup highlighting
        setTimeout(() => {
          setupDiffHighlighting(row);
        }, 0);
      } else {
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error regenerating cell with prompt 1:", error);
      showNotification("Error regenerating cell: " + error.message, "error");
    })
    .finally(() => {
      button.innerHTML = originalContent;
      button.disabled = false;
    });
}

function regenerateWithPrompt2(button) {
  const rowIdx = button.getAttribute("data-row");
  const page = button.getAttribute("data-page");
  const rowsPerPage = button.getAttribute("data-rows");

  const originalContent = button.innerHTML;
  button.innerHTML = '<span class="loading-spinner"></span>';
  button.disabled = true;

  showNotification("Regenerating with prompt 2...", "info");

  const container = button.closest(".cell-container");
  const contentDiv = container.querySelector(".cell-content");
  const textArea = container.querySelector(".editable");

  const formData = new FormData();
  formData.append("row_idx", rowIdx);
  formData.append("page", page);
  formData.append("rows_per_page", rowsPerPage);

  fetch("/regenerate_with_prompt_2", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      return response.json();
    })
    .then((data) => {
      if (data.status === "success") {
        const row = container.closest("tr");
        const colAContent = row.querySelector("td:nth-child(2) .cell-content");
        const cell = contentDiv.closest("td");

        // 1. Update Col B content
        contentDiv.innerHTML = data.highlighted_html;
        textArea.value = data.new_text;

        // 2. Update Col A content (if provided)
        if (data.highlighted_a_html && colAContent) {
          colAContent.innerHTML = data.highlighted_a_html;
        }

        // 3. Update diff status
        cell.classList.remove("same", "different");
        cell.classList.add(data.diff_status);

        // 4. Update color approval status
        cell.classList.remove("approved", "yellow-approved", "red-approved");
        if (data.col_b_approved) {
          const classMap = {
            green: "approved",
            yellow: "yellow-approved",
            red: "red-approved",
          };
          const approvalClass = classMap[data.col_b_type] || "approved"; // Default to green if type is invalid
          cell.classList.add(approvalClass);
        }

        showNotification("Regenerated with prompt 2!");

        // 5. Re-setup highlighting
        setTimeout(() => {
          setupDiffHighlighting(row);
        }, 0);
      } else {
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error regenerating cell with prompt 2:", error);
      showNotification("Error regenerating cell: " + error.message, "error");
    })
    .finally(() => {
      button.innerHTML = originalContent;
      button.disabled = false;
    });
}

// Utility functions for managing body scroll
function disableBodyScroll() {
  const scrollY = window.scrollY;
  document.body.style.position = "fixed";
  document.body.style.top = `-${scrollY}px`;
  document.body.style.width = "100%";
  document.body.dataset.scrollY = scrollY;
}

function enableBodyScroll() {
  const scrollY = document.body.dataset.scrollY || 0;
  document.body.style.position = "";
  document.body.style.top = "";
  document.body.style.width = "";
  window.scrollTo(0, parseInt(scrollY || 0));
  delete document.body.dataset.scrollY;
}

// Modify the regenerateWithCustomPrompt function
function regenerateWithCustomPrompt(button) {
  const rowIdx = button.getAttribute("data-row");
  const page = button.getAttribute("data-page");
  const rowsPerPage = button.getAttribute("data-rows");

  // Get the modal elements
  const customPromptModal = document.getElementById("customPromptModal");
  const promptText = document.getElementById("customPromptText");
  const promptRowIdx = document.getElementById("promptRowIdx");
  const promptPage = document.getElementById("promptPage");
  const promptRowsPerPage = document.getElementById("promptRowsPerPage");

  // Set hidden fields
  promptRowIdx.value = rowIdx;
  promptPage.value = page;
  promptRowsPerPage.value = rowsPerPage;

  // Load saved prompt from localStorage if it exists
  const savedPrompt = localStorage.getItem("customPrompt");

  // Load default template prompt if nothing is saved or use the saved prompt
  if (savedPrompt) {
    promptText.value = savedPrompt;
  } else if (!promptText.value) {
    const defaultPrompt = `Please analyze this content and provide a Bengali translation:\n\nArabic Text: {arabic_text}\n\nPrevious Bengali Translation: {col_b_text}\n\nOriginal Bengali: {col_a_text}\n\nPlease provide an improved Bengali translation incorporating aspects from both the original and previous translation:`;
    promptText.value = defaultPrompt;
    localStorage.setItem("customPrompt", defaultPrompt);
  }

  // Display the modal and disable body scroll
  customPromptModal.style.display = "flex";
  disableBodyScroll();
}

function executeCustomPrompt() {
  const promptText = document.getElementById("customPromptText").value;
  const rowIdx = document.getElementById("promptRowIdx").value;
  const page = document.getElementById("promptPage").value;
  const rowsPerPage = document.getElementById("promptRowsPerPage").value;

  if (!promptText.trim()) {
    showNotification("Please enter a prompt first", "error");
    return;
  }

  // Save the prompt to localStorage
  localStorage.setItem("customPrompt", promptText);

  // Close the modal and re-enable body scroll
  document.getElementById("customPromptModal").style.display = "none";
  enableBodyScroll();

  // Find the custom prompt button for this row to show loading state
  const buttons = document.querySelectorAll(".custom-prompt-btn");
  let targetButton = null;

  for (const btn of buttons) {
    if (btn.getAttribute("data-row") === rowIdx) {
      targetButton = btn;
      break;
    }
  }

  if (!targetButton) {
    showNotification("Error finding the button", "error");
    return;
  }

  const originalContent = targetButton.innerHTML;
  targetButton.innerHTML = '<span class="loading-spinner"></span>';
  targetButton.disabled = true;

  showNotification("Generating with custom prompt...", "info");

  const formData = new FormData();
  formData.append("row_idx", rowIdx);
  formData.append("prompt", promptText);
  formData.append("page", page);
  formData.append("rows_per_page", rowsPerPage);

  fetch("/regenerate_with_custom_prompt", {
    method: "POST",
    body: formData,
  })
    .then((response) => {
      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      return response.json();
    })
    .then((data) => {
      if (data.status === "success") {
        const container = targetButton.closest(".cell-container");
        const contentDiv = container.querySelector(".cell-content");
        const textArea = container.querySelector(".editable");
        const row = container.closest("tr");
        const colAContent = row.querySelector("td:nth-child(2) .cell-content");
        const cell = contentDiv.closest("td");

        // 1. Update Col B content
        contentDiv.innerHTML = data.highlighted_html;
        textArea.value = data.new_text;

        // 2. Update Col A content (if provided)
        if (data.highlighted_a_html && colAContent) {
          colAContent.innerHTML = data.highlighted_a_html;
        }

        // 3. Update diff status
        cell.classList.remove("same", "different");
        cell.classList.add(data.diff_status);

        // 4. Update color approval status
        cell.classList.remove("approved", "yellow-approved", "red-approved");
        if (data.col_b_approved) {
          const classMap = {
            green: "approved",
            yellow: "yellow-approved",
            red: "red-approved",
          };
          const approvalClass = classMap[data.col_b_type] || "approved";
          cell.classList.add(approvalClass);
        }

        showNotification("Generated with custom prompt!");

        // 5. Re-setup highlighting
        setTimeout(() => {
          setupDiffHighlighting(row);
        }, 0);
      } else {
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error generating with custom prompt:", error);
      showNotification("Error: " + error.message, "error");
    })
    .finally(() => {
      targetButton.innerHTML = originalContent;
      targetButton.disabled = false;
    });
}

// Setup custom prompt placeholder insertion
function setupCustomPromptPlaceholders() {
  const placeholders = document.querySelectorAll(".placeholder-item");
  const promptTextarea = document.getElementById("customPromptText");

  placeholders.forEach((placeholder) => {
    placeholder.addEventListener("click", () => {
      // Get current cursor position
      const cursorPos = promptTextarea.selectionStart;
      const placeholderText = placeholder.textContent;

      // Insert placeholder at cursor position
      const textBefore = promptTextarea.value.substring(0, cursorPos);
      const textAfter = promptTextarea.value.substring(cursorPos);
      promptTextarea.value = textBefore + placeholderText + textAfter;

      // Move cursor after inserted placeholder
      promptTextarea.focus();
      promptTextarea.selectionStart = cursorPos + placeholderText.length;
      promptTextarea.selectionEnd = cursorPos + placeholderText.length;
    });
  });
}

// Function to regenerate all cells in parallel using backend endpoint
function regenerateAllCells() {
  // Get the selected color from the dropdown
  const colorSelect = document.getElementById("regenerate-color-select");
  const selectedColor = colorSelect ? colorSelect.value : "any";

  // Find all regenerate buttons
  let regenerateButtons = Array.from(
    document.querySelectorAll(".generate-btn")
  );

  // Filter buttons based on the selected color
  if (selectedColor !== "any") {
    regenerateButtons = regenerateButtons.filter((button) => {
      const cell = button.closest("td");

      if (selectedColor === "none") {
        // Select cells that don't have any color approval
        return (
          !cell.classList.contains("approved") &&
          !cell.classList.contains("yellow-approved") &&
          !cell.classList.contains("red-approved")
        );
      } else if (selectedColor === "green") {
        // Select cells with green approval
        return cell.classList.contains("approved");
      } else if (selectedColor === "yellow") {
        // Select cells with yellow approval
        return cell.classList.contains("yellow-approved");
      } else if (selectedColor === "red") {
        // Select cells with red approval
        return cell.classList.contains("red-approved");
      }

      return true;
    });
  }

  if (regenerateButtons.length === 0) {
    showNotification(
      `No cells to regenerate with filter: ${selectedColor}`,
      "info"
    );
    return;
  }

  const regenerateAllBtn = document.getElementById("regenerate-all-btn");
  const originalContent = regenerateAllBtn.innerHTML;
  let buttonCount = regenerateButtons.length;

  regenerateAllBtn.innerHTML = `<span class="loading-spinner"></span> Processing...`;
  regenerateAllBtn.disabled = true;

  showNotification(`Starting regeneration of ${buttonCount} cells...`, "info");

  // Collect all row IDs to send to backend
  const rowIds = regenerateButtons.map((button) =>
    parseInt(button.getAttribute("data-row"))
  );

  // Set all buttons to loading state
  regenerateButtons.forEach((button) => {
    button.innerHTML = '<span class="loading-spinner"></span>';
    button.disabled = true;
  });

  // Call backend endpoint to process all rows in parallel
  fetch("/regenerate_multiple_cells", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ row_ids: rowIds }),
  })
    .then((response) => {
      if (!response.ok)
        throw new Error(`HTTP error! status: ${response.status}`);
      return response.json();
    })
    .then((data) => {
      if (data.status === "success") {
        showNotification(data.message, "success");

        // Process each result and update UI
        data.results.forEach((result) => {
          if (result.status === "success") {
            // Find the button for this row
            const button = regenerateButtons.find(
              (btn) => parseInt(btn.getAttribute("data-row")) === result.row_idx
            );

            if (button) {
              const container = button.closest(".cell-container");
              const contentDiv = container.querySelector(".cell-content");
              const textArea = container.querySelector(".editable");
              const row = container.closest("tr");
              const colAContent = row.querySelector(
                "td:nth-child(2) .cell-content"
              );
              const cell = contentDiv.closest("td");

              // 1. Update Col B content
              contentDiv.innerHTML = result.highlighted_html;
              if (textArea) textArea.value = result.new_text;

              // 2. Update Col A content (if provided)
              if (result.highlighted_a_html && colAContent) {
                colAContent.innerHTML = result.highlighted_a_html;
              }

              // 3. Update diff status
              cell.classList.remove("same", "different");
              cell.classList.add(result.diff_status);

              // 4. Update color approval status
              cell.classList.remove(
                "approved",
                "yellow-approved",
                "red-approved"
              );

              if (result.col_b_approved) {
                const classMap = {
                  green: "approved",
                  yellow: "yellow-approved",
                  red: "red-approved",
                };
                const approvalClass = classMap[result.col_b_type] || "approved";
                cell.classList.add(approvalClass);
              }

              // 5. Re-setup highlighting
              setupDiffHighlighting(row);
            }
          } else {
            console.error(
              `Error regenerating row ${result.row_idx}:`,
              result.message
            );
            showNotification(
              `Error regenerating row ${result.row_idx}: ${result.message}`,
              "error"
            );
          }
        });
      } else {
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error regenerating cells:", error);
      showNotification("Error regenerating cells: " + error.message, "error");
    })
    .finally(() => {
      // Restore all buttons to original state
      regenerateButtons.forEach((button) => {
        button.innerHTML = '<i class="material-icons">offline_bolt</i>';
        button.disabled = false;
      });

      // Restore regenerate all button
      regenerateAllBtn.innerHTML = originalContent;
      regenerateAllBtn.disabled = false;
    });
}

// Initialize everything when DOM is loaded - optimized with event delegation
document.addEventListener("DOMContentLoaded", () => {
  // Cache DOM elements
  dom.notification = document.getElementById("notification");
  dom.selectionButton = document.getElementById("selectionButton");

  // Initialize theme
  initTheme();

  // Setup form handlers
  setupFormHandlers();

  // Setup text selection
  setupTextSelection();

  // Setup comment functionality
  setupCommentFunctionality();

  // Setup sidebar
  setupSidebar();

  setupRecalculateRatioButton();

  // Setup regenerate all button
  const regenerateAllBtn = document.getElementById("regenerate-all-btn");
  if (regenerateAllBtn) {
    regenerateAllBtn.addEventListener("click", regenerateAllCells);
  }

  // Setup custom prompt modal
  const customPromptModal = document.getElementById("customPromptModal");
  const closeCustomPrompt = customPromptModal.querySelector(".close-comment");
  const promptCancelBtn = document.getElementById("promptCancelBtn");
  const promptGenerateBtn = document.getElementById("promptGenerateBtn");
  const customPromptText = document.getElementById("customPromptText");

  function closeCustomPromptModal() {
    // Save any changes to the prompt content before closing
    if (customPromptText && customPromptText.value.trim()) {
      localStorage.setItem("customPrompt", customPromptText.value);
    }
    customPromptModal.style.display = "none";
    enableBodyScroll();
  }

  closeCustomPrompt.addEventListener("click", closeCustomPromptModal);
  promptCancelBtn.addEventListener("click", closeCustomPromptModal);
  promptGenerateBtn.addEventListener("click", executeCustomPrompt);

  // Close when clicking outside modal
  window.addEventListener("click", (e) => {
    if (e.target === customPromptModal) {
      closeCustomPromptModal();
    }
  });

  // Global escape key handler for all modals
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      // Handle all potential modals
      const commentModal = document.getElementById("commentModal");
      const sidebar = document.getElementById("settings-sidebar");

      if (customPromptModal.style.display === "flex") {
        closeCustomPromptModal();
      }

      if (commentModal && commentModal.style.display === "flex") {
        document.querySelector(".close-comment").click();
      }

      if (sidebar && sidebar.classList.contains("active")) {
        document.getElementById("sidebar-close").click();
      }

      // Handle any other active modals here
    }
  });

  // Setup placeholder insertion
  setupCustomPromptPlaceholders();

  // Event delegation for common elements
  document.addEventListener("click", (e) => {
    const editBtn = e.target.closest(".edit-btn");
    if (editBtn) {
      makeEditable(
        editBtn.closest(".cell-container").querySelector(".cell-content")
      );
      return;
    }

    const saveBtn = e.target.closest(".save-btn");
    if (saveBtn) {
      console.log("Save button clicked:", saveBtn); // Debug log
      saveCell(saveBtn);
      return;
    }

    const generateBtn = e.target.closest(".generate-btn");
    if (generateBtn) {
      regenerateCell(generateBtn);
      return;
    }

    const regenerateBtn1 = e.target.closest(".regenerate-btn-1");
    if (regenerateBtn1) {
      regenerateWithPrompt1(regenerateBtn1);
      return;
    }

    const regenerateBtn2 = e.target.closest(".regenerate-btn-2");
    if (regenerateBtn2) {
      regenerateWithPrompt2(regenerateBtn2);
      return;
    }

    const customPromptBtn = e.target.closest(".custom-prompt-btn");
    if (customPromptBtn) {
      regenerateWithCustomPrompt(customPromptBtn);
      return;
    }

    // Comment button click is handled by its specific listener setup in setupCommentFunctionality
  });

  // Double click to edit cells
  document.querySelectorAll(".cell-content[data-editable]").forEach((div) => {
    div.addEventListener("dblclick", function () {
      makeEditable(this);
    });
  });

  // Setup diff highlighting for all rows
  document.querySelectorAll("tbody tr").forEach(setupDiffHighlighting);

  // Setup Arabic button click
  document.querySelectorAll(".show-arabic-btn").forEach((btn) => {
    btn.addEventListener("click", function (event) {
      event.stopPropagation();
      showArabicText(event.currentTarget);
    });
  });

  // Setup show Bangla button click
  document.querySelectorAll(".show-bangla-btn").forEach((btn) => {
    btn.addEventListener("click", function (event) {
      event.stopPropagation();
      showBanglaText(event.currentTarget);
    });
  });

  // Setup generate Bangla button click
  document.querySelectorAll(".generate-bangla-btn").forEach((btn) => {
    btn.addEventListener("click", function (event) {
      event.stopPropagation();
      generateBanglaText(event.currentTarget);
    });
  });
});

// Comment functionality
function setupCommentFunctionality() {
  const commentModal = document.getElementById("commentModal");
  const commentText = document.getElementById("commentText");
  const commentRowIdx = document.getElementById("commentRowIdx");
  const commentPage = document.getElementById("commentPage");
  const commentRowsPerPage = document.getElementById("commentRowsPerPage");
  const commentSaveBtn = document.getElementById("commentSaveBtn");
  const commentCancelBtn = document.getElementById("commentCancelBtn");
  const closeComment = document.querySelector(".close-comment");

  // Open comment modal
  document.addEventListener("click", (e) => {
    const commentBtn = e.target.closest(".comment-btn");
    if (commentBtn) {
      const rowIdx = commentBtn.getAttribute("data-row");
      const page = commentBtn.getAttribute("data-page");
      const rowsPerPage = commentBtn.getAttribute("data-rows");

      // Get existing comment if any
      fetch(`/get_comment?row_idx=${rowIdx}`)
        .then((response) => response.json())
        .then((data) => {
          commentText.value = data.comment || "";
          commentRowIdx.value = rowIdx;
          commentPage.value = page;
          commentRowsPerPage.value = rowsPerPage;
          commentModal.style.display = "flex";
          disableBodyScroll();
        })
        .catch((error) => {
          console.error("Error fetching comment:", error);
          commentText.value = "";
          commentRowIdx.value = rowIdx;
          commentPage.value = page;
          commentRowsPerPage.value = rowsPerPage;
          commentModal.style.display = "flex";
          disableBodyScroll();
        });
    }
  });

  // Close modal functions
  function closeModal() {
    commentModal.style.display = "none";
    enableBodyScroll();
  }

  closeComment.addEventListener("click", closeModal);
  commentCancelBtn.addEventListener("click", closeModal);

  // Close when clicking outside modal
  window.addEventListener("click", (e) => {
    if (e.target === commentModal) {
      closeModal();
    }
  });

  // Save comment
  commentSaveBtn.addEventListener("click", () => {
    const formData = new FormData();
    formData.append("row_idx", commentRowIdx.value);
    formData.append("comment", commentText.value);

    fetch("/save_comment", {
      method: "POST",
      body: formData,
    })
      .then((response) => response.json())
      .then((data) => {
        if (data.status === "success") {
          showNotification("Comment saved successfully!");
          closeModal();
        } else {
          showNotification("Error: " + data.message, "error");
        }
      })
      .catch((error) => {
        console.error("Error saving comment:", error);
        showNotification("Error saving comment", "error");
      });
  });
}

// Add near the top or within DOMContentLoaded
function setupSidebar() {
  const sidebar = document.getElementById("settings-sidebar");
  const toggleBtn = document.getElementById("sidebar-toggle");
  const closeBtn = document.getElementById("sidebar-close");
  const overlay = document.getElementById("sidebar-overlay");

  if (!sidebar || !toggleBtn || !closeBtn || !overlay) {
    console.warn("Sidebar elements not found. Skipping sidebar setup.");
    return;
  }

  function openSidebar() {
    sidebar.classList.add("active");
    overlay.classList.add("active");
    disableBodyScroll();
    // Optional: Add class to body if needed for layout adjustments
    // document.body.classList.add('sidebar-active');
  }

  function closeSidebar() {
    sidebar.classList.remove("active");
    overlay.classList.remove("active");
    enableBodyScroll();
    // document.body.classList.remove('sidebar-active');
  }

  toggleBtn.addEventListener("click", (e) => {
    e.stopPropagation(); // Prevent triggering body click listener
    openSidebar();
  });

  closeBtn.addEventListener("click", closeSidebar);
  overlay.addEventListener("click", closeSidebar);

  // Close sidebar if clicked outside of it
  // document.addEventListener('click', (e) => {
  //     if (sidebar.classList.contains('active') && !sidebar.contains(e.target) && e.target !== toggleBtn) {
  //          closeSidebar();
  //     }
  // });

  // Prevent clicks inside sidebar from closing it
  sidebar.addEventListener("click", (e) => {
    e.stopPropagation();
  });
}

// Show Arabic text for a cell
function showArabicText(button) {
  const row = button.getAttribute("data-row");
  if (!row) {
    console.error("No row index found on button");
    showNotification("Error: Missing row index", "error");
    return;
  }

  const cellContainer = button.closest(".cell-container");
  const contentDiv = cellContainer.querySelector(".cell-content");

  // Reset other buttons
  const showBanglaBtn = cellContainer.querySelector(".show-bangla-btn");
  const generateBanglaBtn = cellContainer.querySelector(".generate-bangla-btn");

  if (showBanglaBtn) {
    showBanglaBtn.classList.remove("active");
    // Reset Bangla button icon
    showBanglaBtn.querySelector("i").textContent = "spellcheck";
  }
  if (generateBanglaBtn) generateBanglaBtn.classList.remove("active");

  // If button is already active, show original content
  if (button.classList.contains("active")) {
    if (contentDiv.hasAttribute("data-original-content")) {
      contentDiv.innerHTML = contentDiv.getAttribute("data-original-content");
      button.classList.remove("active");
      // Change icon back to translate for Arabic
      button.querySelector("i").textContent = "translate";
    }
    return;
  }

  // Save original content if not already saved
  if (!contentDiv.hasAttribute("data-original-content")) {
    contentDiv.setAttribute("data-original-content", contentDiv.innerHTML);
  }

  // Get the actual row value (adjust for Excel row numbering)
  const actualRow = parseInt(row) + 2;

  // Fetch Arabic text
  fetch(`/get_arabic_text?row_idx=${actualRow}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.status === "success") {
        contentDiv.innerHTML = data.arabic_text.replace(/\n/g, "<br>");
        button.classList.add("active");
        // Change icon to description for Arabic
        button.querySelector("i").textContent = "description";

        // Save Arabic content for future use
        contentDiv.setAttribute("data-arabic-content", contentDiv.innerHTML);
      } else {
        console.error("Error fetching Arabic:", data.message);
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error fetching Arabic text:", error);
      showNotification("Error fetching Arabic text", "error");
    });
}

// Show previously generated Bangla text
function showBanglaText(button) {
  const row = button.getAttribute("data-row");
  if (!row) {
    console.error("No row index found on button");
    showNotification("Error: Missing row index", "error");
    return;
  }

  const cellContainer = button.closest(".cell-container");
  const contentDiv = cellContainer.querySelector(".cell-content");

  // Reset other buttons
  const showArabicBtn = cellContainer.querySelector(".show-arabic-btn");
  const generateBanglaBtn = cellContainer.querySelector(".generate-bangla-btn");

  if (showArabicBtn) {
    showArabicBtn.classList.remove("active");
    // Reset Arabic button icon
    showArabicBtn.querySelector("i").textContent = "translate";
  }
  if (generateBanglaBtn) generateBanglaBtn.classList.remove("active");

  // If button is already active, show original content
  if (button.classList.contains("active")) {
    if (contentDiv.hasAttribute("data-original-content")) {
      contentDiv.innerHTML = contentDiv.getAttribute("data-original-content");
      button.classList.remove("active");
      // Reset icon to default spellcheck
      button.querySelector("i").textContent = "spellcheck";

      // Reset all buttons to inactive state to clearly indicate we're showing original content
      if (showArabicBtn) {
        showArabicBtn.classList.remove("active");
        showArabicBtn.querySelector("i").textContent = "translate";
      }
      if (generateBanglaBtn) {
        generateBanglaBtn.classList.remove("active");
      }
    }
    return;
  }

  // If bangla content exists, show it
  if (contentDiv.hasAttribute("data-bangla-content")) {
    // Save original content if not already saved
    if (!contentDiv.hasAttribute("data-original-content")) {
      contentDiv.setAttribute("data-original-content", contentDiv.innerHTML);
    }

    contentDiv.innerHTML = contentDiv.getAttribute("data-bangla-content");
    button.classList.add("active");

    // Change icon based on the source of Bangla content
    const banglaSource = contentDiv.getAttribute("data-bangla-source");
    if (banglaSource === "ai") {
      // AI-translated content
      button.querySelector("i").textContent = "auto_stories";
    } else {
      // Original or manually translated Bangla content
      button.querySelector("i").textContent = "translate";
    }
  } else {
    // If no Bangla translation exists yet, generate one
    generateBanglaText(cellContainer.querySelector(".generate-bangla-btn"));
  }
}

// Generate new Bangla translation
function generateBanglaText(button) {
  const row = button.getAttribute("data-row");
  if (!row) {
    console.error("No row index found on button");
    showNotification("Error: Missing row index", "error");
    return;
  }

  const cellContainer = button.closest(".cell-container");
  const contentDiv = cellContainer.querySelector(".cell-content");

  // Reset all button states to ensure clear starting point
  const showArabicBtn = cellContainer.querySelector(".show-arabic-btn");
  const showBanglaBtn = cellContainer.querySelector(".show-bangla-btn");

  // If button is already active, toggle back to original content
  if (button.classList.contains("active")) {
    if (contentDiv.hasAttribute("data-original-content")) {
      // Restore original content
      contentDiv.innerHTML = contentDiv.getAttribute("data-original-content");

      // Reset all button states
      button.classList.remove("active");

      if (showArabicBtn) {
        showArabicBtn.classList.remove("active");
        showArabicBtn.querySelector("i").textContent = "translate";
      }

      if (showBanglaBtn) {
        showBanglaBtn.classList.remove("active");
        showBanglaBtn.querySelector("i").textContent = "spellcheck";
      }
    }
    return;
  }

  // Save original content if not already saved
  if (!contentDiv.hasAttribute("data-original-content")) {
    contentDiv.setAttribute("data-original-content", contentDiv.innerHTML);
  }

  // Get the actual row value (adjust for Excel row numbering)
  const actualRow = parseInt(row) + 2;

  // Show loading spinner and disable button
  const originalContent = button.innerHTML;
  button.innerHTML = '<span class="loading-spinner"></span>';
  button.disabled = true;

  // Show loading notification
  showNotification("Translating to Bangla...", "info");

  // Fetch AI-translated Bangla text
  fetch(`/translate_arabic_to_bangla?row_idx=${actualRow}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.status === "success") {
        // Update content
        contentDiv.innerHTML = data.translated_bangla.replace(/\n/g, "<br>");

        // Save Bangla content for future use
        contentDiv.setAttribute("data-bangla-content", contentDiv.innerHTML);
        contentDiv.setAttribute("data-bangla-source", "ai"); // Mark as AI-generated

        // Reset all button states first
        if (showArabicBtn) {
          showArabicBtn.classList.remove("active");
          showArabicBtn.querySelector("i").textContent = "translate";
        }

        // Set button states to show we're displaying AI-translated Bangla
        button.classList.add("active");

        if (showBanglaBtn) {
          showBanglaBtn.classList.add("active");
          showBanglaBtn.querySelector("i").textContent = "auto_stories";
          showBanglaBtn.setAttribute("data-has-translation", "true");
        }

        showNotification("Translation completed!");
      } else {
        console.error("Error translating to Bangla:", data.message);
        showNotification("Error: " + data.message, "error");
      }
    })
    .catch((error) => {
      console.error("Error fetching Bangla translation:", error);
      showNotification("Error fetching Bangla translation", "error");
    })
    .finally(() => {
      // Restore button's original appearance when done
      button.innerHTML = originalContent;
      button.disabled = false;
    });
}

function setupRecalculateRatioButton() {
  const recalculateBtn = document.getElementById("recalculate-ratio-btn");

  if (!recalculateBtn) {
    console.warn("Recalculate ratio button not found. Skipping setup.");
    return;
  }

  recalculateBtn.addEventListener("click", function () {
    recalculateBtn.disabled = true;
    recalculateBtn.textContent = "Recalculating...";

    fetch("/recalculate_ratios", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
    })
      .then((response) => response.json())
      .then((data) => {
        if (data.status === "success") {
          showNotification(data.message);
          if (
            confirm(
              "Ratios have been recalculated. Reload the page to see the updated values?"
            )
          ) {
            window.location.reload();
          }
        } else {
          showNotification(data.message, "error");
        }
      })
      .catch((error) => {
        console.error("Error recalculating ratios:", error);
        showNotification("Failed to recalculate ratios", "error");
      })
      .finally(() => {
        recalculateBtn.disabled = false;
        recalculateBtn.textContent = "Recalculate Ratios";
      });
  });
}
