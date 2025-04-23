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

    // Comment button click is handled by its specific listener setup in setupCommentFunctionality
  });

  // Double click to edit cells
  document.querySelectorAll(".cell-content").forEach((div) => {
    div.addEventListener("dblclick", function () {
      makeEditable(this);
    });
  });

  // Setup diff highlighting for all rows
  document.querySelectorAll("tbody tr").forEach(setupDiffHighlighting);

  // Add event listeners for toggle and translate buttons
  document.querySelectorAll(".toggle-arabic-btn").forEach((btn) => {
    btn.addEventListener("click", function (event) {
      event.stopPropagation();
      toggleArabicText(event.currentTarget);
    });
  });

  document.querySelectorAll(".translate-bangla-btn").forEach((btn) => {
    btn.addEventListener("click", function (event) {
      event.stopPropagation();
      translateToBangla(event.currentTarget);
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
          commentModal.style.display = "block";
        })
        .catch((error) => {
          console.error("Error fetching comment:", error);
          commentText.value = "";
          commentRowIdx.value = rowIdx;
          commentPage.value = page;
          commentRowsPerPage.value = rowsPerPage;
          commentModal.style.display = "block";
        });
    }
  });

  // Close modal functions
  function closeModal() {
    commentModal.style.display = "none";
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
    // Optional: Add class to body if needed for layout adjustments
    // document.body.classList.add('sidebar-active');
  }

  function closeSidebar() {
    sidebar.classList.remove("active");
    overlay.classList.remove("active");
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

// Modified toggleArabicText to handle state transitions
function toggleArabicText(button) {
  console.log("Toggle Arabic for button:", button);
  const row = button.getAttribute("data-row");
  if (!row) {
    console.error("No row index found on toggle button");
    showNotification("Error: Missing row index", "error");
    return;
  }

  console.log("Toggle Arabic for row:", row);

  const cellContainer = button.closest(".cell-container");
  const contentDiv = cellContainer.querySelector(".cell-content");
  const translateButton = cellContainer.querySelector(".translate-bangla-btn");

  // Check current state
  const isShowingArabic = button.classList.contains("showing-arabic");
  const isShowingBangla = button.classList.contains("showing-bangla");

  if (isShowingArabic) {
    // Switch back to original content
    if (contentDiv.hasAttribute("data-original-content")) {
      contentDiv.innerHTML = contentDiv.getAttribute("data-original-content");
      button.classList.remove("showing-arabic");
      button.querySelector("i").textContent = "translate";
    }
  } else if (isShowingBangla) {
    // Switch to Arabic text
    const actualRow = parseInt(row) + 2; // Adjust for 0-based indexing and header row
    fetch(`/get_arabic_text?row_idx=${actualRow}`)
      .then((response) => response.json())
      .then((data) => {
        if (data.status === "success") {
          contentDiv.innerHTML = data.arabic_text.replace(/\n/g, "<br>");
          button.classList.remove("showing-bangla");
          button.classList.add("showing-arabic");
          button.querySelector("i").textContent = "description";
          // Reset translate button if exists
          if (translateButton) {
            translateButton.classList.remove("showing-bangla");
          }
        } else {
          console.error("Error fetching Arabic:", data.message);
          showNotification("Error: " + data.message, "error");
        }
      })
      .catch((error) => {
        console.error("Error fetching Arabic text:", error);
        showNotification("Error fetching Arabic text", "error");
      });
  } else {
    // Switch to Arabic text
    const actualRow = parseInt(row) + 2;
    fetch(`/get_arabic_text?row_idx=${actualRow}`)
      .then((response) => response.json())
      .then((data) => {
        if (data.status === "success") {
          if (!contentDiv.hasAttribute("data-original-content")) {
            contentDiv.setAttribute(
              "data-original-content",
              contentDiv.innerHTML
            );
          }
          contentDiv.innerHTML = data.arabic_text.replace(/\n/g, "<br>");
          button.classList.add("showing-arabic");
          button.querySelector("i").textContent = "description";
          // Reset translate button if exists
          if (translateButton) {
            translateButton.classList.remove("showing-bangla");
          }
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
}

// New function to handle AI translation to Bangla
function translateToBangla(button) {
  const row = button.getAttribute("data-row");
  if (!row) {
    console.error("No row index found on translate button");
    showNotification("Error: Missing row index", "error");
    return;
  }

  console.log("Translate to Bangla for row:", row);

  const cellContainer = button.closest(".cell-container");
  const contentDiv = cellContainer.querySelector(".cell-content");
  const toggleArabicBtn = cellContainer.querySelector(".toggle-arabic-btn");

  // Save original content if not already saved
  if (!contentDiv.hasAttribute("data-original-content")) {
    contentDiv.setAttribute("data-original-content", contentDiv.innerHTML);
  }

  // Get the actual row value (adjust for Excel row numbering)
  const actualRow = parseInt(row) + 2;

  // Show loading state
  showNotification("Translating to Bangla...", "info");

  // Fetch AI-translated Bangla text
  fetch(`/translate_arabic_to_bangla?row_idx=${actualRow}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.status === "success") {
        contentDiv.innerHTML = data.translated_bangla.replace(/\n/g, "<br>");
        button.classList.add("showing-bangla");
        // Update toggle button state
        if (toggleArabicBtn) {
          toggleArabicBtn.classList.remove("showing-arabic");
          toggleArabicBtn.classList.add("showing-bangla");
          toggleArabicBtn.querySelector("i").textContent = "g_translate";
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
