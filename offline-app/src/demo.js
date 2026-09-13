/* Guided demo. Plain script (not a module) so it keeps working even if app.js
   fails to load, e.g. when index.html is opened directly as a file:// page. */
(() => {
  const panel = document.querySelector("#demo-panel");
  const guide = document.querySelector("#guide-character");
  const title = document.querySelector("#demo-title");
  const message = document.querySelector("#demo-message");
  const expression = document.querySelector("#expression");
  const stepButtons = Array.from(document.querySelectorAll(".demo-step"));
  const backButton = document.querySelector("#demo-back");
  const nextButton = document.querySelector("#demo-next");
  const workspaceSteps = document.querySelectorAll(".workspace .step");
  const seenKey = "slideform-demo-seen";

  const steps = [
    { title: "Choose the presentation", message: "Select the PowerPoint file you want to convert. Slideform reads only the file you choose on this device.", expression: "Focused", image: "assets/guide-character.png" },
    { title: "Confirm the selection", message: "Slides containing Demand Priority are selected automatically. Update the slide numbers or choose all native tables when needed.", expression: "Reviewing", image: "assets/guide-reviewing.png" },
    { title: "Download the workbook", message: "Select Convert to Excel. The completed workbook downloads with the same name as your source presentation.", expression: "Complete", image: "assets/guide-complete.png" },
  ];

  // Preload so switching steps never shows a blank frame while the image fetches.
  steps.forEach((step) => { new Image().src = step.image; });

  let index = 0;

  function show(next) {
    index = Math.max(0, Math.min(steps.length - 1, next));
    const step = steps[index];
    title.textContent = step.title;
    message.textContent = step.message;
    expression.textContent = step.expression;
    guide.src = step.image;
    stepButtons.forEach((button, buttonIndex) => {
      const active = buttonIndex === index;
      button.classList.toggle("active", active);
      button.setAttribute("aria-selected", String(active));
      button.tabIndex = active ? 0 : -1;
    });
    workspaceSteps.forEach((step, stepIndex) => step.classList.toggle("demo-focus", stepIndex === index));
    nextButton.textContent = index === steps.length - 1 ? "Start converting" : "Next step";
    backButton.disabled = index === 0;
  }

  function open() {
    panel.classList.add("open");
    panel.setAttribute("aria-hidden", "false");
    show(0);
    document.querySelector("#demo-close").focus();
  }

  function close() {
    if (panel.contains(document.activeElement)) document.querySelector("#demo-trigger").focus();
    panel.classList.remove("open");
    panel.setAttribute("aria-hidden", "true");
    workspaceSteps.forEach((step) => step.classList.remove("demo-focus"));
  }

  document.querySelector("#demo-trigger").addEventListener("click", open);
  document.querySelector("#demo-close").addEventListener("click", close);
  backButton.addEventListener("click", () => show(index - 1));
  nextButton.addEventListener("click", () => {
    if (index === steps.length - 1) { close(); return; }
    show(index + 1);
  });
  stepButtons.forEach((button) => button.addEventListener("click", () => show(Number(button.dataset.step))));
  panel.addEventListener("click", (event) => { if (event.target === panel) close(); });
  panel.addEventListener("keydown", (event) => {
    if (event.key === "Escape") close();
    else if (event.key === "ArrowRight") show(index + 1);
    else if (event.key === "ArrowLeft") show(index - 1);
  });

  // Show new visitors the guide once, automatically.
  try {
    if (!localStorage.getItem(seenKey)) {
      localStorage.setItem(seenKey, "1");
      open();
    }
  } catch (_error) {
    // Storage may be unavailable (private browsing); the demo stays manual-open only.
  }
})();
