(() => {
  const settings = window.KIU_SEAS_AUDIT || {};
  const formUrl = typeof settings.studentForm === "string" ? settings.studentForm.trim() : "";
  const deadline = typeof settings.deadline === "string" ? settings.deadline.trim() : "";
  const formLink = document.querySelector("#student-form-link");
  const shareButton = document.querySelector("#share-link");
  const setupNotice = document.querySelector("#setup-notice");
  const deadlineText = document.querySelector("#deadline");

  if (deadline && deadlineText) deadlineText.textContent = deadline;

  if (!/^https:\/\/docs\.google\.com\/forms\//i.test(formUrl)) {
    setupNotice.hidden = false;
    formLink.setAttribute("aria-disabled", "true");
    formLink.removeAttribute("target");
    shareButton.disabled = true;
    return;
  }

  formLink.href = formUrl;
  shareButton.addEventListener("click", () => {
    const pageUrl = window.location.href.split("#")[0];
    const message = [
      "*KIU Western Campus - SEAS Postgraduate Students' Audit*",
      "",
      "All registered MSc and PhD students should complete the private audit and progress-monitoring form:",
      pageUrl,
      "",
      "Do not post personal or academic information in this WhatsApp group.",
      "",
      "*Issued through the Associate Dean Research via the departmental Research Coordinator*"
    ].join("\n");
    window.open(`https://wa.me/?text=${encodeURIComponent(message)}`, "_blank", "noopener,noreferrer");
  });
})();
