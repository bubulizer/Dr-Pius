(() => {
  const config = window.KIU_AUDIT_FORMS || {};
  const urls = { phd: config.phd || "", masters: config.masters || "" };
  const notice = document.querySelector("#setup-notice");
  let missing = false;

  document.querySelectorAll(".form-link").forEach(link => {
    const url = urls[link.dataset.form];
    if (!url || !/^https:\/\/docs\.google\.com\/forms\//.test(url)) {
      missing = true;
      link.setAttribute("aria-disabled", "true");
      link.addEventListener("click", event => event.preventDefault());
      return;
    }
    link.href = url;
    link.target = "_blank";
    link.rel = "noopener";
  });

  document.querySelectorAll(".share-link").forEach(button => {
    const level = button.dataset.form;
    const url = urls[level];
    if (!url || !/^https:\/\/docs\.google\.com\/forms\//.test(url)) {
      button.setAttribute("aria-disabled", "true");
      button.disabled = true;
      return;
    }
    button.addEventListener("click", () => {
      const label = level === "phd" ? "PhD" : "Master's";
      const text = `KIU Western Campus ${label} students' private audit form: ${url}`;
      window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, "_blank", "noopener");
    });
  });

  if (missing) notice.hidden = false;
})();
