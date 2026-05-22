document.querySelectorAll('form[action="#"]').forEach((form) => {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector('button[type="submit"]');
    if (!button) return;
    const original = button.textContent;
    button.textContent = "Signal Sent";
    button.disabled = true;
    window.setTimeout(() => {
      button.textContent = original;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
});

document.querySelectorAll(".nav-toggle").forEach((button) => {
  const targetId = button.getAttribute("aria-controls");
  const nav = targetId ? document.getElementById(targetId) : null;
  if (!nav) return;

  button.addEventListener("click", () => {
    const isOpen = button.getAttribute("aria-expanded") === "true";
    button.setAttribute("aria-expanded", String(!isOpen));
    nav.classList.toggle("is-open", !isOpen);
  });

  nav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      button.setAttribute("aria-expanded", "false");
      nav.classList.remove("is-open");
    });
  });
});

const snippets = [
  [
    'app.scene.load("orbital-hangar");',
    'app.camera.flyTo("command-deck");',
    'app.collab.syncPresence();',
    'app.deploy.target("webxr-preview");'
  ],
  [
    'project.branch("launch-sequence");',
    'materials.set("carbon-alloy");',
    'physics.gravity(0, -9.8, 0);',
    'runtime.publish("edge-region-04");'
  ],
  [
    'viewer.attach("commerce-configurator");',
    'analytics.track("hotspot-opened");',
    'xr.enter("product-walkthrough");',
    'scene.snapshot("stakeholder-review");'
  ]
];

const codeTarget = document.getElementById("code-snippet");
if (codeTarget) {
  let pointer = 0;
  window.setInterval(() => {
    pointer = (pointer + 1) % snippets.length;
    codeTarget.textContent = snippets[pointer].join("\n");
  }, 2800);
}

document.querySelectorAll("[data-industry-explorer]").forEach((explorer) => {
  const tabs = explorer.querySelectorAll(".industry-tab");
  const images = explorer.querySelectorAll(".industry-image");
  const cards = explorer.querySelectorAll(".industry-info-card");

  const activateIndustry = (industry) => {
    tabs.forEach((tab) => {
      const isActive = tab.dataset.industry === industry;
      tab.classList.toggle("is-active", isActive);
      tab.setAttribute("aria-selected", String(isActive));
    });

    images.forEach((image) => {
      image.classList.toggle("is-active", image.dataset.industryPanel === industry);
    });

    cards.forEach((card) => {
      card.classList.toggle("is-active", card.dataset.industryCopy === industry);
    });
  };

  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      activateIndustry(tab.dataset.industry);
    });
  });
});
