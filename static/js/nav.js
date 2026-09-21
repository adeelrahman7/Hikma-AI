// Shared sidebar/nav logic for landingpage.html, upload.html, and study.html.
// Single source of truth for header and sidebar behavior.

const PAGE_TRANSITION_MS = 300;

// Call this instead of setting window.location.href directly, from any
// inline script (button onclick handlers, fetch callbacks, etc.), so
// programmatic navigation gets the same slide-out as clicked links.
function navigateWithTransition(url) {
  document.body.classList.add("page-exit");
  setTimeout(() => {
    window.location.href = url;
  }, PAGE_TRANSITION_MS);
}
window.navigateWithTransition = navigateWithTransition;

// Intercepts clicks on same-page-set links (relative/absolute .html hrefs)
// so the current page slides out before the browser navigates away.
// Leaves external links, hash links, and new-tab links alone.
function interceptInternalLinks() {
  document.addEventListener("click", (e) => {
    const link = e.target.closest("a");
    if (!link) return;

    const href = link.getAttribute("href");
    if (!href) return;
    if (href.startsWith("#")) return;
    if (href.startsWith("http://") || href.startsWith("https://")) return;
    if (link.target === "_blank") return;
    if (!href.endsWith(".html")) return;

    e.preventDefault();
    navigateWithTransition(href);
  });
}

document.addEventListener("DOMContentLoaded", interceptInternalLinks);

function initNav() {
  const menuBtn = document.getElementById("menuBtn");
  const sidebar = document.getElementById("sidebar");
  const sidebarOverlay = document.getElementById("sidebarOverlay");
  const sidebarClose = document.getElementById("sidebarClose");
  const authToggleBtn = document.getElementById("authToggleBtn");
  const sidebarAvatar = document.getElementById("sidebarAvatar");
  const sidebarUsername = document.getElementById("sidebarUsername");
  const sidebarEmail = document.getElementById("sidebarEmail");

  function openSidebar() {
    sidebar?.classList.add("open");
    sidebarOverlay?.classList.add("open");
  }

  function closeSidebar() {
    sidebar?.classList.remove("open");
    sidebarOverlay?.classList.remove("open");
  }

  menuBtn?.addEventListener("click", openSidebar);
  sidebarClose?.addEventListener("click", closeSidebar);
  sidebarOverlay?.addEventListener("click", closeSidebar);
  
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeSidebar();
  });

  // Demo login/logout toggle
  let loggedIn = true;
  authToggleBtn?.addEventListener("click", () => {
    loggedIn = !loggedIn;
    if (loggedIn) {
      if (sidebarAvatar) sidebarAvatar.textContent = "AR";
      if (sidebarUsername) sidebarUsername.textContent = "Adeel Rahman";
      if (sidebarEmail) sidebarEmail.textContent = "adeel@hikma.ai";
      authToggleBtn.innerHTML = `
        <span class="sidebar-icon-circle">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M9 20H5.5A1.5 1.5 0 0 1 4 18.5v-13A1.5 1.5 0 0 1 5.5 4H9"/>
            <path d="M16 16l4-4-4-4"/>
            <path d="M20 12H9"/>
          </svg>
        </span>
        Log out
      `;
      authToggleBtn.classList.add("logout");
    } else {
      if (sidebarAvatar) sidebarAvatar.textContent = "?";
      if (sidebarUsername) sidebarUsername.textContent = "Guest";
      if (sidebarEmail) sidebarEmail.textContent = "Not signed in";
      authToggleBtn.innerHTML = `
        <span class="sidebar-icon-circle">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M15 4h3.5A1.5 1.5 0 0 1 20 5.5v13a1.5 1.5 0 0 1-1.5 1.5H15"/>
            <path d="M8 8l-4 4 4 4"/>
            <path d="M4 12h11"/>
          </svg>
        </span>
        Log in
      `;
      authToggleBtn.classList.remove("logout");
    }
  });
}

function loadNav() {
  // 1. Inject Top Navigation Bar with Center Slot
  const navContainer = document.getElementById("main-nav");
  if (navContainer) {
    navContainer.innerHTML = `
      <div class="brand-left">
        <button class="menu-btn" id="menuBtn" aria-label="Open sidebar menu">
          <span></span>
          <span></span>
          <span></span>
        </button>
        <a href="/static/landingpage.html" class="brand-link">
          <img src="/static/favicon.svg" alt="Hikma AI Logo" class="brand-logo" style="width: 28px; height: 28px;" />
          <span class="brand-title">Hikma AI</span>
        </a>
      </div>

      <div class="nav-center" id="navCenter"></div>
    `;
  }

  // 2. Relocate study navigation pill into navbar if present
  const studyPill = document.querySelector(".study-nav-pill");
  const navCenter = document.getElementById("navCenter");
  if (studyPill && navCenter) {
    navCenter.appendChild(studyPill);
    const oldBar = document.querySelector(".study-nav-bar");
    if (oldBar) oldBar.remove();
  }

  // 3. Fetch and Inject Sidebar Partial
  const navRoot = document.getElementById("nav-root");
  if (!navRoot) return;

  fetch("/static/partials/sidebar.html")
    .then(res => {
      if (!res.ok) throw new Error(`Failed to load sidebar partial: HTTP ${res.status}`);
      return res.text();
    })
    .then(html => {
      navRoot.innerHTML = html;
      initNav();
    })
    .catch(err => {
      console.error("Nav failed to load:", err);
    });
}

document.addEventListener("DOMContentLoaded", loadNav);