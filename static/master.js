/* master.js — FOET Landing Page: Refresh button only */

const syncBtn = document.getElementById("sync-btn");

function setSyncState(on) {
  if (!syncBtn) return;
  syncBtn.classList.toggle("syncing", on);
  syncBtn.disabled = on;
  const txt = syncBtn.querySelector(".sync-btn-text");
  if (txt) txt.textContent = on ? "Refreshing..." : "Refresh";
}

async function masterSync() {
  setSyncState(true);
  try {
    const res  = await fetch("/api/departments");
    const data = await res.json();
    if (data.departments) {
      await Promise.all(
        data.departments
          .filter(d => d.sync_configured)
          .map(d => fetch(`/api/${d.key}/sync`, { method: "POST" }).catch(() => {}))
      );
    }
  } catch (_) { /* silent */ }
  finally {
    setSyncState(false);
  }
}
