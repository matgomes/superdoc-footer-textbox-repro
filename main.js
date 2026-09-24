import { SuperDoc, DOCX } from "superdoc";
import "superdoc/style.css";

const params = new URLSearchParams(location.search);
const fixture = params.get("doc") ?? "A-textbox-nowrap-in-pct-table.docx";
const select = document.getElementById("fixture");
const status = document.getElementById("status");

select.value = fixture;
select.addEventListener("change", () => {
  location.search = `?doc=${select.value}`;
});

window.__repro = { fixture, ready: false, exceptions: [] };

function describe(event) {
  const error = event?.error;
  return error instanceof Error ? error.message : String(event?.code ?? error);
}

function reportOverlay() {
  const overlay = document.querySelector(".v2-document-loading-overlay__progress-value");
  window.__repro.overlay = overlay ? overlay.textContent.trim() : null;
  status.textContent = [
    `onReady: ${window.__repro.ready}`,
    `loading overlay: ${window.__repro.overlay ?? "gone"}`,
    `exceptions: ${window.__repro.exceptions.join(" | ") || "none"}`,
  ].join("   ·   ");
}

const response = await fetch(`/${fixture}`);
const file = new File([await response.blob()], fixture, {
  type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
});

new SuperDoc({
  selector: "#editor",
  documentMode: "editing",
  document: { id: "repro", type: DOCX, data: file },
  onReady: () => {
    window.__repro.ready = true;
    reportOverlay();
  },
  onException: (event) => {
    window.__repro.exceptions.push(describe(event));
    reportOverlay();
  },
});

setInterval(reportOverlay, 1000);
