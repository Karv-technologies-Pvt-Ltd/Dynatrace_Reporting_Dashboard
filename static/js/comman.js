// ===============================
// Auto logout after inactivity
// ===============================
let timer;
const logoutTime = 5 * 60 * 1000; // 5 minutes

function resetTimer() {
  clearTimeout(timer);
  timer = setTimeout(() => {
    window.location.href = "/logout";
  }, logoutTime);
}

["mousemove", "keypress", "click", "scroll"].forEach(event =>
  document.addEventListener(event, resetTimer)
);

resetTimer();

// ===============================
// Show / hide access token
// ===============================
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("toggleToken");

  if (btn) {
    btn.addEventListener("click", () => {
      const input = document.getElementById("access_token");
      const show = input.type === "password";
      input.type = show ? "text" : "password";
      btn.textContent = show ? "Hide" : "Show";
    });
  }

  // ===============================
  // Filter management zones
  // ===============================
  const filterInput = document.getElementById("mzFilter");
  const zoneSelect = document.getElementById("management_zone");

  if (filterInput && zoneSelect) {
    filterInput.addEventListener("input", () => {
      const query = filterInput.value.toLowerCase();

      Array.from(zoneSelect.options).forEach(opt => {
        if (opt.value === "All") {
          opt.hidden = false;
          return;
        }
        opt.hidden = !opt.text.toLowerCase().includes(query);
      });
    });
  }
});
