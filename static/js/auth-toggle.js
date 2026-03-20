document.addEventListener("DOMContentLoaded", () => {
    console.log("auth-toggle.js loaded");
    const messages = document.querySelector(".django-messages");
    if (messages) {
        setTimeout(() => {
            messages.style.display = "none";
        }, 3000);
    }
});
// function togglePassword(id, icon) {
//   const input = document.getElementById(id);

//   if (input.type === "password") {
//     input.type = "text";
//     icon.classList.remove("bi-eye");
//     icon.classList.add("bi-eye-slash");
//   } else {
//     input.type = "password";
//     icon.classList.remove("bi-eye-slash");
//     icon.classList.add("bi-eye");
//   }
// }
