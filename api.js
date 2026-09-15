// ==========================================================================
// 🔗 API CLIENT FOR GITHUB PAGES
// ==========================================================================
const BACKEND_API_URL = "https://script.google.com/macros/s/AKfycbxe6vrQshGAPglgmmZykljXazDNh1ZwWSRdFDAip1WWWRky24lmBo-qicLK3_i10AvT/exec";
const TOKEN_KEY = "school_persistent_token";

async function callBackendApi(action, data = {}) {
  try {
    const res = await fetch(BACKEND_API_URL, {
      method: "POST",
      // text/plain ជួយកាត់បន្ថយបញ្ហា CORS Preflight នៅលើ Browser
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: action, data: data })
    });
    if (!res.ok) throw new Error("HTTP error " + res.status);
    return await res.json();
  } catch (err) {
    console.error("API Call Error:", err);
    throw err;
  }
}

function getStoredToken() {
  return localStorage.getItem(TOKEN_KEY) || "";
}

function setStoredToken(token) {
  localStorage.setItem(TOKEN_KEY, token);
}

function removeStoredToken() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem("krou_teacher_sex");
  localStorage.removeItem("krou_active_page");
}
