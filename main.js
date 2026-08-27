import * as pdfjs from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import { marked } from "marked";
import * as XLSX from "https://esm.sh/xlsx/xlsx.mjs";

// PDF.js Worker initialization
pdfjs.GlobalWorkerOptions.workerSrc = pdfWorker;

import { GoogleGenerativeAI } from "@google/generative-ai";

// DOM Elements
const chatContainer = document.getElementById("chat-container");
const inputForm = document.getElementById("input-area");
const userInput = document.getElementById("user-input");
const sendBtn = document.getElementById("send-btn");
const typingIndicator = document.getElementById("typing-indicator");
const openSettings = document.getElementById("open-settings");
const closeSettings = document.getElementById("close-settings");
const saveSettings = document.getElementById("save-settings");
const settingsModal = document.getElementById("settings-modal");
const apiKeyInput = document.getElementById("api-key-input");

const fileInput = document.getElementById("file-input");
const attachBtn = document.getElementById("attach-btn");
const fileList = document.getElementById("file-list");
const quotaWarning = document.getElementById("quota-warning");
const retryTimer = document.getElementById("retry-timer");
const clearChatBtn = document.getElementById("clear-chat");
const roleBtns = document.querySelectorAll(".role-btn");

// Mobile menu toggle
const mobileMenuToggle = document.getElementById("mobile-menu-toggle");
const menuIconOpen = document.getElementById("menu-icon-open");
const menuIconClose = document.getElementById("menu-icon-close");
const sidebar = document.querySelector(".sidebar");

if (mobileMenuToggle) {
  mobileMenuToggle.addEventListener("click", () => {
    const isOpen = sidebar.classList.toggle("sidebar--open");
    menuIconOpen.style.display = isOpen ? "none" : "block";
    menuIconClose.style.display = isOpen ? "block" : "none";
  });
}

// Close mobile menu when a role is selected
document.querySelectorAll(".role-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    if (window.innerWidth <= 640) {
      sidebar.classList.remove("sidebar--open");
      menuIconOpen.style.display = "block";
      menuIconClose.style.display = "none";
    }
  });
});


// State
let apiKey = import.meta.env.VITE_GEMINI_API_KEY || localStorage.getItem("gemini_api_key") || "";
let chatHistories = JSON.parse(localStorage.getItem("gemini_chat_histories")) || { "HR": [], "Teknisi": [], "Finance": [] };
let genAI = null;
let model = null;
let knowledgeContext = "";
let uploadedFiles = JSON.parse(localStorage.getItem("gemini_knowledge_files")) || [];
let currentRole = localStorage.getItem("gemini_current_role") || "HR";

// Initial Load
if (apiKey) {
  initGemini();
  apiKeyInput.value = apiKey;
}

// Set initial active role UI
roleBtns.forEach(btn => {
  if (btn.dataset.role === currentRole) {
    btn.classList.add("active");
  } else {
    btn.classList.remove("active");
  }
});

function renderChatHistory() {
  chatContainer.innerHTML = "";
  const history = chatHistories[currentRole] || [];
  
  if (history.length === 0) {
    appendMessage("bot", `Halo! Saya adalah **Knowledge Twin** untuk role **${currentRole}**. Silakan unggah PDF atau tanya sesuatu.`);
  } else {
    history.forEach(entry => {
      const role = entry.role === "user" ? "user" : "bot";
      const text = entry.parts[0].text;
      const cleanText = role === "user" ? text.split("PERTANYAAN USER: ").pop() : text;
      appendMessage(role, cleanText);
    });
  }
}

// Render initial history
renderChatHistory();

if (uploadedFiles.length > 0) {
  rebuildKnowledgeContext();
  updateFileListUI();
}

function initGemini() {
  try {
    genAI = new GoogleGenerativeAI(apiKey);
    const systemInstructions = {
      "HR": "Anda adalah 'Knowledge Twin' untuk peran HR (Human Resources). Tugas Anda adalah memberikan jawaban berdasarkan dokumen HR yang diunggah. Berbicaralah dengan nada profesional, ramah, dan solutif sebagaimana layaknya seorang staf HR.",
      "Teknisi": "Anda adalah 'Knowledge Twin' untuk peran Teknisi. Tugas Anda adalah memberikan jawaban berdasarkan dokumen teknis yang diunggah. Berbicaralah dengan nada teknis, detail, dan praktis sebagaimana layaknya seorang teknisi yang ahli.",
      "Finance": "Anda adalah 'Knowledge Twin' untuk peran Finance (Keuangan). Tugas Anda adalah memberikan jawaban berdasarkan dokumen keuangan (laporan, spreadsheet) yang diunggah. Berbicaralah dengan nada akurat, teliti, dan formal sebagaimana layaknya seorang analis keuangan."
    };

    model = genAI.getGenerativeModel({ 
      model: "gemini-3-flash-preview",
      systemInstruction: systemInstructions[currentRole] || systemInstructions["HR"]
    });
  } catch (error) {
    console.error("Failed to initialize Gemini:", error);
  }
}

function saveToStorage() {
  localStorage.setItem("gemini_chat_histories", JSON.stringify(chatHistories));
  localStorage.setItem("gemini_knowledge_files", JSON.stringify(uploadedFiles));
  localStorage.setItem("gemini_current_role", currentRole);
}

// PDF Reader Logic
async function readPdf(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
  let fullText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map(item => item.str).join(" ");
    fullText += pageText + "\n";
  }
  return fullText;
}

// Excel/CSV Reader Logic
async function readExcel(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  let fullText = "";

  workbook.SheetNames.forEach(sheetName => {
    fullText += `SHEET: ${sheetName}\n`;
    const worksheet = workbook.Sheets[sheetName];
    // Convert to CSV format as it's tokens-efficient and structured
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    fullText += csv + "\n\n";
  });
  
  return fullText;
}

// UI Helpers
function appendMessage(role, text) {
  const msgDiv = document.createElement("div");
  msgDiv.className = `message ${role}-message`;
  
  if (role === "bot") {
    msgDiv.innerHTML = marked.parse(text);
  } else {
    // Simple sanitization for user messages
    msgDiv.textContent = text;
  }
  
  chatContainer.appendChild(msgDiv);
  chatContainer.scrollTop = chatContainer.scrollHeight;
}

function toggleTyping(show) {
  typingIndicator.style.display = show ? "block" : "none";
  chatContainer.scrollTop = chatContainer.scrollHeight;
}

function updateFileListUI() {
  fileList.innerHTML = "";
  const roleFiles = uploadedFiles.filter(f => f.role === currentRole || (!f.role && currentRole === "HR"));
  
  roleFiles.forEach((file) => {
    // Find absolute index in uploadedFiles for deletion
    const absIndex = uploadedFiles.findIndex(f => f.name === file.name && (f.role === file.role));
    
    const item = document.createElement("div");
    item.className = "file-item";
    item.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>
      <span>${file.name}</span>
      <button class="remove-file" data-index="${absIndex}" style="background:none; border:none; color:#ef4444; cursor:pointer; font-weight:bold; margin-left:4px;">×</button>
    `;
    fileList.appendChild(item);
  });

  document.querySelectorAll(".remove-file").forEach(btn => {
    btn.onclick = (e) => {
      const idx = parseInt(e.target.dataset.index);
      uploadedFiles.splice(idx, 1);
      rebuildKnowledgeContext();
      updateFileListUI();
      saveToStorage();
    };
  });
}

function rebuildKnowledgeContext() {
  const roleFiles = uploadedFiles.filter(f => f.role === currentRole || (!f.role && currentRole === "HR"));
  knowledgeContext = roleFiles.map(f => f.content).join("\n\n---\n\n");
}

function showQuotaWarning(retrySeconds) {
  quotaWarning.style.display = "flex";
  let timeLeft = retrySeconds || 60;
  
  const timer = setInterval(() => {
    const mins = Math.floor(timeLeft / 60);
    const secs = timeLeft % 60;
    retryTimer.innerText = `${mins}:${secs.toString().padStart(2, '0')}`;
    
    if (timeLeft <= 0) {
      clearInterval(timer);
      quotaWarning.style.display = "none";
    }
    timeLeft--;
  }, 1000);
}

// Event Handlers
openSettings.addEventListener("click", () => {
  settingsModal.style.display = "flex";
});

closeSettings.addEventListener("click", () => {
  settingsModal.style.display = "none";
});

saveSettings.addEventListener("click", () => {
  const newKey = apiKeyInput.value.trim();
  if (newKey) {
    apiKey = newKey;
    localStorage.setItem("gemini_api_key", apiKey);
    initGemini();
    settingsModal.style.display = "none";
    appendMessage("bot", "Pengaturan disimpan!");
  }
});

clearChatBtn.addEventListener("click", () => {
  if (confirm(`Hapus riwayat chat untuk role ${currentRole}?`)) {
    chatHistories[currentRole] = [];
    localStorage.setItem("gemini_chat_histories", JSON.stringify(chatHistories));
    renderChatHistory();
    settingsModal.style.display = "none";
  }
});

attachBtn.addEventListener("click", () => {
  fileInput.click();
});

fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  toggleTyping(true);
  try {
    let text = "";
    const extension = file.name.split('.').pop().toLowerCase();

    if (extension === 'pdf') {
      text = await readPdf(file);
    } else if (['xlsx', 'xls', 'csv'].includes(extension)) {
      text = await readExcel(file);
    } else {
      throw new Error("Format file tidak didukung.");
    }

    uploadedFiles.push({ 
      name: file.name, 
      content: text, 
      role: currentRole 
    });
    rebuildKnowledgeContext();
    updateFileListUI();
    saveToStorage();
    appendMessage("bot", `Dokumen **${file.name}** berhasil dipelajari untuk role **${currentRole}**.`);
  } catch (error) {
    console.error("File Read Error:", error);
    appendMessage("bot", `Gagal membaca file: ${error.message}`);
  } finally {
    toggleTyping(false);
    fileInput.value = "";
  }
});

// Role Switcher Event Listeners
roleBtns.forEach(btn => {
  btn.addEventListener("click", () => {
    const newRole = btn.dataset.role;
    if (newRole === currentRole) return;

    currentRole = newRole;
    
    // Update UI
    roleBtns.forEach(b => b.classList.toggle("active", b.dataset.role === currentRole));
    
    // Re-initialize Gemini with new system instructions
    initGemini();
    
    // Update file list and context
    rebuildKnowledgeContext();
    updateFileListUI();
    
    // Render the chat history for the new role
    renderChatHistory();
    
    saveToStorage();
  });
});

inputForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const text = userInput.value.trim();
  
  if (!text) return;
  if (!apiKey) {
    settingsModal.style.display = "flex";
    return;
  }

  appendMessage("user", text);
  userInput.value = "";
  sendBtn.disabled = true;
  toggleTyping(true);
  quotaWarning.style.display = "none";

  try {
    const finalPrompt = knowledgeContext 
      ? `PENGETAHUAN DARI DOKUMEN:\n${knowledgeContext}\n\nPERTANYAAN USER: ${text}`
      : text;

    const result = await model.generateContentStream({
      contents: [...(chatHistories[currentRole] || []), { role: "user", parts: [{ text: finalPrompt }] }]
    });

    let botText = "";
    const msgDiv = document.createElement("div");
    msgDiv.className = `message bot-message`;
    chatContainer.appendChild(msgDiv);
    
    for await (const chunk of result.stream) {
      const chunkText = chunk.text();
      botText += chunkText;
      
      // Use marked for real-time rendering
      msgDiv.innerHTML = marked.parse(botText);
      
      chatContainer.scrollTop = chatContainer.scrollHeight;
      toggleTyping(false);
    }
    
    // Update local state and storage
    if (!chatHistories[currentRole]) chatHistories[currentRole] = [];
    chatHistories[currentRole].push({ role: "user", parts: [{ text }] });
    chatHistories[currentRole].push({ role: "model", parts: [{ text: botText }] });
    saveToStorage();

  } catch (error) {
    console.error("Full Error Object:", error);
    const isQuotaError = error.message?.includes("429") || error.status === 429;
    
    if (isQuotaError) {
      showQuotaWarning(60); 
      appendMessage("bot", "⚠️ **Kuota Gratis Habis.** Silakan tunggu sekitar 1 menit.");
    } else {
      appendMessage("bot", "Terjadi kesalahan saat menghubungi Gemini.");
    }
  } finally {
    toggleTyping(false);
    sendBtn.disabled = false;
  }
});

window.addEventListener("click", (e) => {
  if (e.target === settingsModal) {
    settingsModal.style.display = "none";
  }
});
