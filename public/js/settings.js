const KEY = "claudeMail.settings.v1";

export const MODELS = [
  { id: "claude-haiku-5-5", label: "Haiku 5.5 (fastest, lowest cost)" },
  { id: "claude-sonnet-5-5", label: "Sonnet 5.5 (balanced)" },
  { id: "claude-opus-5-5", label: "Opus 5.5 (most capable)" },
];

const DEFAULTS = { apiKey: "", model: MODELS[0].id, name: "", clientId: "", tenant: "common" };
let memory = null;

export function loadSettings() {
  if (!memory) {
    let stored = {};
    try {
      stored = JSON.parse(localStorage.getItem(KEY) || "{}");
    } catch {
      /* storage unavailable or corrupt: fall back to defaults */
    }
    memory = { ...DEFAULTS, ...stored };
  }
  return { ...memory };
}

// Returns false when the browser blocked saving to disk (settings then last only for this session).
export function saveSettings(patch) {
  memory = { ...loadSettings(), ...patch };
  try {
    localStorage.setItem(KEY, JSON.stringify(memory));
    return true;
  } catch {
    return false;
  }
}
