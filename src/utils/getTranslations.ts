import pt from "../webparts/treeView/i18n/pt.json";
import es from "../webparts/treeView/i18n/es.json";

const TRANSLATIONS = { pt, es } as const;
type Lang = keyof typeof TRANSLATIONS; // 'pt' | 'es'
export type Translations = typeof pt;

function isLang(code: string): code is Lang {
  return code === "pt" || code === "es";
}

export function getUserLanguage(): Lang {
  const candidates = (
    navigator.languages?.length ? navigator.languages : [navigator.language]
  )
    .filter(Boolean)
    .map((l) => l.toLowerCase());

  // respeita a ordem do navegador
  for (const l of candidates) {
    const base = l.split("-")[0]; // 'pt-br' -> 'pt'
    if (isLang(base)) return base;
  }

  return "pt"; // fallback
}

export function getTranslations(): Translations {
  return TRANSLATIONS[getUserLanguage()];
}

// logs úteis
console.log("Idiomas preferidos:", navigator.languages);
console.log("Idioma escolhido:", getUserLanguage());
console.log("Traduções carregadas:", getTranslations());
