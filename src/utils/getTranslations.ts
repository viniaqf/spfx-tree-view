import pt from "../webparts/treeView/i18n/pt.json"; 
import es from '../webparts/treeView/i18n/es.json';

export type Translations = typeof pt;

export function getUserLanguage(): string {
  const languages = navigator.languages?.map(l => l.toLowerCase()) || [];
  console.log(`Idiomas preferidos: ${languages.join(', ')}`);

  if (languages.some(l => l.startsWith('es'))) return 'es';

  return 'pt';
}

console.log(`Idioma do usuário: ${getUserLanguage()}`);

export function getTranslations(): Translations {
  const lang = getUserLanguage();
  switch (lang) {
    case 'es': return es;
    default: return pt;
  }
}

console.log(" Traduções carregadas:", getTranslations());