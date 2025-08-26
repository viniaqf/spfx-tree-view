// src/utils/localCssInjector.ts
import { SPComponentLoader } from "@microsoft/sp-loader";
import { loadStyles } from "@microsoft/load-themed-styles";

//Injeta css pelo arquivo carregado dentro da biblioteca de estilos
export function injectCssOnce(
  cssPath: string,
  key: string = "treeview_css_injected"
): void {
  if (!cssPath) return;

  const w = window as any;
  w.__cssFlags = w.__cssFlags || {};
  if (w.__cssFlags[key]) return;

  let finalUrl = cssPath;
  const isAbsolute = /^https?:\/\//i.test(cssPath);
  if (!isAbsolute) {
    const webRel =
      (w._spPageContextInfo && w._spPageContextInfo.webServerRelativeUrl) || "";
    if (cssPath.startsWith("/")) {
      finalUrl = cssPath;
    } else {
      finalUrl = `${webRel.replace(/\/$/, "")}/${cssPath}`;
    }
    const origin = location.origin.replace(/\/$/, "");
    finalUrl = `${origin}${finalUrl.startsWith("/") ? "" : "/"}${finalUrl}`;
  }

  const already = Array.from(
    document.head.querySelectorAll('link[rel="stylesheet"]')
  ).some((l) => (l as HTMLLinkElement).href === finalUrl);
  if (already) {
    w.__cssFlags[key] = true;
    return;
  }

  SPComponentLoader.loadCss(finalUrl);
  w.__cssFlags[key] = true;
}

//Injeta css por código, dentro da pasta style/spfx_style.ts
export function injectCssStringOnce(
  cssText: string,
  key: string = "treeview_css_string_injected"
): void {
  if (!cssText) return;

  const w = window as any;
  w.__cssFlags = w.__cssFlags || {};
  if (w.__cssFlags[key]) return;

  // loadStyles insere um <style> no <head>
  loadStyles(cssText);
  w.__cssFlags[key] = true;
}
