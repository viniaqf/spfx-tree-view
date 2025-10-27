import React, { useEffect, useRef, useState } from 'react';
import pnp from 'sp-pnp-js';
import { getTranslations, getUserLanguage } from '../../../utils/getTranslations';
import { HIDE_SWITCHER_CSS } from '../../../styles/spfx_style';


interface IframePreviewProps {
    url: string;
    width?: string;
    height?: string;
    title?: string;
    useFallback?: boolean; // se true tenta buscar dados via PnP quando iframe for bloqueado
    listTitle?: string; // título da lista (para o fallback)
    filterField?: string; // campo usado no filtro (para o fallback)
    filterValue?: string; // valor do filtro (para o fallback)
    emptyMessage?: string;
    newTitle?: string
}

const containerStyle: React.CSSProperties = {
    border: '1px solid #e1e1e1',
    borderRadius: 6,
    padding: 8,
    boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
    background: 'white',
    maxWidth: '100%'
};

const placeholderStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'white',
    width: '100%',
    height: '100%',
    minHeight: '200px'
};

const placeholderTextStyle: React.CSSProperties = {
    fontSize: 16,
    color: '#444',
    textAlign: 'center',
    lineHeight: 1.4
};


export default function IframePreview(props: IframePreviewProps) {
    const {
        url,
        width = '100%',
        height = '800px',
        useFallback = true,
        listTitle,
        filterField,
        filterValue,
        emptyMessage,
        newTitle,
    } = props;

    const iframeRef = useRef<HTMLIFrameElement | null>(null);
    const observerRef = useRef<MutationObserver | null>(null);
    const [loaded, setLoaded] = useState(false);
    const [iframeBlocked, setIframeBlocked] = useState(false);
    const [checking, setChecking] = useState(false);
    const [items, setItems] = useState<any[] | null>(null);
    const [fallbackError, setFallbackError] = useState<string | null>(null);


    const t = getTranslations();

    function setHeaderTitleText(doc: Document, text: string): boolean {
        const el = doc.querySelector(
            'div[data-automationid="headerTitleButton"], button[data-automationid="headerTitleButton"]'
        ) as HTMLElement | null;
        if (!el) return false;

        el.textContent = text;
        el.setAttribute('title', text);
        el.style.display = 'flex';
        el.style.alignItems = 'center';
        return true;
    }



    // Tenta detectar bloqueio do iframe. Se for bloqueado por X-Frame-Options/CSP, acessar contentDocument vai lançar.
    const onIframeLoad = () => {
        setLoaded(true);
        try {
            const iframe = iframeRef.current;
            const doc = iframe?.contentDocument || iframe?.contentWindow?.document;

            if (doc) {
                // INJETAR <style> NO HEAD DO IFRAME (apenas 1 vez)
                if (!doc.head.querySelector('style[data-injected-by="treeview"]')) {
                    const style = doc.createElement('style');
                    style.setAttribute('data-injected-by', 'treeview');
                    style.textContent = HIDE_SWITCHER_CSS;
                    doc.head.appendChild(style);
                }

                //Força ocultar imediatamente (caso haja estilos inline ou re-render)
                const hideNow = () => {
                    doc
                        .querySelectorAll(
                            'button[data-automationid="librariesDropdownButton"], button[class^="librariesDropdown_"]'
                        )
                        .forEach((el) => {
                            (el as HTMLElement).style.setProperty('display', 'none', 'important');
                        });
                };
                hideNow();

                // OBSERVA re-renderizações dentro do iframe e reaplica o hide
                if (observerRef.current) observerRef.current.disconnect();
                observerRef.current = new MutationObserver(() => hideNow());
                observerRef.current.observe(doc.body, { childList: true, subtree: true });

                //  Sinaliza que o iframe NÃO está bloqueado
                setIframeBlocked(false);

                if (newTitle && newTitle.trim()) {
                    const injectNow = () => setHeaderTitleText(doc, newTitle!.trim());

                    if (!injectNow()) {
                        // se o header ainda não está no DOM, observa até aparecer
                        const mo = new MutationObserver(() => {
                            if (injectNow()) {
                                mo.disconnect();
                            }
                        });
                        mo.observe(doc.body, { childList: true, subtree: true });
                    }
                }


                const bodyText = doc.body?.innerText || '';
                if (bodyText.trim().length === 0) {
                    setTimeout(() => {
                        try {
                            const d = iframeRef.current?.contentDocument || iframeRef.current?.contentWindow?.document;
                            const bt = d?.body?.innerText || '';
                            if (bt.trim().length === 0) {
                                setIframeBlocked(true);
                            }
                        } catch {
                            setIframeBlocked(true);
                        }
                    }, 400);
                }
            }
        } catch (e) {
            // Cross-origin ou bloqueio total: não dá para acessar o DOM do iframe
            setIframeBlocked(true);
        }

    };




    // timeout para detectar quando onLoad não ocorrer (algumas vezes o browser ignora o load)
    useEffect(() => {
        observerRef.current?.disconnect();
        setLoaded(false);
        setIframeBlocked(false);
        setItems(null);
        setFallbackError(null);
    }, [url]);

    // Se detectou bloqueio e o usuário quer fallback, buscamos dados via PnP
    useEffect(() => {
        if (!iframeBlocked) return;
        if (!useFallback) return;
        if (!listTitle || !filterField || !filterValue) return;

        let canceled = false;
        const fetchItems = async () => {
            setChecking(true);
            setFallbackError(null);
            try {
                // Primeira tentativa: filtrar por valor textual (útil para Choice/Texto)
                const filterText = `${filterField} eq '${filterValue.replace(/'/g, "\\'")}'`;
                let q = pnp.sp.web.lists.getByTitle(listTitle).items.filter(filterText).select('ID', 'Title', 'FileRef', 'FileLeafRef').top(500);
                let result = await q.get();

                // Se não encontrou e o valor parece numérico, tenta LookupId
                if ((!result || result.length === 0) && !isNaN(Number(filterValue))) {
                    try {
                        const lookupFilter = `${filterField}Id eq ${Number(filterValue)}`;
                        result = await pnp.sp.web.lists.getByTitle(listTitle).items.filter(lookupFilter).select('ID', 'Title', 'FileRef', 'FileLeafRef').top(500).get();
                    } catch (e) {

                    }
                }

                if (!canceled) {
                    setItems(result || []);
                }
            } catch (error) {
                if (!canceled) {
                    setFallbackError(String(error));
                }
            } finally {
                if (!canceled) setChecking(false);
            }
        };

        fetchItems();
        return () => { canceled = true; };
    }, [iframeBlocked, useFallback, listTitle, filterField, filterValue]);

    // IframePreview.tsx — SUBSTITUIR O return INTEIRO por este
    return (
        <div style={{ ...containerStyle, width }}>
            {!url ? (
                // Caso 1: sem URL → placeholder inicial
                <div style={{ ...placeholderStyle, height }}>
                    <div style={placeholderTextStyle}>
                        {emptyMessage || "Por favor, selecione um item para visualizar."}
                    </div>
                </div>
            ) : !iframeBlocked ? (
                // Caso 2: com URL e iframe OK
                <div style={{ width: '100%', height }}>
                    <iframe
                        ref={iframeRef}
                        src={url}
                        width="100%"
                        height="100%"
                        style={{ border: 'none', minHeight: height }}
                        onLoad={onIframeLoad}

                    />
                </div>
            ) : (
                // Caso 3: com URL, mas bloqueado → fallback
                <div>
                    <div style={{ marginBottom: 8, color: '#9a1b0d' }}>
                        {t.iframe_load_error}
                    </div>

                    {useFallback ? (
                        <div>
                            <div style={{ marginBottom: 8 }}></div>
                            {checking}
                            {items && (
                                <div>
                                    {items.length === 0 && <div>{t.no_items_found_filter}</div>}
                                    {items.length > 0 && (
                                        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                                            <thead>
                                                <tr>
                                                    <th style={{ textAlign: 'left', padding: 6, borderBottom: '1px solid #eee' }}>#</th>
                                                    <th style={{ textAlign: 'left', padding: 6, borderBottom: '1px solid #eee' }}>Título</th>
                                                    <th style={{ textAlign: 'left', padding: 6, borderBottom: '1px solid #eee' }}>Arquivo</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {items.map((it: any) => (
                                                    <tr key={it.ID}>
                                                        <td style={{ padding: 6, borderBottom: '1px solid #f5f5f5' }}>{it.ID}</td>
                                                        <td style={{ padding: 6, borderBottom: '1px solid #f5f5f5' }}>{it.Title}</td>
                                                        <td style={{ padding: 6, borderBottom: '1px solid #f5f5f5' }}>
                                                            {it.FileRef ? (
                                                                <a href={`${it.FileRef}?web=1`} target="_blank" rel="noopener noreferrer">
                                                                    {it.FileLeafRef || it.FileRef}
                                                                </a>
                                                            ) : ('-')}
                                                        </td>
                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    )}
                                </div>
                            )}
                        </div>
                    ) : (
                        <div></div>
                    )}
                </div>
            )}
        </div>
    );
}
