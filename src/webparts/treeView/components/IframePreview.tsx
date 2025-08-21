import React, { useEffect, useRef, useState } from 'react';
import pnp from 'sp-pnp-js';
import { getTranslations, getUserLanguage } from '../../../utils/getTranslations';

interface IframePreviewProps {
    url: string;
    width?: string; // ex: "100%"
    height?: string; // ex: "600px"
    title?: string;
    useFallback?: boolean; // se true tenta buscar dados via PnP quando iframe for bloqueado
    listTitle?: string; // título da lista (para o fallback)
    filterField?: string; // campo usado no filtro (para o fallback)
    filterValue?: string; // valor do filtro (para o fallback)
}

const containerStyle: React.CSSProperties = {
    border: '1px solid #e1e1e1',
    borderRadius: 6,
    padding: 8,
    boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
    background: 'white',
    maxWidth: '100%'
};

export default function IframePreview(props: IframePreviewProps) {
    const {
        url,
        width = '100%',
        height = '800px',
        useFallback = true,
        listTitle,
        filterField,
        filterValue
    } = props;

    const iframeRef = useRef<HTMLIFrameElement | null>(null);
    const [loaded, setLoaded] = useState(false);
    const [iframeBlocked, setIframeBlocked] = useState(false);
    const [checking, setChecking] = useState(false);

    // Fallback state
    const [items, setItems] = useState<any[] | null>(null);
    const [fallbackError, setFallbackError] = useState<string | null>(null);

    const t = getTranslations();

    // Tenta detectar bloqueio do iframe. Se for bloqueado por X-Frame-Options/CSP, acessar contentDocument vai lançar.
    const onIframeLoad = () => {
        setLoaded(true);
        try {
            // Se for same-origin, conseguimos acessar o document sem erro.
            const doc = iframeRef.current?.contentDocument || iframeRef.current?.contentWindow?.document;
            if (doc) {
                // Se a página estiver vazia (alguns bloqueios resultam em body vazio) tenta marcar como bloqueado
                const bodyText = doc.body?.innerText || '';
                if (bodyText.trim().length === 0) {
                    // Pode ser bloqueio ou página que realmente está vazia
                    // aguarda um pequeno tempo antes de decidir
                    setTimeout(() => {
                        try {
                            const d = iframeRef.current?.contentDocument || iframeRef.current?.contentWindow?.document;
                            const bt = d?.body?.innerText || '';
                            if (bt.trim().length === 0) {
                                setIframeBlocked(true);
                            }
                        } catch (e) {
                            setIframeBlocked(true);
                        }
                    }, 400);
                } else {
                    setIframeBlocked(false);
                }
            }
        } catch (e) {
            // Acesso negado -> cross-origin ou bloqueado explicitamente
            setIframeBlocked(true);
        }
    };

    // timeout para detectar quando onLoad não ocorrer (algumas vezes o browser ignora o load)
    useEffect(() => {
        setLoaded(false);
        setIframeBlocked(false);
        setItems(null);
        setFallbackError(null);

        // const to = setTimeout(() => {
        //     if (!loaded) {
        //         // Se não carregou em 15s, considere bloqueado (vamos permitir recarregar manualmente)
        //         setIframeBlocked(true);
        //     }
        // }, 100);
        // return () => clearTimeout(to);
        // eslint-disable-next-line react-hooks/exhaustive-deps
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
                        // ignora
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

    return (
        <div style={{ ...containerStyle, width }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <div style={{ display: 'flex', gap: 8 }}>
                    <button onClick={() => { setIframeBlocked(false); setLoaded(false); if (iframeRef.current) iframeRef.current.src = url; }} title="Recarregar" aria-label="Recarregar">Recarregar</button>
                    <button onClick={() => window.open(url, '_blank')} title="Abrir em nova aba" aria-label="Abrir em nova aba">Abrir em nova aba</button>
                </div>
            </div>

            {!iframeBlocked && (
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
            )}

            {iframeBlocked && (
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
                                                                <a href={`${it.FileRef}?web=1`} target="_blank" rel="noopener noreferrer">{it.FileLeafRef || it.FileRef}</a>
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
