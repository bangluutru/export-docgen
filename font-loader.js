/**
 * Font Loader — Lazy-loads Noto Sans JP for PDF export
 * Font is loaded on-demand (not at page startup) to keep initial load fast.
 * 
 * Strategy:
 * 1. Try to load from local file (NotoSansJP-Regular.ttf)
 * 2. If local fails, load from Google Fonts CDN
 */
const FontLoader = (() => {
    let fontDataCache = null; // cached base64 string
    let loadingPromise = null;

    // CDN fallback URL for the font
    const FONT_URLS = [
        'NotoSansJP-Regular.ttf',  // local
        'https://cdn.jsdelivr.net/fontsource/fonts/noto-sans-jp@latest/japanese-400-normal.woff2',
    ];

    /**
     * Load the Noto Sans JP font and return its base64 string.
     * First call fetches from server; subsequent calls return cache.
     * @returns {Promise<string>} base64-encoded font data
     */
    async function loadNotoSansJP() {
        if (fontDataCache) return fontDataCache;
        if (loadingPromise) return loadingPromise;

        loadingPromise = (async () => {
            try {
                // Try local first, then CDN fallback
                let buffer = null;
                for (const url of FONT_URLS) {
                    try {
                        const resp = await fetch(url);
                        if (resp.ok) {
                            buffer = await resp.arrayBuffer();
                            console.log(`[FontLoader] Loaded font from: ${url} (${(buffer.byteLength / 1024).toFixed(0)} KB)`);
                            break;
                        }
                    } catch (e) {
                        console.warn(`[FontLoader] Failed to load from ${url}:`, e.message);
                    }
                }

                if (!buffer) {
                    throw new Error('Could not load Noto Sans JP font from any source');
                }

                // Convert ArrayBuffer to base64
                const bytes = new Uint8Array(buffer);
                let binary = '';
                const chunkSize = 8192;
                for (let i = 0; i < bytes.length; i += chunkSize) {
                    const chunk = bytes.subarray(i, i + chunkSize);
                    binary += String.fromCharCode.apply(null, chunk);
                }
                fontDataCache = btoa(binary);
                return fontDataCache;
            } catch (err) {
                loadingPromise = null; // allow retry
                throw err;
            }
        })();

        return loadingPromise;
    }

    /**
     * Register the Noto Sans JP font with a jsPDF document instance.
     * @param {jsPDF} doc — jsPDF document
     */
    async function registerFont(doc) {
        const base64Data = await loadNotoSansJP();

        // Register Regular weight
        doc.addFileToVFS('NotoSansJP-Regular.ttf', base64Data);
        doc.addFont('NotoSansJP-Regular.ttf', 'NotoSansJP', 'normal');

        // Register Bold (same variable font — jsPDF will use it for bold style)
        doc.addFileToVFS('NotoSansJP-Bold.ttf', base64Data);
        doc.addFont('NotoSansJP-Bold.ttf', 'NotoSansJP', 'bold');

        doc.setFont('NotoSansJP', 'normal');
    }

    /**
     * Check if font is already cached (loaded).
     * @returns {boolean}
     */
    function isLoaded() {
        return fontDataCache !== null;
    }

    return { loadNotoSansJP, registerFont, isLoaded };
})();
