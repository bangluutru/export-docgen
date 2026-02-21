/**
 * Font Loader — Lazy-loads Noto Sans for PDF export (Vietnamese + Latin support)
 * Font is loaded on-demand (not at page startup) to keep initial load fast.
 * 
 * Strategy:
 * 1. Try to load from local file (NotoSans-Regular.ttf)
 * 2. If local fails, load from CDN
 */
const FontLoader = (() => {
    let fontDataCache = null; // cached base64 string
    let loadingPromise = null;

    // CDN fallback URL for the font
    const FONT_URLS = [
        'NotoSans-Regular.ttf',  // local
        'https://cdn.jsdelivr.net/fontsource/fonts/noto-sans@latest/vietnamese-400-normal.ttf',
    ];

    /**
     * Load the Noto Sans font and return its base64 string.
     * First call fetches from server; subsequent calls return cache.
     * @returns {Promise<string>} base64-encoded font data
     */
    async function loadNotoSans() {
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
                            // Validate that we got a real font file, not HTML (SPA fallback)
                            const contentType = resp.headers.get('content-type') || '';
                            if (contentType.includes('text/html')) {
                                console.warn(`[FontLoader] ${url} returned HTML (not a font), skipping`);
                                continue;
                            }
                            buffer = await resp.arrayBuffer();
                            // Extra check: valid TTF/OTF starts with specific magic bytes
                            const header = new Uint8Array(buffer.slice(0, 4));
                            const magic = String.fromCharCode(...header);
                            if (buffer.byteLength < 1000 || (magic !== '\x00\x01\x00\x00' && magic !== 'OTTO' && magic !== 'true' && magic !== 'typ1')) {
                                console.warn(`[FontLoader] ${url} is not a valid font file (${buffer.byteLength} bytes), skipping`);
                                buffer = null;
                                continue;
                            }
                            console.log(`[FontLoader] Loaded font from: ${url} (${(buffer.byteLength / 1024).toFixed(0)} KB)`);
                            break;
                        }
                    } catch (e) {
                        console.warn(`[FontLoader] Failed to load from ${url}:`, e.message);
                    }
                }

                if (!buffer) {
                    throw new Error('Could not load Noto Sans font from any source');
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
        const base64Data = await loadNotoSans();

        // Register Regular weight
        doc.addFileToVFS('NotoSans-Regular.ttf', base64Data);
        doc.addFont('NotoSans-Regular.ttf', 'NotoSans', 'normal');

        // Register Bold (use same font for bold — jsPDF will simulate bold)
        doc.addFileToVFS('NotoSans-Bold.ttf', base64Data);
        doc.addFont('NotoSans-Bold.ttf', 'NotoSans', 'bold');

        doc.setFont('NotoSans', 'normal');
    }

    /**
     * Check if font is already cached (loaded).
     * @returns {boolean}
     */
    function isLoaded() {
        return fontDataCache !== null;
    }

    return { loadNotoSans, registerFont, isLoaded };
})();
