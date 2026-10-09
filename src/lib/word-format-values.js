/** Native highlight strings, including Word's standard palette and no highlight. */
const HIGHLIGHTS = {
    yellow: '#FFFF00', lime: '#00FF00', turquoise: '#00FFFF', pink: '#FF00FF',
    blue: '#0000FF', red: '#FF0000', darkblue: '#000080', teal: '#008080',
    green: '#008000', purple: '#800080', darkred: '#800000', olive: '#808000',
    gray: '#808080', lightgray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
    darkyellow: '#808000', darkgray: '#808080', cyan: '#00FFFF', magenta: '#FF00FF',
    darkgreen: '#008000', darkcyan: '#008080', darkmagenta: '#800080',
};

export function comparableHighlightColor(value) {
    if (value === null) return null;
    if (typeof value !== 'string') return value;
    const name = value.toLowerCase().replace(/[\s_-]+/g, '');
    if (name === 'none') return null;
    return HIGHLIGHTS[name] || value.toUpperCase();
}

export function nativeHighlightColor(value) {
    const color = comparableHighlightColor(value);
    if (color === null || typeof color === 'string' && /^#[\da-f]{6}$/i.test(color)) return color;
    throw new Error(`Unknown formatting highlightColor: ${value}. Use a standard color name, #RRGGBB, or none.`);
}
