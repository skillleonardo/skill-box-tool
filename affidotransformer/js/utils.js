/**
 * Pulisce e formatta il testo estratto.
 * @param {string} text - Il testo grezzo.
 * @returns {string} - Il testo pulito e formattato.
 */
export function cleanAndFormatText(text) {
    if (!text) return "";
    let clean = text.replace(/\s+/g, ' ').trim();
    clean = clean.replace(/([.,:;?!])(?=\S)/g, '$1 ');
    clean = clean.replace(/(?:^|[.!?]\s+)([a-z])/g, function(match) {
        return match.toUpperCase();
    });
    return clean;
}
