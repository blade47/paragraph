/**
 * Cleans HTML pasted from Microsoft Word documents
 *
 * Removes ALL formatting except the basic inline tools supported by EditorJS:
 * - Bold (<b>, <strong>)
 * - Italic (<i>, <em>)
 * - Underline (<u>)
 * - Mark/Highlight (<mark>)
 * - Links (<a href="">)
 *
 * Strips out:
 * - All inline styles (color, font-family, font-size, etc.)
 * - All Office namespace tags (<o:p>, <w:sdt>, etc.)
 * - All classes and attributes (except href on links)
 * - All unsupported tags
 *
 * Use this for ALL content - abstract generation, marketing, custom pages, etc.
 *
 * @param {string} html - Raw HTML from Word paste event
 * @returns {string} - Cleaned HTML with only supported formatting
 *
 * @example
 * const dirtyHTML = '<p class="MsoNormal" style="color:red;font-family:Arial"><o:p><b>Bold</b> text</o:p></p>';
 * const cleaned = cleanWordHTML(dirtyHTML);
 * // Returns: '<p><b>Bold</b> text</p>'
 */
export default function cleanWordHTML(html: string): string;
/**
 * Detects if HTML content originated from Microsoft Word
 *
 * @param {string} html - HTML content to analyze
 * @returns {boolean} - True if content appears to be from Word
 *
 * @example
 * isWordHTML('<p class="MsoNormal">Text</p>'); // Returns: true
 * isWordHTML('<p>Normal text</p>'); // Returns: false
 */
export declare function isWordHTML(html: string): boolean;
