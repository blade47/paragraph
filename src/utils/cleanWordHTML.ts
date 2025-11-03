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
export default function cleanWordHTML(html: string): string {
  if (!html || html.trim().length === 0) {
    return '';
  }

  let cleaned = html;

  // Step 1: Remove Microsoft Office namespace tags
  // These break HTML parsing and serve no purpose outside Office
  cleaned = cleaned.replace(/<\/?o:[^>]+>/gi, ''); // Office namespace: <o:p>
  cleaned = cleaned.replace(/<\/?w:[^>]+>/gi, ''); // Word namespace: <w:sdt>
  cleaned = cleaned.replace(/<\/?v:[^>]+>/gi, ''); // VML namespace: <v:shape>
  cleaned = cleaned.replace(/<\/?st1:[^>]+>/gi, ''); // Smart tags: <st1:place>
  cleaned = cleaned.replace(/<\/?m:[^>]+>/gi, ''); // Math namespace: <m:oMath>

  // Step 2: Remove conditional comments (IE-specific, break modern browsers)
  // Example: <!--[if gte mso 9]><xml>...</xml><![endif]-->
  cleaned = cleaned.replace(/<!--\[if[^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, '');
  cleaned = cleaned.replace(/<!--\[if[^\]]*\]>[\s\S]*?-->/gi, '');

  // Step 3: Remove ALL inline styles (colors, fonts, sizes, etc.)
  // We don't support custom styling in any editor
  cleaned = cleaned.replace(/\s*style="[^"]*"/gi, '');
  cleaned = cleaned.replace(/\s*style='[^']*'/gi, '');

  // Step 4: Remove ALL classes
  // Word classes (MsoNormal, etc.) and any other classes are not needed
  cleaned = cleaned.replace(/\s*class="[^"]*"/gi, '');
  cleaned = cleaned.replace(/\s*class='[^']*'/gi, '');

  // Step 5: Remove <font> tags and face attributes (legacy font styling)
  cleaned = cleaned.replace(/<font[^>]*>/gi, '');
  cleaned = cleaned.replace(/<\/font>/gi, '');
  cleaned = cleaned.replace(/\s*face="[^"]*"/gi, '');

  // Step 6: Remove other Word-specific attributes
  cleaned = cleaned.replace(/\s*lang="[^"]*"/gi, '');
  cleaned = cleaned.replace(/\s*xml:lang="[^"]*"/gi, '');

  // Step 7: Remove all other attributes EXCEPT href on <a> tags
  // This regex preserves href but removes everything else
  cleaned = cleaned.replace(/<(?!a\s)(\w+)[^>]*?(?<!href="[^"]*")>/gi, '<$1>');
  // Clean up <a> tags to only keep href
  cleaned = cleaned.replace(/<a\s+[^>]*?(href="[^"]*")[^>]*?>/gi, '<a $1>');

  // Step 8: Normalize typography (smart quotes to straight quotes)
  cleaned = cleaned.replace(/[""]/g, '"'); // Smart double quotes → "
  cleaned = cleaned.replace(/['']/g, "'"); // Smart single quotes → '
  cleaned = cleaned.replace(/…/g, '...'); // Ellipsis → ...
  cleaned = cleaned.replace(/–/g, '-'); // En dash → -
  cleaned = cleaned.replace(/—/g, '--'); // Em dash → --

  // Step 9: Normalize whitespace
  cleaned = cleaned.replace(/&nbsp;/g, ' '); // &nbsp; → space
  cleaned = cleaned.replace(/&#160;/g, ' '); // &#160; → space
  cleaned = cleaned.replace(/\u00A0/g, ' '); // Non-breaking space → space
  cleaned = cleaned.replace(/\u202F/g, ' '); // Narrow non-breaking space → space
  cleaned = cleaned.replace(/\uFEFF/g, ''); // Zero-width space → remove

  // Collapse multiple spaces (but not inside tags)
  cleaned = cleaned.replace(/([^>])\s{2,}([^<])/g, '$1 $2');

  // Step 10: Remove empty elements created by cleaning
  // <span>, <font>, <div> that became empty
  cleaned = cleaned.replace(/<span[^>]*>\s*<\/span>/gi, '');
  cleaned = cleaned.replace(/<span[^>]*><\/span>/gi, '');
  cleaned = cleaned.replace(/<div[^>]*>\s*<\/div>/gi, '');
  cleaned = cleaned.replace(/<div[^>]*><\/div>/gi, '');

  // Step 11: Remove unsupported tags while preserving content
  // Keep only: p, br, b, strong, i, em, u, mark, a
  // Remove: span, div, h1-h6, ul, ol, li, etc. (content stays)
  const supportedTags = ['p', 'br', 'b', 'strong', 'i', 'em', 'u', 'mark', 'a'];
  const tagRegex = /<\/?(\w+)[^>]*>/gi;
  cleaned = cleaned.replace(tagRegex, (match, tagName) => {
    if (supportedTags.includes(tagName.toLowerCase())) {
      return match; // Keep supported tags
    }
    return ''; // Remove unsupported tags
  });

  // Step 12: Final cleanup - remove completely empty elements
  let previousLength;
  do {
    previousLength = cleaned.length;
    cleaned = cleaned.replace(/<(\w+)[^>]*>\s*<\/\1>/gi, '');
  } while (cleaned.length < previousLength);

  // Step 13: Normalize tag spacing and trim
  cleaned = cleaned.replace(/<(\w+)\s+>/g, '<$1>');
  cleaned = cleaned.replace(/\s+>/g, '>');

  return cleaned.trim();
}

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
export function isWordHTML(html: string): boolean {
  if (!html || html.trim().length === 0) {
    return false;
  }

  // Check for common Word indicators
  const wordIndicators = [
    /class="?Mso/i, // MsoNormal, MsoTitle, etc.
    /<o:/i, // Office namespace tags
    /<w:/i, // Word namespace tags
    /<v:/i, // VML namespace tags
    /xmlns:o=/i, // Office XML namespace declaration
    /xmlns:w=/i, // Word XML namespace declaration
    /urn:schemas-microsoft-com:office/i, // Microsoft Office schema
  ];

  return wordIndicators.some((indicator) => indicator.test(html));
}
