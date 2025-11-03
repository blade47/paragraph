(function(){"use strict";try{if(typeof document<"u"){var e=document.createElement("style");e.appendChild(document.createTextNode(".ce-paragraph{line-height:1.6em;outline:none}.ce-paragraph--right{text-align:right}.ce-paragraph--center{text-align:center}.ce-paragraph--left{text-align:left}.ce-paragraph--justify{text-align:justify}.ce-block:only-of-type .ce-paragraph[data-placeholder-active]:empty:before,.ce-block:only-of-type .ce-paragraph[data-placeholder-active][data-empty=true]:before{content:attr(data-placeholder-active)}.ce-paragraph p:first-of-type{margin-top:0}.ce-paragraph p:last-of-type{margin-bottom:0}")),document.head.appendChild(e)}}catch(a){console.error("vite-plugin-css-injected-by-js",a)}})();
const c = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24"><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M18 7L6 7"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M18 17H6"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M16 12L8 12"/></svg>', g = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24"><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M18 7L6 7"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M18 17H6"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M18 12L6 12"/></svg>', h = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24"><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M17 7L5 7"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M17 17H5"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M13 12L5 12"/></svg>', p = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24"><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M19 7L7 7"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M19 17H7"/><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M19 12L11 12"/></svg>', d = '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" fill="none" viewBox="0 0 24 24"><path stroke="currentColor" stroke-linecap="round" stroke-width="2" d="M8 9V7.2C8 7.08954 8.08954 7 8.2 7L12 7M16 9V7.2C16 7.08954 15.9105 7 15.8 7L12 7M12 7L12 17M12 17H10M12 17H14"/></svg>';
function u(r) {
  const e = document.createElement("div");
  e.innerHTML = r.trim();
  const t = document.createDocumentFragment();
  return t.append(...Array.from(e.childNodes)), t;
}
function m(r) {
  if (!r || r.trim().length === 0)
    return "";
  let e = r;
  e = e.replace(/<\/?o:[^>]+>/gi, ""), e = e.replace(/<\/?w:[^>]+>/gi, ""), e = e.replace(/<\/?v:[^>]+>/gi, ""), e = e.replace(/<\/?st1:[^>]+>/gi, ""), e = e.replace(/<\/?m:[^>]+>/gi, ""), e = e.replace(/<!--\[if[^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, ""), e = e.replace(/<!--\[if[^\]]*\]>[\s\S]*?-->/gi, ""), e = e.replace(/\s*style="[^"]*"/gi, ""), e = e.replace(/\s*style='[^']*'/gi, ""), e = e.replace(/\s*class="[^"]*"/gi, ""), e = e.replace(/\s*class='[^']*'/gi, ""), e = e.replace(/<font[^>]*>/gi, ""), e = e.replace(/<\/font>/gi, ""), e = e.replace(/\s*face="[^"]*"/gi, ""), e = e.replace(/\s*lang="[^"]*"/gi, ""), e = e.replace(/\s*xml:lang="[^"]*"/gi, ""), e = e.replace(new RegExp('<(?!a\\s)(\\w+)[^>]*?(?<!href="[^"]*")>', "gi"), "<$1>"), e = e.replace(/<a\s+[^>]*?(href="[^"]*")[^>]*?>/gi, "<a $1>"), e = e.replace(/[""]/g, '"'), e = e.replace(/['']/g, "'"), e = e.replace(/…/g, "..."), e = e.replace(/–/g, "-"), e = e.replace(/—/g, "--"), e = e.replace(/&nbsp;/g, " "), e = e.replace(/&#160;/g, " "), e = e.replace(/\u00A0/g, " "), e = e.replace(/\u202F/g, " "), e = e.replace(/\uFEFF/g, ""), e = e.replace(/([^>])\s{2,}([^<])/g, "$1 $2"), e = e.replace(/<span[^>]*>\s*<\/span>/gi, ""), e = e.replace(/<span[^>]*><\/span>/gi, ""), e = e.replace(/<div[^>]*>\s*<\/div>/gi, ""), e = e.replace(/<div[^>]*><\/div>/gi, "");
  const t = ["p", "br", "b", "strong", "i", "em", "u", "mark", "a"], n = /<\/?(\w+)[^>]*>/gi;
  e = e.replace(n, (l, o) => t.includes(o.toLowerCase()) ? l : "");
  let s;
  do
    s = e.length, e = e.replace(/<(\w+)[^>]*>\s*<\/\1>/gi, "");
  while (e.length < s);
  return e = e.replace(/<(\w+)\s+>/g, "<$1>"), e = e.replace(/\s+>/g, ">"), e.trim();
}
function w(r) {
  return !r || r.trim().length === 0 ? !1 : [
    /class="?Mso/i,
    // MsoNormal, MsoTitle, etc.
    /<o:/i,
    // Office namespace tags
    /<w:/i,
    // Word namespace tags
    /<v:/i,
    // VML namespace tags
    /xmlns:o=/i,
    // Office XML namespace declaration
    /xmlns:w=/i,
    // Word XML namespace declaration
    /urn:schemas-microsoft-com:office/i
    // Microsoft Office schema
  ].some((t) => t.test(r));
}
/**
 * Base Paragraph Block for the Editor.js.
 * Represents a regular text block
 *
 * @author CodeX (team@codex.so)
 * @copyright CodeX 2018
 * @license The MIT License (MIT)
 */
var a = /* @__PURE__ */ ((r) => (r.LEFT = "left", r.CENTER = "center", r.RIGHT = "right", r.JUSTIFY = "justify", r))(a || {});
class i {
  /**
   * Default placeholder for Paragraph Tool
   *
   * @returns {string}
   * @class
   */
  static get DEFAULT_PLACEHOLDER() {
    return "";
  }
  /**
   * Allowed paragraph alignments
   *
   * @public
   * @returns {{left: string, center: string, right: string}}
  */
  static get ALIGNMENTS() {
    return {
      left: "left".toString(),
      center: "center".toString(),
      right: "right".toString(),
      justify: "justify".toString()
    };
  }
  /**
   * Default paragraph alignment
   *
   * @public
   * @returns {string}
   */
  static get DEFAULT_ALIGNMENT() {
    return "left";
  }
  /**
   * Render plugin`s main Element and fill it with saved data
   *
   * @param {object} params - constructor params
   * @param {ParagraphData} params.data - previously saved data
   * @param {ParagraphConfig} params.config - user config for Tool
   * @param {object} params.api - editor.js api
   * @param {boolean} readOnly - read only mode flag
   */
  constructor({ data: e, config: t, api: n, readOnly: s }) {
    this.api = n, this.config = {
      placeholder: t.placeholder,
      preserveBlank: t.preserveBlank
    }, this.readOnly = s, this._CSS = {
      block: this.api.styles.block,
      wrapper: "ce-paragraph",
      alignment: {
        left: "ce-paragraph--left",
        center: "ce-paragraph--center",
        right: "ce-paragraph--right",
        justify: "ce-paragraph--justify"
      }
    }, this.readOnly || (this.onKeyUp = this.onKeyUp.bind(this)), this._placeholder = t.placeholder ? t.placeholder : i.DEFAULT_PLACEHOLDER, this._data = {
      text: "",
      alignment: i.DEFAULT_ALIGNMENT
    }, this._element = null, this._preserveBlank = t.preserveBlank ?? !1, this.data = e;
  }
  /**
   * Check if text content is empty and set empty string to inner html.
   * We need this because some browsers (e.g. Safari) insert <br> into empty contenteditanle elements
   *
   * @param {KeyboardEvent} e - key up event
   */
  onKeyUp(e) {
    if (e.code !== "Backspace" && e.code !== "Delete" || !this._element)
      return;
    const { textContent: t } = this._element;
    t === "" && (this._element.innerHTML = "");
  }
  /**
   * Create Tool's view
   *
   * @returns {HTMLDivElement}
   * @private
   */
  drawView() {
    const e = document.createElement("DIV");
    return e.classList.add(this._CSS.wrapper, this._CSS.block, this._CSS.alignment[this._data.alignment]), e.contentEditable = "false", e.dataset.placeholderActive = this.api.i18n.t(this._placeholder), this._data.text && (e.innerHTML = this._data.text), this.readOnly || (e.contentEditable = "true", e.addEventListener("keyup", this.onKeyUp)), e;
  }
  /**
   * Return Tool's view
   *
   * @returns {HTMLDivElement}
   */
  render() {
    return this._element = this.drawView(), this._element;
  }
  /**
   * Method that specified how to merge two Text blocks.
   * Called by Editor.js by backspace at the beginning of the Block
   *
   * @param {ParagraphData} data
   * @public
   */
  merge(e) {
    if (!this._element)
      return;
    this._data.text += e.text, this._data.alignment = e.alignment;
    const t = u(e.text);
    this._element.appendChild(t), this._element.normalize();
  }
  /**
   * Validate Paragraph block data:
   * - check for emptiness
   *
   * @param {ParagraphData} savedData — data received after saving
   * @returns {boolean} false if saved data is not correct, otherwise true
   * @public
   */
  validate(e) {
    return !(e.text.trim() === "" && !this._preserveBlank);
  }
  /**
   * Extract Tool's data from the view
   *
   * @param {HTMLDivElement} toolsContent - Paragraph tools rendered view
   * @returns {ParagraphData} - saved data
   * @public
   */
  save() {
    var e;
    return this._data.text = ((e = this._element) == null ? void 0 : e.innerHTML) ?? "", this.data;
  }
  /**
   * Apply visual representation of activated tune
   * @param tune 
   * @param status 
   */
  applyTune(e, t) {
    var n;
    (n = this._element) == null || n.classList.toggle(`${this._CSS.alignment[e]}`, t);
  }
  /**
   * @returns TunesMenuConfig
   */
  renderSettings() {
    return i.alignmentTunes.map((e) => ({
      icon: e.icon,
      label: this.api.i18n.t(e.title),
      name: e.name,
      toggle: e.toggle,
      closeOnActivate: !0,
      isActive: this.data.alignment.toString() === e.name,
      onActivate: () => {
        if (typeof e.action == "function") {
          e.action(e.name);
          return;
        }
        this.tuneToggled(e.name);
      }
    }));
  }
  /**
   * On paste callback fired from Editor.
   *
   * Handles pasted content with special processing for Microsoft Word HTML.
   * Word documents contain proprietary tags and formatting that can break
   * LaTeX compilation, so we clean them before insertion.
   *
   * @param {HTMLPasteEvent} event - event with pasted data
   */
  onPaste(e) {
    let t = e.detail.data.innerHTML;
    w(t) && (t = m(t));
    const n = {
      text: t,
      alignment: this.stringToAlignmentEnum(e.detail.data.style.textAlign) || i.DEFAULT_ALIGNMENT
    };
    this._data = n, window.requestAnimationFrame(() => {
      this._element && (this._element.innerHTML = this._data.text || "");
    });
  }
  /**
   * Stores all Tool's data
   */
  set data(e) {
    this._data.text = e.text || "", this._data.alignment = e.alignment || i.DEFAULT_ALIGNMENT;
  }
  /**
   * Get current Tools`s data
   */
  get data() {
    return this._data;
  }
  /**
   * Convert string to @ParagraphAlignmentsEnum
   * @param value 
   * @returns 
   */
  stringToAlignmentEnum(e) {
    for (const t in a)
      if (a[t] === e)
        return e;
  }
  /**
   * Tune has been toggled
   * @param tuneName 
   * @returns @void
   */
  tuneToggled(e) {
    const t = this.stringToAlignmentEnum(e);
    t && (this.resetTunes(), this.setTune(t), this.applyTune(t, !0));
  }
  /**
   * Set one tune
   * @param tune 
   * @param force - tune state
   */
  setTune(e) {
    this._data.alignment = e;
  }
  /**
   * Remove all tunes
   */
  resetTunes() {
    i.alignmentTunes.forEach((e) => {
      const t = this.stringToAlignmentEnum(e.name);
      t && this.applyTune(t, !1);
    });
  }
  /**
   * Enable Conversion Toolbar. Paragraph can be converted to/from other tools
   * @returns {ConversionConfig}
   */
  static get conversionConfig() {
    return {
      export: "text",
      // to convert Paragraph to other block, use 'text' property of saved data
      import: "text"
      // to covert other block's exported string to Paragraph, fill 'text' property of tool data
    };
  }
  /**
   * Sanitizer rules
   * @returns {SanitizerConfig} - Edtior.js sanitizer config
   */
  static get sanitize() {
    return {
      text: {
        br: !0
      }
    };
  }
  /**
   * Returns true to notify the core that read-only mode is supported
   *
   * @returns {boolean}
   */
  static get isReadOnlySupported() {
    return !0;
  }
  /**
   * Used by Editor paste handling API.
   * Provides configuration to handle P tags.
   *
   * @returns {PasteConfig} - Paragraph Paste Setting
   */
  static get pasteConfig() {
    return {
      tags: ["P"]
    };
  }
  /**
   * Icon and title for displaying at the Toolbox
   *
   * @returns {ToolboxConfig} - Paragraph Toolbox Setting
   */
  static get toolbox() {
    return {
      icon: d,
      title: "Text"
    };
  }
  /**
   * Available paragraph tools
   */
  static get alignmentTunes() {
    return [
      {
        name: "left",
        icon: h,
        title: "Align left",
        toggle: !0
      },
      {
        name: "center",
        icon: c,
        title: "Align center",
        toggle: !0
      },
      {
        name: "right",
        icon: p,
        title: "Align right",
        toggle: !0
      },
      {
        name: "justify",
        icon: g,
        title: "Justify",
        toggle: !0
      }
    ];
  }
}
export {
  i as default
};
