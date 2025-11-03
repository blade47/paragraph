import { API, ConversionConfig, HTMLPasteEvent, PasteConfig, SanitizerConfig, ToolConfig, ToolboxConfig } from '@editorjs/editorjs';
import { ActionConfig } from './types/types';
import { TunesMenuConfig } from '@editorjs/editorjs/types/tools';

/**
 * Base Paragraph Block for the Editor.js.
 * Represents a regular text block
 *
 * @author CodeX (team@codex.so)
 * @copyright CodeX 2018
 * @license The MIT License (MIT)
 */
/**
 * @typedef {object} ParagraphConfig
 * @property {string} placeholder - placeholder for the empty paragraph
 * @property {boolean} preserveBlank - Whether or not to keep blank paragraphs when saving editor data
 */
export interface ParagraphConfig extends ToolConfig {
    /**
     * Placeholder for the empty paragraph
     */
    placeholder?: string;
    /**
     * Whether or not to keep blank paragraphs when saving editor data
     */
    preserveBlank?: boolean;
}
/**
 * @typedef {object} ParagraphData
 * @description Tool's input and output data format
 * @property {string} text — Paragraph's content. Can include HTML tags: <a><b><i>
 */
export interface ParagraphData {
    /**
     * Paragraph's content
     */
    text: string;
    /**
     * Paragraph alignment
     */
    alignment: ParagraphAlignmentsEnum;
}
/**
 * @typedef {object} ParagraphParams
 * @description Constructor params for the Paragraph tool, use to pass initial data and settings
 * @property {ParagraphData} data - Preload data for the paragraph.
 * @property {ParagraphConfig} config - The configuration for the paragraph.
 * @property {API} api - The Editor.js API.
 * @property {boolean} readOnly - Is paragraph is read-only.
 */
interface ParagraphParams {
    /**
     * Initial data for the paragraph
     */
    data: ParagraphData;
    /**
     * Paragraph tool configuration
     */
    config: ParagraphConfig;
    /**
     * Editor.js API
     */
    api: API;
    /**
     * Is paragraph read-only.
     */
    readOnly: boolean;
}
interface ParagraphAlignments {
    left: string;
    center: string;
    right: string;
    justify: string;
}
declare enum ParagraphAlignmentsEnum {
    LEFT = "left",
    CENTER = "center",
    RIGHT = "right",
    JUSTIFY = "justify"
}
export default class Paragraph {
    /**
     * Default placeholder for Paragraph Tool
     *
     * @returns {string}
     * @class
     */
    static get DEFAULT_PLACEHOLDER(): string;
    /**
     * Allowed paragraph alignments
     *
     * @public
     * @returns {{left: string, center: string, right: string}}
    */
    static get ALIGNMENTS(): ParagraphAlignments;
    /**
     * Default paragraph alignment
     *
     * @public
     * @returns {string}
     */
    static get DEFAULT_ALIGNMENT(): ParagraphAlignmentsEnum;
    /**
     * The Editor.js API
     */
    api: API;
    /**
     * Paragraph's configuration
     */
    config: ParagraphConfig;
    /**
     * Is Paragraph Tool read-only
     */
    readOnly: boolean;
    /**
     * Paragraph Tool's CSS classes
     */
    private _CSS;
    /**
     * Placeholder for Paragraph Tool
     */
    private _placeholder;
    /**
     * Paragraph's data
     */
    private _data;
    /**
     * Paragraph's main Element
     */
    private _element;
    /**
     * Whether or not to keep blank paragraphs when saving editor data
     */
    private _preserveBlank;
    /**
     * Render plugin`s main Element and fill it with saved data
     *
     * @param {object} params - constructor params
     * @param {ParagraphData} params.data - previously saved data
     * @param {ParagraphConfig} params.config - user config for Tool
     * @param {object} params.api - editor.js api
     * @param {boolean} readOnly - read only mode flag
     */
    constructor({ data, config, api, readOnly }: ParagraphParams);
    /**
     * Check if text content is empty and set empty string to inner html.
     * We need this because some browsers (e.g. Safari) insert <br> into empty contenteditanle elements
     *
     * @param {KeyboardEvent} e - key up event
     */
    onKeyUp(e: KeyboardEvent): void;
    /**
     * Create Tool's view
     *
     * @returns {HTMLDivElement}
     * @private
     */
    drawView(): HTMLDivElement;
    /**
     * Return Tool's view
     *
     * @returns {HTMLDivElement}
     */
    render(): HTMLDivElement;
    /**
     * Method that specified how to merge two Text blocks.
     * Called by Editor.js by backspace at the beginning of the Block
     *
     * @param {ParagraphData} data
     * @public
     */
    merge(data: ParagraphData): void;
    /**
     * Validate Paragraph block data:
     * - check for emptiness
     *
     * @param {ParagraphData} savedData — data received after saving
     * @returns {boolean} false if saved data is not correct, otherwise true
     * @public
     */
    validate(savedData: ParagraphData): boolean;
    /**
     * Extract Tool's data from the view
     *
     * @param {HTMLDivElement} toolsContent - Paragraph tools rendered view
     * @returns {ParagraphData} - saved data
     * @public
     */
    save(): ParagraphData;
    /**
     * Apply visual representation of activated tune
     * @param tune
     * @param status
     */
    applyTune(tune: ParagraphAlignmentsEnum, status: boolean): void;
    /**
     * @returns TunesMenuConfig
     */
    renderSettings(): TunesMenuConfig;
    /**
     * On paste callback fired from Editor.
     *
     * Handles pasted content with special processing for Microsoft Word HTML.
     * Word documents contain proprietary tags and formatting that can break
     * LaTeX compilation, so we clean them before insertion.
     *
     * @param {HTMLPasteEvent} event - event with pasted data
     */
    onPaste(event: HTMLPasteEvent): void;
    /**
     * Stores all Tool's data
     */
    private set data(value);
    /**
     * Get current Tools`s data
     */
    private get data();
    /**
     * Convert string to @ParagraphAlignmentsEnum
     * @param value
     * @returns
     */
    private stringToAlignmentEnum;
    /**
     * Tune has been toggled
     * @param tuneName
     * @returns @void
     */
    private tuneToggled;
    /**
     * Set one tune
     * @param tune
     * @param force - tune state
     */
    private setTune;
    /**
     * Remove all tunes
     */
    private resetTunes;
    /**
     * Enable Conversion Toolbar. Paragraph can be converted to/from other tools
     * @returns {ConversionConfig}
     */
    static get conversionConfig(): ConversionConfig;
    /**
     * Sanitizer rules
     * @returns {SanitizerConfig} - Edtior.js sanitizer config
     */
    static get sanitize(): SanitizerConfig;
    /**
     * Returns true to notify the core that read-only mode is supported
     *
     * @returns {boolean}
     */
    static get isReadOnlySupported(): boolean;
    /**
     * Used by Editor paste handling API.
     * Provides configuration to handle P tags.
     *
     * @returns {PasteConfig} - Paragraph Paste Setting
     */
    static get pasteConfig(): PasteConfig;
    /**
     * Icon and title for displaying at the Toolbox
     *
     * @returns {ToolboxConfig} - Paragraph Toolbox Setting
     */
    static get toolbox(): ToolboxConfig;
    /**
     * Available paragraph tools
     */
    static get alignmentTunes(): Array<ActionConfig>;
}
export {};
