"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.convertHtmlToWordRuns = convertHtmlToWordRuns;
exports.stripHtmlTags = stripHtmlTags;
const docx_1 = require("docx");
/**
 * Converts HTML content to Word document TextRun elements
 * Handles basic formatting: bold, italic, underline, and line breaks
 */
function convertHtmlToWordRuns(htmlContent) {
    if (!htmlContent || htmlContent.trim() === '') {
        return [new docx_1.TextRun({ text: '', size: 24, font: "Arial" })];
    }
    // Parse HTML and convert to TextRun elements
    const runs = [];
    // Split by line breaks first to handle <br> tags
    const lines = htmlContent.split(/<br\s*\/?>/i);
    lines.forEach((line, lineIndex) => {
        if (lineIndex > 0) {
            // Add line break for each <br> tag
            runs.push(new docx_1.TextRun({ text: '\n', size: 24, font: "Arial" }));
        }
        if (line.trim() === '') {
            return;
        }
        // Process formatting within the line
        const processedRuns = processFormattedText(line);
        runs.push(...processedRuns);
    });
    return runs.length > 0 ? runs : [new docx_1.TextRun({ text: htmlContent, size: 24, font: "Arial" })];
}
/**
 * Processes a line of text and converts HTML formatting to TextRun properties
 * Handles nested formatting and multiple formatting tags
 */
function processFormattedText(text) {
    if (!text || text.trim() === '') {
        return [];
    }
    const runs = [];
    // Use a stack-based approach to handle nested formatting
    const formatStack = [];
    let currentFormatting = { bold: false, italic: false, underline: false };
    // Split by HTML tags while preserving the tags
    const segments = text.split(/(<[^>]*>)/g);
    for (const segment of segments) {
        if (segment.startsWith('<') && segment.endsWith('>')) {
            // This is a tag - update formatting state
            const tagMatch = segment.match(/<\/?([biu]|strong|em)/i);
            if (tagMatch) {
                const tagName = tagMatch[1].toLowerCase();
                const isClosing = segment.startsWith('</');
                if (isClosing) {
                    // Pop from stack to restore previous formatting
                    if (formatStack.length > 0) {
                        currentFormatting = formatStack.pop();
                    }
                    else {
                        // Reset to default if no previous state
                        currentFormatting = { bold: false, italic: false, underline: false };
                    }
                }
                else {
                    // Push current state to stack and update formatting
                    formatStack.push({ ...currentFormatting });
                    switch (tagName) {
                        case 'b':
                        case 'strong':
                            currentFormatting.bold = true;
                            break;
                        case 'i':
                        case 'em':
                            currentFormatting.italic = true;
                            break;
                        case 'u':
                            currentFormatting.underline = true;
                            break;
                    }
                }
            }
        }
        else if (segment.trim()) {
            // This is text content - create a TextRun with current formatting
            runs.push(createTextRun(segment, currentFormatting));
        }
    }
    return runs.length > 0 ? runs : [createTextRun(text, { bold: false, italic: false, underline: false })];
}
/**
 * Creates a TextRun with the specified formatting
 */
function createTextRun(text, formatting) {
    const textRun = new docx_1.TextRun({
        text: text,
        size: 24,
        font: "Arial"
    });
    // Apply formatting
    if (formatting.bold) {
        textRun.bold = true;
    }
    if (formatting.italic) {
        textRun.italics = true;
    }
    if (formatting.underline) {
        textRun.underline = {};
    }
    return textRun;
}
/**
 * Alternative simpler approach - strips HTML tags and returns plain text
 * Use this if you want to completely remove formatting
 */
function stripHtmlTags(htmlContent) {
    if (!htmlContent)
        return '';
    return htmlContent
        .replace(/<br\s*\/?>/gi, '\n') // Convert <br> to newlines
        .replace(/<[^>]*>/g, '') // Remove all HTML tags
        .replace(/&nbsp;/g, ' ') // Convert non-breaking spaces
        .replace(/&amp;/g, '&') // Convert HTML entities
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .trim();
}
