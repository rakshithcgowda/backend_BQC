import { TextRun } from 'docx';
/**
 * Converts HTML content to Word document TextRun elements
 * Handles basic formatting: bold, italic, underline, and line breaks
 */
export declare function convertHtmlToWordRuns(htmlContent: string): TextRun[];
/**
 * Alternative simpler approach - strips HTML tags and returns plain text
 * Use this if you want to completely remove formatting
 */
export declare function stripHtmlTags(htmlContent: string): string;
//# sourceMappingURL=htmlToWord.d.ts.map