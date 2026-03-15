declare interface CellStyle {
    fill?: Rgba;
    color?: Rgba;
    fontSize?: Pt;
    fontWeight?: FontWeight;
    fontStyle?: FontStyle;
    align?: TextAlign;
    borderTop?: boolean;
    borderBottom?: boolean;
    borderLeft?: boolean;
    borderRight?: boolean;
}

declare type FontStyle = 'normal' | 'italic';

declare type FontWeight = 'normal' | 'bold';

export declare interface GeneratedPDF {
    download(filename?: string): void;
    toBlob(): Blob;
    toDataUrl(): string;
}

declare interface ImageElement {
    type: 'image';
    x: Pt;
    y: Pt;
    width: Pt;
    height: Pt;
    src: string;
    alt?: string;
}

export declare interface IRDocument {
    format: SourceFormat;
    pages: IRPage[];
    metadata: {
        title?: string;
        author?: string;
        created?: string;
    };
}

export declare type IRElement = TextElement | ImageElement | ShapeElement | TableElement;

export declare interface IRPage {
    width: Pt;
    height: Pt;
    background: Rgba;
    elements: IRElement[];
    label?: string;
}

export declare function isSupportedFile(file: File): boolean;

declare interface ParseError {
    code: ParseErrorCode;
    message: string;
    detail?: string;
}

declare type ParseErrorCode = 'INVALID_FILE' | 'UNSUPPORTED_FORMAT' | 'CORRUPT_ZIP' | 'MISSING_ENTRY' | 'XML_PARSE_FAILED' | 'UNKNOWN';

declare type ParseResult = {
    ok: true;
    doc: IRDocument;
} | {
    ok: false;
    error: ParseError;
};

export declare interface PDFGeneratorOptions {
    pageNumbers?: boolean;
    title?: string;
    author?: string;
}

declare type Pt = number;

declare type Rgba = {
    r: number;
    g: number;
    b: number;
    a: number;
};

declare interface ShapeElement {
    type: 'shape';
    kind: ShapeKind;
    x: Pt;
    y: Pt;
    width: Pt;
    height: Pt;
    fill?: Rgba;
    stroke?: Rgba;
    strokeWidth: Pt;
}

declare type ShapeKind = 'rect' | 'ellipse' | 'line' | 'arrow';

declare type SourceFormat = 'xlsx' | 'pptx';

export declare const SUPPORTED_FORMATS: readonly [".pptx", ".xlsx"];

export declare type SupportedFormat = typeof SUPPORTED_FORMATS[number];

declare interface TableCell {
    value: string;
    colspan: number;
    rowspan: number;
    style: CellStyle;
}

declare interface TableElement {
    type: 'table';
    x: Pt;
    y: Pt;
    width: Pt;
    colWidths: Pt[];
    rows: TableRow[];
}

declare interface TableRow {
    cells: TableCell[];
    height: Pt;
}

declare type TextAlign = 'left' | 'center' | 'right' | 'justify';

declare interface TextElement {
    type: 'text';
    x: Pt;
    y: Pt;
    width: Pt;
    height: Pt;
    align: TextAlign;
    runs: TextRun[];
    lineHeight: number;
}

declare interface TextRun {
    text: string;
    fontSize: Pt;
    fontFamily: string;
    fontWeight: FontWeight;
    fontStyle: FontStyle;
    color: Rgba;
    underline: boolean;
    strikethrough: boolean;
    url?: string;
}

/**
 * Parse a file to the intermediate representation without generating a PDF.
 * Useful for inspecting the parsed structure.
 */
export declare function toIR(file: File): Promise<ParseResult>;

/**
 * Convert a .pptx or .xlsx File to a PDF.
 *
 * @example
 * const result = await toPDF(file)
 * if (result.ok) {
 *   result.pdf.download('slides.pdf')
 * }
 */
export declare function toPDF(file: File, opts?: ToPDFOptions): Promise<ToPDFResult>;

export declare interface ToPDFOptions extends PDFGeneratorOptions {
    /** Subset of page indices to include (0-indexed). Default: all pages */
    pages?: number[];
}

export declare type ToPDFResult = {
    ok: true;
    pdf: GeneratedPDF;
} | {
    ok: false;
    error: string;
};

export { }
