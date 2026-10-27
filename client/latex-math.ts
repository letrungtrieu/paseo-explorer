// Bridges Paseo's built-in md4c WASM parser and KaTeX MathML engine so `$...$` and `$$...$$` formulas render in chat.

export interface PaseoKatexEngine {
  renderToString(
    latexExpression: string,
    options?: {
      output?: "mathml" | "html" | "htmlAndMathml";
      displayMode?: boolean;
      throwOnError?: boolean;
      trust?: boolean;
    },
  ): string;
}

interface MarkdownParserFlags {
  underline?: boolean;
  latexMath?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  highlight?: boolean;
  hardSoftBreaks?: boolean;
  preserveBlankLines?: boolean;
  admonitions?: boolean;
}

interface PaseoMarkdownParserModule {
  parseMarkdownSync?: (markdownSource: string, flags?: MarkdownParserFlags) => unknown;
  parseMarkdown?: (markdownSource: string, flags?: MarkdownParserFlags) => Promise<unknown>;
  __paseoOfficeLatexPatched?: boolean;
}

interface PaseoKatexRendererProps {
  content: string;
  katex: PaseoKatexEngine | null | undefined;
  displayMode?: boolean;
  style?: Record<string, unknown>;
}

interface PaseoKatexRendererModule {
  KaTeXRenderer?: (props: PaseoKatexRendererProps) => unknown;
  __paseoOfficeLatexPatched?: boolean;
}

interface PaseoMarkdownBuildStylesModule {
  buildStyles?: (normalizedStyle: {
    paragraph?: { color?: string; fontSize?: number; fontFamily?: string };
  }) => {
    paragraph?: Record<string, unknown>;
    mathInline?: Record<string, unknown>;
    mathDisplay?: Record<string, unknown>;
  };
  __paseoOfficeLatexPatched?: boolean;
}

export interface MathDomElement {
  textContent: string | null;
  innerHTML?: string;
  style?: Record<string, string | undefined>;
  getAttribute(attributeName: string): string | null;
  setAttribute(attributeName: string, attributeValue: string): void;
  querySelector(selector: string): MathDomElement | null;
  querySelectorAll(selector: string): ArrayLike<MathDomElement>;
}

type MetroRequireFn = (moduleId: number) => unknown;

const KNOWN_BUILD_STYLES_MODULE_ID = 3974;
const KNOWN_MD4C_PARSER_MODULE_ID = 3975;
const KNOWN_KATEX_RENDERER_MODULE_ID = 3981;
const KNOWN_KATEX_ENGINE_MODULE_ID = 3989;
const HARDCODED_LIGHT_MATH_COLOR = "#1f2937";

let cachedKatexEngine: PaseoKatexEngine | null = null;

// Retrieve Paseo's Metro module loader (`globalThis.__r`) when running inside the Paseo web/desktop renderer.
function getMetroRequire(): MetroRequireFn | null {
  const globalScope = globalThis as unknown as { __r?: MetroRequireFn };
  return typeof globalScope.__r === "function" ? globalScope.__r : null;
}

// Extract a valid KaTeX engine instance from a raw Metro module export object.
function extractKatexFromModuleExports(moduleExports: unknown): PaseoKatexEngine | null {
  if (!moduleExports || typeof moduleExports !== "object") return null;
  const candidateRecord = moduleExports as {
    renderToString?: unknown;
    default?: { renderToString?: unknown };
  };
  if (typeof candidateRecord.renderToString === "function") {
    return candidateRecord as unknown as PaseoKatexEngine;
  }
  if (candidateRecord.default && typeof candidateRecord.default.renderToString === "function") {
    return candidateRecord.default as unknown as PaseoKatexEngine;
  }
  return null;
}

// Locate Paseo's bundled KaTeX engine so formulas can be converted to browser-native MathML without extra dependencies.
export function resolvePaseoKatexEngine(): PaseoKatexEngine | null {
  if (cachedKatexEngine) return cachedKatexEngine;
  const metroRequire = getMetroRequire();
  if (!metroRequire) return null;

  try {
    const directExports = metroRequire(KNOWN_KATEX_ENGINE_MODULE_ID);
    const directEngine = extractKatexFromModuleExports(directExports);
    if (directEngine) {
      cachedKatexEngine = directEngine;
      return directEngine;
    }
  } catch {
    // Fall back to nearby module scan if a Paseo update shifted Metro module IDs.
  }

  for (let candidateId = 3900; candidateId <= 4100; candidateId++) {
    try {
      const candidateExports = metroRequire(candidateId);
      const discoveredEngine = extractKatexFromModuleExports(candidateExports);
      if (discoveredEngine) {
        cachedKatexEngine = discoveredEngine;
        return discoveredEngine;
      }
    } catch {
      // Ignore uninitialized or unrelated Metro module IDs during fallback discovery.
    }
  }

  return null;
}

// Render a single LaTeX string into MathML markup using Paseo's KaTeX engine.
export function renderLatexToMathMl(
  katexEngine: PaseoKatexEngine,
  latexExpression: string,
  displayMode: boolean,
): string | null {
  const trimmedLatex = latexExpression.trim();
  if (!trimmedLatex) return null;
  try {
    return katexEngine.renderToString(trimmedLatex, {
      output: "mathml",
      displayMode,
      throwOnError: false,
      trust: false,
    });
  } catch {
    return null;
  }
}

// Decode HTML entities (`&amp;`, `&lt;`, `&gt;`) inside a raw HTML text segment before passing it to KaTeX.
function decodeHtmlEntitiesForLatex(rawHtmlFragment: string): string {
  return rawHtmlFragment
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

// Replace `$$...$$` and `$...$` segments in HTML/text with MathML while preserving `<code>`, `<pre>`, and `<a>` tags.
export function replaceDollarMathWithMathMl(
  rawHtmlOrTextContent: string,
  katexEngine: PaseoKatexEngine,
): { htmlContent: string; hasRenderedMath: boolean } {
  let hasRenderedMath = false;
  const openTag = "<";

  // Split out protected `<code>`, `<pre>`, `<a>`, and `<math>` blocks so `$` inside code/links is never touched.
  const protectedTagRegex =
    /(<(?:code|pre|a|math)\b[^>]*>[\s\S]*?<\/(?:code|pre|a|math)>)/gi;
  const segments = rawHtmlOrTextContent.split(protectedTagRegex);

  const transformedSegments = segments.map((segmentText, segmentIndex) => {
    // Odd indices are the captured protected `<code>`/`<pre>`/`<a>`/`<math>` blocks; return them untouched.
    if (segmentIndex % 2 === 1) {
      return segmentText;
    }

    // 1. Convert block display formulas (`$$...$$`, allowing `<br>` line breaks inside) first.
    const afterBlockMath = segmentText.replace(
      /\$\$([\s\S]+?)\$\$/g,
      (fullMatch, blockFormula: string) => {
        const cleanFormula = decodeHtmlEntitiesForLatex(blockFormula);
        const mathMlMarkup = renderLatexToMathMl(katexEngine, cleanFormula, true);
        if (!mathMlMarkup) return fullMatch;
        hasRenderedMath = true;
        return `${openTag}span role="math" data-paseo-math="display" style="display:block;margin:8px 0;text-align:center;">${mathMlMarkup}${openTag}/span>`;
      },
    );

    // 2. Convert inline formulas (`$...$`), ignoring currency numbers and HTML tags.
    return afterInlineMathSegment(afterBlockMath, katexEngine, openTag, () => {
      hasRenderedMath = true;
    });
  });

  return { htmlContent: transformedSegments.join(""), hasRenderedMath };
}

function afterInlineMathSegment(
  segmentHtml: string,
  katexEngine: PaseoKatexEngine,
  openTag: string,
  onRendered: () => void,
): string {
  return segmentHtml.replace(
    /(?<!\\|\$)\$([^\s$<>][^$\n<>]*?[^\s$<>]|[a-zA-Z0-9])\$/g,
    (fullMatch, inlineFormula: string) => {
      if (/^\d+(?:[.,]\d+)?$/.test(inlineFormula.trim())) {
        return fullMatch;
      }
      const cleanFormula = decodeHtmlEntitiesForLatex(inlineFormula);
      const mathMlMarkup = renderLatexToMathMl(katexEngine, cleanFormula, false);
      if (!mathMlMarkup) return fullMatch;
      onRendered();
      return `${openTag}span role="math" data-paseo-math="inline">${mathMlMarkup}${openTag}/span>`;
    },
  );
}

// Upgrade already-mounted `.enrm-text` paragraphs and math role placeholders in the DOM to MathML.
export function decorateMountedLatexElements(
  rootDocument: MathDomElement,
  explicitKatexEngine?: PaseoKatexEngine | null,
): void {
  const katexEngine = explicitKatexEngine ?? resolvePaseoKatexEngine();
  if (!katexEngine) return;

  // 1. Copy the textMessage color from `.enrm-text p` or `.enrm-text li` onto its parent `.enrm-text` container
  // so block-level `<span role="math">` elements inherit the exact same text color and background.
  const markdownContainers = Array.from(rootDocument.querySelectorAll(".enrm-text"));
  for (const markdownContainer of markdownContainers) {
    const sampleTextElement = markdownContainer.querySelector("p, li");
    const messageFontColor = sampleTextElement?.style?.color;
    if (messageFontColor && markdownContainer.style) {
      markdownContainer.style.color = messageFontColor;
    }
  }

  // 2. Upgrade native math role nodes and strip Paseo's light-mode `#F3F4F6` background and `#1F2937` text color.
  const mathRoleElements = Array.from(rootDocument.querySelectorAll('span[role="math"]'));
  for (const mathSpan of mathRoleElements) {
    if (mathSpan.style) {
      mathSpan.style.color = "inherit";
      mathSpan.style.backgroundColor = "transparent";
    }

    if (mathSpan.querySelector("math") || mathSpan.getAttribute("data-paseo-math-ready") === "true") {
      continue;
    }
    const rawLatexLabel = mathSpan.getAttribute("aria-label") ?? "";
    if (!rawLatexLabel) continue;
    const isDisplayFormula = (mathSpan.textContent ?? "").trim().startsWith("$$");
    const mathMlMarkup = renderLatexToMathMl(katexEngine, rawLatexLabel, isDisplayFormula);
    if (mathMlMarkup && "innerHTML" in mathSpan) {
      mathSpan.innerHTML = mathMlMarkup;
      mathSpan.setAttribute("data-paseo-math-ready", "true");
    }
  }

  // 3. Upgrade already-mounted `.enrm-text p` and leaf `.enrm-text li > span` blocks holding unparsed `$...$` text.
  const markdownTextBlocks = Array.from(
    rootDocument.querySelectorAll(".enrm-text p, .enrm-text li > span, .enrm-text li"),
  );
  for (const textBlock of markdownTextBlocks) {
    if (
      textBlock.getAttribute("data-paseo-math-scanned") === "true" ||
      textBlock.querySelector("ul, ol, p")
    ) {
      continue;
    }
    const rawBlockText = textBlock.textContent ?? "";
    if (!rawBlockText.includes("$")) continue;

    const sourceMarkup =
      typeof textBlock.innerHTML === "string" && textBlock.innerHTML.length > 0
        ? textBlock.innerHTML
        : rawBlockText;
    const { htmlContent, hasRenderedMath } = replaceDollarMathWithMathMl(sourceMarkup, katexEngine);
    textBlock.setAttribute("data-paseo-math-scanned", "true");
    if (hasRenderedMath && "innerHTML" in textBlock) {
      textBlock.innerHTML = htmlContent;
    }
  }
}

// Normalize KaTeXRenderer inline style so formulas share the surrounding textMessage fontColor and background.
function normalizeKatexContainerStyle(
  rawStyle: Record<string, unknown> | undefined,
  displayMode: boolean | undefined,
): Record<string, unknown> {
  const rawColor = typeof rawStyle?.color === "string" ? rawStyle.color.trim().toLowerCase() : "";
  const resolvedColor =
    !rawColor || rawColor === HARDCODED_LIGHT_MATH_COLOR ? "inherit" : rawStyle?.color;
  return {
    ...rawStyle,
    color: resolvedColor,
    backgroundColor: "transparent",
    ...(displayMode ? { padding: "4px 0" } : null),
  };
}

// Enable `latexMath: true` on Paseo's md4c parser and inject KaTeX into `KaTeXRenderer` for all new Markdown renders.
export function enablePaseoLatexMathRuntime(): () => void {
  const metroRequire = getMetroRequire();
  if (!metroRequire) {
    return () => {};
  }

  const restoreCallbacks: Array<() => void> = [];
  const katexEngine = resolvePaseoKatexEngine();

  try {
    const buildStylesModule = metroRequire(
      KNOWN_BUILD_STYLES_MODULE_ID,
    ) as PaseoMarkdownBuildStylesModule | null;
    if (
      buildStylesModule &&
      !buildStylesModule.__paseoOfficeLatexPatched &&
      typeof buildStylesModule.buildStyles === "function"
    ) {
      const originalBuildStyles = buildStylesModule.buildStyles;
      buildStylesModule.buildStyles = (normalizedStyle) => {
        const builtStyles = originalBuildStyles(normalizedStyle);
        const paragraphColor = normalizedStyle?.paragraph?.color ?? "inherit";
        if (builtStyles && typeof builtStyles === "object") {
          builtStyles.mathInline = {
            ...builtStyles.mathInline,
            color: paragraphColor,
            backgroundColor: "transparent",
          };
          builtStyles.mathDisplay = {
            ...builtStyles.mathDisplay,
            color: paragraphColor,
            backgroundColor: "transparent",
            padding: 4,
          };
        }
        return builtStyles;
      };
      buildStylesModule.__paseoOfficeLatexPatched = true;

      restoreCallbacks.push(() => {
        buildStylesModule.buildStyles = originalBuildStyles;
        buildStylesModule.__paseoOfficeLatexPatched = false;
      });
    }
  } catch {
    // Ignore if buildStyles module is unavailable in the current host environment.
  }

  try {
    const parserModule = metroRequire(KNOWN_MD4C_PARSER_MODULE_ID) as PaseoMarkdownParserModule | null;
    if (
      parserModule &&
      !parserModule.__paseoOfficeLatexPatched &&
      typeof parserModule.parseMarkdownSync === "function" &&
      typeof parserModule.parseMarkdown === "function"
    ) {
      const originalParseSync = parserModule.parseMarkdownSync;
      const originalParseAsync = parserModule.parseMarkdown;

      parserModule.parseMarkdownSync = (markdownSource, flags) =>
        originalParseSync(markdownSource, { ...flags, latexMath: true });
      parserModule.parseMarkdown = (markdownSource, flags) =>
        originalParseAsync(markdownSource, { ...flags, latexMath: true });
      parserModule.__paseoOfficeLatexPatched = true;

      restoreCallbacks.push(() => {
        parserModule.parseMarkdownSync = originalParseSync;
        parserModule.parseMarkdown = originalParseAsync;
        parserModule.__paseoOfficeLatexPatched = false;
      });
    }
  } catch {
    // Ignore if md4c parser module is unavailable in the current host environment.
  }

  try {
    const rendererModule = metroRequire(
      KNOWN_KATEX_RENDERER_MODULE_ID,
    ) as PaseoKatexRendererModule | null;
    if (
      rendererModule &&
      !rendererModule.__paseoOfficeLatexPatched &&
      typeof rendererModule.KaTeXRenderer === "function" &&
      katexEngine
    ) {
      const originalKaTeXRenderer = rendererModule.KaTeXRenderer;
      rendererModule.KaTeXRenderer = (props) =>
        originalKaTeXRenderer({
          ...props,
          katex: props.katex ?? katexEngine,
          style: normalizeKatexContainerStyle(props.style, props.displayMode),
        });
      rendererModule.__paseoOfficeLatexPatched = true;

      restoreCallbacks.push(() => {
        rendererModule.KaTeXRenderer = originalKaTeXRenderer;
        rendererModule.__paseoOfficeLatexPatched = false;
      });
    }
  } catch {
    // Ignore if KaTeXRenderer module is unavailable in the current host environment.
  }

  return () => {
    for (const restoreFn of restoreCallbacks) {
      restoreFn();
    }
  };
}
