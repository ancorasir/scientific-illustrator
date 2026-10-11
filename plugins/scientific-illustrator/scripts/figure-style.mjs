import { readFileSync } from "node:fs";

const data = JSON.parse(readFileSync(new URL("../styles/figure-profiles.json", import.meta.url), "utf8"));
export const PROFILE_NAMES = Object.freeze(Object.keys(data.profiles));
export const PARAGRAPH_FIELDS = Object.freeze(["line_spacing_pt", "paragraph_space_before_pt", "paragraph_space_after_pt"]);
const ptPerMm = 72 / 25.4;

function finite(value, name, minimum, maximum = 10000) {
  if (typeof value !== "number" || !Number.isFinite(value) || value < minimum || value > maximum) {
    throw new Error(`${name} must be a finite number between ${minimum} and ${maximum}.`);
  }
  return value;
}

// Pure, read-only computation. All preset sizes refer to the final cropped figure.
export function resolveFigureStyle(args) {
  const profileName = args.profile ?? "print-standard";
  const profile = data.profiles[profileName];
  if (!profile) throw new Error(`Unknown figure profile: ${profileName}`);
  const roleName = args.role ?? "body";
  const role = profile.roles[roleName];
  if (!role) throw new Error(`Unknown text role: ${roleName}`);
  const finalWidth = finite(args.final_width_mm, "final_width_mm", 1);
  const sourceWidth = finite(args.source_figure_width_mm ?? finalWidth, "source_figure_width_mm", 1);
  const insertionScale = finalWidth / sourceWidth;
  const sourceScale = 1 / insertionScale;
  const font = args.font_family ?? profile.fontFamily;
  if (typeof font !== "string" || !font.trim() || /[\r\n\0]/.test(font)) throw new Error("font_family must be a nonempty font name.");
  const size = finite(args.font_size_pt ?? role.fontPt, "font_size_pt", 1, 400);
  const leading = finite(args.leading_ratio ?? role.leadingRatio, "leading_ratio", 1, 3);
  const [horizontal, vertical] = role.paddingMm;
  const textStyle = {
    font_name: font, font_size: size * sourceScale, bold: role.bold,
    line_spacing_pt: size * leading * sourceScale,
    paragraph_space_before_pt: 0, paragraph_space_after_pt: 0,
    margin_left: horizontal * ptPerMm * sourceScale,
    margin_right: horizontal * ptPerMm * sourceScale,
    margin_top: vertical * ptPerMm * sourceScale,
    margin_bottom: vertical * ptPerMm * sourceScale,
    word_wrap: true, text_autofit: "none",
  };
  if (textStyle.font_size < 1 || textStyle.font_size > 400 || textStyle.line_spacing_pt < 1 || textStyle.line_spacing_pt > 1200 || Object.entries(textStyle).some(([k,v]) => k.startsWith("margin_") && v > 1000)) {
    throw new Error("Source scale exceeds the drawing API limits. Use a canvas closer to final figure size.");
  }
  const lineWidth = profile.lineWidthPt * sourceScale;
  if (lineWidth > 50) throw new Error("Source line width exceeds the drawing API limit.");
  const warnings = [];
  const budget = { checked: false };
  for (const [field,min] of [["final_height_mm",1],["body_width_mm",1],["body_height_mm",1],["caption_height_mm",0],["paragraph_gap_mm",0]]) {
    if (args[field] !== undefined) finite(args[field], field, min);
  }
  if (args.body_width_mm !== undefined && finalWidth > args.body_width_mm) warnings.push("Figure width exceeds body width.");
  const budgetFields = ["final_height_mm", "body_height_mm", "caption_height_mm", "paragraph_gap_mm"];
  const provided = budgetFields.filter(k => args[k] !== undefined);
  if (provided.length && provided.length !== budgetFields.length) {
    warnings.push("Height budget incomplete: provide final_height_mm, body_height_mm, caption_height_mm and paragraph_gap_mm explicitly.");
  } else if (provided.length) {
    budget.checked = true;
    budget.remaining_mm = args.body_height_mm - args.final_height_mm - args.caption_height_mm - args.paragraph_gap_mm;
    budget.fits = budget.remaining_mm >= 0;
    if (!budget.fits) warnings.push("Figure, caption and paragraph gaps exceed body height.");
  }
  return {
    profile: profileName, role: roleName,
    final_size: { width_mm: finalWidth, font_pt: size, line_spacing_pt: size * leading, line_width_pt: profile.lineWidthPt },
    insertion_scale: insertionScale, text_style: textStyle,
    line_style: { line_width: lineWidth },
    spacing_mm: Object.fromEntries(Object.entries(profile.spacingMm).map(([k,v]) => [k, v * sourceScale])),
    height_budget: budget, warnings,
    limits: "Computed style and space budget only; not font availability, text fit, native rendering or document pagination. Use crop width, not slide width. Presets are optional; do not restyle existing user objects implicitly.",
  };
}

// Call before dispatch. Older backends must never silently ignore new typography.
export function assertParagraphSupport(backend, args) {
  function hasParagraphSettings(value) {
    if (!value || typeof value !== "object") return false;
    let found = false;
    for (const field of PARAGRAPH_FIELDS) {
      if (value[field] !== undefined) {
        finite(value[field], field, field === "line_spacing_pt" ? 1 : 0, 1200);
        found = true;
      }
    }
    for (const child of Object.values(value)) {
      if (child && typeof child === "object") found = hasParagraphSettings(child) || found;
    }
    return found;
  }
  const hasSettings = hasParagraphSettings(args);
  if (backend !== "ooxml" && hasSettings) {
    throw new Error(`Explicit paragraph spacing is currently supported only by the OOXML working-copy backend; ${backend} would ignore it. Select OOXML in a fresh session, or omit these settings with an explicit limitation.`);
  }
}
