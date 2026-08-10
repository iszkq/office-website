export const locales = [
  "am",
  "ar",
  "bg",
  "bn",
  "ca",
  "cs",
  "da",
  "de",
  "el",
  "en",
  "es-419",
  "es",
  "et",
  "fa",
  "fi",
  "fil",
  "fr",
  "gu",
  "he",
  "hi",
  "hr",
  "hu",
  "id",
  "it",
  "ja",
  "kn",
  "ko",
  "lt",
  "lv",
  "ml",
  "mr",
  "ms",
  "nl",
  "no",
  "pl",
  "pt-BR",
  "pt-PT",
  "ro",
  "ru",
  "sk",
  "sl",
  "sr",
  "sv",
  "sw",
  "ta",
  "te",
  "th",
  "tr",
  "uk",
  "vi",
  "zh-CN",
  "zh-TW",
] as const;

export type Locale = (typeof locales)[number];
export type Language = Locale | "auto";

export const Locale = {
  EN: "en",
  JA: "ja",
  KO: "ko",
  ZH_CN: "zh-CN",
  ZH_TW: "zh-TW",
} as const satisfies Record<string, Locale>;

export const LocaleExtend = {
  Auto: "auto",
} as const;

const localeDisplayNames = new Intl.DisplayNames(["en"], { type: "language" });

export const LocaleName = Object.fromEntries(
  locales.map((locale) => [locale, localeDisplayNames.of(locale) ?? locale])
) as Record<Locale, string>;

export const standardizeLocale = (input: string): Locale => {
  const normalized = input.replace("_", "-").toLowerCase();
  const exactMatch = locales.find((locale) => locale.toLowerCase() === normalized);
  if (exactMatch) return exactMatch;

  const language = normalized.split("-")[0];
  return locales.find((locale) => locale.toLowerCase() === language) ?? Locale.EN;
};
