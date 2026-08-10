import { create } from "zustand";
import { persist } from "zustand/middleware";
import { useSyncExternalStore } from "react";
import { EditorServer } from "@/utils/editor/server";
import {
  Language,
  Locale,
  LocaleExtend,
  standardizeLocale,
} from "@/utils/locale";
import { type OfficeTheme } from "@/utils/editor/types";

/**
 * Resolves the language setting to an actual locale code.
 * If the language is set to "auto", it detects the browser's preferred language.
 */
function resolveLanguage(language: Language): Locale {
  if (language === LocaleExtend.Auto) {
    const browserLang =
      typeof navigator !== "undefined"
        ? navigator.language || (navigator as any).userLanguage
        : "en";
    return standardizeLocale(browserLang || "en");
  }
  return language as Locale;
}

interface AppState {
  // Document State
  server: EditorServer;

  // Settings State
  language: Language;
  theme: OfficeTheme;

  // Actions
  setState: (
    state: Partial<Pick<AppState, "language" | "theme">>,
  ) => void;
}

export const useAppStore = create<AppState>()(
  persist(
    (set) => ({
      // Document Initial State
      server: new EditorServer(),

      // Settings Initial State
      language: LocaleExtend.Auto,
      theme: "theme-white",

      // Settings Actions
      setState: (newState) => set((state) => ({ ...state, ...newState })),
    }),
    {
      name: "office-state",
      // Only persist settings, skip server instance
      partialize: (state) => ({
        language: state.language,
        theme: state.theme,
      }),
    },
  ),
);

/**
 * Hook to check if persist rehydration has completed.
 * Returns false during SSR and before localStorage state is loaded,
 * then true once the persisted state has been applied.
 */
export function useHasHydrated(): boolean {
  return useSyncExternalStore(
    (callback) => {
      const unsub = useAppStore.persist.onFinishHydration(callback);
      return unsub;
    },
    () => useAppStore.persist.hasHydrated(),
    () => false, // SSR: always false
  );
}

/**
 * Hook to get the resolved language (reactive).
 * When language setting is "auto", returns the detected browser language.
 * Re-renders automatically when language setting changes.
 */
export function useResolvedLanguage(): Locale {
  return useAppStore((state) => resolveLanguage(state.language));
}
