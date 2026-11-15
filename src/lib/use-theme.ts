"use client";

import { useState, useEffect, useCallback } from "react";

type Theme = "gendocx-light" | "gendocx-dark";

function getTheme(): Theme {
  if (typeof document === "undefined") return "gendocx-light";
  const t = document.documentElement.getAttribute("data-theme");
  return t === "gendocx-dark" ? "gendocx-dark" : "gendocx-light";
}

export function useTheme() {
  const [theme, setTheme] = useState<Theme>(getTheme);

  useEffect(() => {
    // Sync on mount in case SSR initial differs
    setTheme(getTheme());

    const observer = new MutationObserver(() => {
      setTheme(getTheme());
    });
    observer.observe(document.documentElement, {
      attributes: true,
      attributeFilter: ["data-theme"],
    });
    return () => observer.disconnect();
  }, []);

  const toggleTheme = useCallback(() => {
    const next: Theme = getTheme() === "gendocx-light" ? "gendocx-dark" : "gendocx-light";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("gendocx-theme", next);
    // State updates via MutationObserver
  }, []);

  return { theme, toggleTheme, isDark: theme === "gendocx-dark" };
}