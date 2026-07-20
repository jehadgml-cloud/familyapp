/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./src/pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/components/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/app/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: {
          DEFAULT: "#10b981", // Emerald
          hover: "#059669",
          light: "#ecfdf5",
        },
        secondary: {
          DEFAULT: "#0f766e", // Deep Teal
          light: "#f0fdfa",
        },
        accent: "#3b82f6", // Blue
        danger: "#ef4444",
        warning: "#f59e0b",
        background: "#f8fafc",
        textMain: "#0f172a",
        textMuted: "#64748b",
        borderDark: "#e2e8f0",
      },
      fontFamily: {
        cairo: ["Cairo", "system-ui", "sans-serif"],
      },
      boxShadow: {
        premium: "0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05)",
        premiumLg: "0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05)",
      },
    },
  },
  plugins: [],
}
