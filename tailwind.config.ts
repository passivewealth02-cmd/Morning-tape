import type { Config } from "tailwindcss";

const config: Config = {
  content: [
    "./pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        paper: "#F4EDE0",
        "paper-dark": "#E8DFCC",
        "paper-light": "#FBF6EB",
        ink: "#1A1612",
        "ink-dim": "#5C544A",
        accent: "#8B0000",
        "green-mark": "#2D5016",
      },
    },
  },
  plugins: [],
};

export default config;
