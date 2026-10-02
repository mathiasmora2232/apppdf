import type { ReactNode } from "react";
import "./globals.css";

export const metadata = {
  title: "SignFlow",
  description: "Firma documental por lotes y multifirmante",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
