import type { NextConfig } from "next";
import { buildSecurityHeaders } from "./src/lib/security-headers";

// CSP esta em modo Report-Only por padrao: observa violacoes sem bloquear.
// Para desligar (troubleshooting): CSP_REPORT_ONLY=off
// Para coletar violacoes em endpoint: CSP_REPORT_URI=https://...
const nextConfig: NextConfig = {
  async headers() {
    return [
      {
        source: "/:path*",
        headers: buildSecurityHeaders(),
      },
    ];
  },
};

export default nextConfig;
