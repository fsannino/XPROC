// Headers HTTP de seguranca aplicados globalmente em next.config.ts.
// Extraido em modulo separado para ser testavel e nao depender de runtime Next.

type Header = { key: string; value: string };

type Env = {
  NODE_ENV?: string;
  CSP_REPORT_ONLY?: string;
  CSP_REPORT_URI?: string;
};

export const cspDirectives = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-inline'",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' data: https://fonts.gstatic.com",
  "img-src 'self' data: blob: https://avatars.githubusercontent.com",
  "connect-src 'self'",
  "frame-ancestors 'none'",
  "base-uri 'self'",
  "form-action 'self'",
  "object-src 'none'",
];

export function buildCspValue(reportUri?: string): string {
  const directives = reportUri ? [...cspDirectives, `report-uri ${reportUri}`] : cspDirectives;
  return directives.join("; ");
}

export function buildSecurityHeaders(env: Env = process.env): Header[] {
  const headers: Header[] = [
    { key: "X-Frame-Options", value: "DENY" },
    { key: "X-Content-Type-Options", value: "nosniff" },
    { key: "Referrer-Policy", value: "strict-origin-when-cross-origin" },
    { key: "Permissions-Policy", value: "camera=(), microphone=(), geolocation=()" },
  ];

  if (env.NODE_ENV === "production") {
    headers.push({ key: "Strict-Transport-Security", value: "max-age=31536000; includeSubDomains" });
  }

  if (env.CSP_REPORT_ONLY !== "off") {
    headers.push({
      key: "Content-Security-Policy-Report-Only",
      value: buildCspValue(env.CSP_REPORT_URI),
    });
  }

  return headers;
}
