import { describe, it, expect } from 'vitest'
import { buildCspValue, buildSecurityHeaders, cspDirectives } from '@/lib/security-headers'

describe('buildCspValue', () => {
  it('junta diretivas com "; "', () => {
    const value = buildCspValue()
    expect(value).toContain("default-src 'self'")
    expect(value).toContain("frame-ancestors 'none'")
    expect(value.split('; ').length).toBe(cspDirectives.length)
  })

  it('anexa report-uri quando fornecido', () => {
    const value = buildCspValue('https://csp.example.com/report')
    expect(value).toContain('report-uri https://csp.example.com/report')
    expect(value.split('; ').length).toBe(cspDirectives.length + 1)
  })

  it('omite report-uri quando ausente', () => {
    const value = buildCspValue()
    expect(value).not.toContain('report-uri')
  })

  it('inclui font-src com fonts.gstatic.com', () => {
    expect(buildCspValue()).toContain('font-src')
    expect(buildCspValue()).toContain('https://fonts.gstatic.com')
  })

  it('inclui style-src com fonts.googleapis.com e unsafe-inline (Tailwind/RF)', () => {
    const value = buildCspValue()
    expect(value).toMatch(/style-src[^;]*'unsafe-inline'/)
    expect(value).toContain('https://fonts.googleapis.com')
  })
})

describe('buildSecurityHeaders', () => {
  it('inclui X-Frame-Options DENY', () => {
    const headers = buildSecurityHeaders({})
    expect(headers).toContainEqual({ key: 'X-Frame-Options', value: 'DENY' })
  })

  it('inclui X-Content-Type-Options nosniff', () => {
    const headers = buildSecurityHeaders({})
    expect(headers).toContainEqual({ key: 'X-Content-Type-Options', value: 'nosniff' })
  })

  it('NAO inclui HSTS em dev (NODE_ENV !== production)', () => {
    const headers = buildSecurityHeaders({ NODE_ENV: 'development' })
    expect(headers.some((h) => h.key === 'Strict-Transport-Security')).toBe(false)
  })

  it('inclui HSTS em production', () => {
    const headers = buildSecurityHeaders({ NODE_ENV: 'production' })
    const hsts = headers.find((h) => h.key === 'Strict-Transport-Security')
    expect(hsts?.value).toContain('max-age=31536000')
    expect(hsts?.value).toContain('includeSubDomains')
  })

  it('inclui CSP Report-Only por padrao', () => {
    const headers = buildSecurityHeaders({})
    const csp = headers.find((h) => h.key === 'Content-Security-Policy-Report-Only')
    expect(csp).toBeDefined()
    expect(csp?.value).toContain("default-src 'self'")
  })

  it('omite CSP quando CSP_REPORT_ONLY=off', () => {
    const headers = buildSecurityHeaders({ CSP_REPORT_ONLY: 'off' })
    expect(headers.some((h) => h.key === 'Content-Security-Policy-Report-Only')).toBe(false)
  })

  it('NAO emite CSP em modo enforcing (so report-only nesta sprint)', () => {
    const headers = buildSecurityHeaders({})
    expect(headers.some((h) => h.key === 'Content-Security-Policy')).toBe(false)
  })

  it('encadeia report-uri quando CSP_REPORT_URI definido', () => {
    const headers = buildSecurityHeaders({ CSP_REPORT_URI: 'https://csp.example.com/report' })
    const csp = headers.find((h) => h.key === 'Content-Security-Policy-Report-Only')
    expect(csp?.value).toContain('report-uri https://csp.example.com/report')
  })
})
