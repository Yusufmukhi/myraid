import QRCode from 'qrcode'

export interface AntiTamperMetadata {
  projectId: string
  orgId?: string
  orgName: string
  revision: string
  totalAmount: number
  generatedAt: string
  checksum: string
  verifyUrl: string
}

export async function generateChecksum(data: string): Promise<string> {
  if (typeof crypto !== 'undefined' && crypto.subtle) {
    try {
      const msgUint8 = new TextEncoder().encode(data)
      const hashBuffer = await crypto.subtle.digest('SHA-256', msgUint8)
      const hashArray = Array.from(new Uint8Array(hashBuffer))
      return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('')
    } catch {
      // Fallback
    }
  }
  // Simple deterministic hash fallback
  let hash = 0
  for (let i = 0; i < data.length; i++) {
    hash = (hash << 5) - hash + data.charCodeAt(i)
    hash |= 0
  }
  return Math.abs(hash).toString(16).padStart(16, '0')
}

export async function generateQuotationQRCode(
  projectId: string,
  orgName: string,
  revision: string,
  totalAmount: number,
  serverSignature?: string,
): Promise<{ qrDataUrl: string; checksum: string; shortCode: string }> {
  // If server provided an HMAC cryptographic signature, use it as unforgeable verification token
  const fullHash = serverSignature && serverSignature.length >= 16
    ? serverSignature
    : await generateChecksum(`${projectId}|${orgName}|${revision}|${Math.round(totalAmount)}|SECURE_QUOTATION_V1`)

  const shortCode = `VREF-${fullHash.slice(0, 4).toUpperCase()}-${fullHash.slice(4, 8).toUpperCase()}-${Math.round(totalAmount).toString().slice(-4)}`

  // Deep verification payload
  const verifyPayload = JSON.stringify({
    v: 1,
    p: projectId,
    o: orgName,
    r: revision,
    t: Math.round(totalAmount),
    sig: fullHash,
    c: shortCode,
    ts: new Date().toISOString().split('T')[0],
  })

  try {
    const qrDataUrl = await QRCode.toDataURL(verifyPayload, {
      width: 140,
      margin: 1,
      color: {
        dark: '#0f172a',
        light: '#ffffff',
      },
      errorCorrectionLevel: 'M',
    })
    return { qrDataUrl, checksum: fullHash, shortCode }
  } catch (err) {
    console.error('QR code generation error:', err)
    return { qrDataUrl: '', checksum: fullHash, shortCode }
  }
}
