let cachedIp: string | null = null
let ipPromise: Promise<string> | null = null

export function getClientDevice(): string {
  if (typeof navigator === 'undefined') return 'Chrome on Windows Laptop'
  const ua = navigator.userAgent
  const isMobile = /Mobi|Android|iPhone|iPad|iPod|Windows Phone/i.test(ua)

  // 1. Determine Device / Platform (Apple iOS must be checked before Android / Linux!)
  let osName = 'Windows Laptop'
  if (/iPhone/i.test(ua)) {
    osName = 'iPhone'
  } else if (/iPad/i.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)) {
    osName = 'iPad'
  } else if (/Android/i.test(ua)) {
    osName = isMobile ? 'Android Mobile' : 'Android Tablet'
  } else if (/Windows/i.test(ua)) {
    osName = isMobile ? 'Windows Mobile' : 'Windows Laptop'
  } else if (/Macintosh|Mac OS X/i.test(ua)) {
    osName = 'MacBook'
  } else if (/Linux/i.test(ua)) {
    osName = isMobile ? 'Android Mobile' : 'Linux PC'
  }

  // 2. Determine Browser accurately across iOS, Android and Desktop
  let browser = 'Safari'
  if (/CriOS/i.test(ua)) {
    browser = 'Chrome'
  } else if (/FxiOS/i.test(ua)) {
    browser = 'Firefox'
  } else if (/EdgiOS|Edg\//i.test(ua)) {
    browser = 'Edge'
  } else if (/Chrome\//i.test(ua)) {
    browser = 'Chrome'
  } else if (/Firefox\//i.test(ua)) {
    browser = 'Firefox'
  } else if (/Safari\//i.test(ua)) {
    browser = 'Safari'
  }

  return `${browser} on ${osName}`
}

export function formatAuditDevice(rawDevice?: string | null): { label: string; isMobile: boolean; icon: string } {
  if (!rawDevice || rawDevice.trim() === '' || rawDevice === 'Desktop Browser') {
    return { label: 'Chrome on Windows Laptop', isMobile: false, icon: '💻' }
  }

  const trimmed = rawDevice.trim()
  const d = trimmed.toLowerCase()

  // 1. Detect browser name
  let browser = 'Chrome'
  if (d.includes('crios') || d.includes('chrome')) browser = 'Chrome'
  else if (d.includes('edgios') || d.includes('edge') || d.includes('edg/')) browser = 'Edge'
  else if (d.includes('fxios') || d.includes('firefox')) browser = 'Firefox'
  else if (d.includes('safari')) browser = 'Safari'

  // 2. APPLE iOS DEVICES (iPhone, iPad, iPod) - MUST be evaluated before generic 'phone' or 'mobile'!
  if (d.includes('iphone') || d.includes('ios') || d.includes('ipod')) {
    const isSafari = d.includes('safari') && !d.includes('chrome') && !d.includes('crios')
    const finalBrowser = isSafari ? 'Safari' : browser
    return { label: `${finalBrowser} on iPhone`, isMobile: true, icon: '📱' }
  }

  if (d.includes('ipad')) {
    return { label: `${browser} on iPad`, isMobile: true, icon: '📱' }
  }

  // 3. ANDROID DEVICES
  if (d.includes('android')) {
    const isTablet = d.includes('tablet') && !d.includes('mobile')
    return { label: `${browser} on Android ${isTablet ? 'Tablet' : 'Mobile'}`, isMobile: true, icon: '📱' }
  }

  // 4. DESKTOP PLATFORMS
  if (d.includes('macintosh') || d.includes('mac os') || d.includes('macbook') || d.includes('darwin')) {
    return { label: `${browser} on MacBook`, isMobile: false, icon: '💻' }
  }

  if (d.includes('windows') || d.includes('win32') || d.includes('win64')) {
    return { label: `${browser} on Windows Laptop`, isMobile: false, icon: '💻' }
  }

  if (d.includes('linux')) {
    return { label: `${browser} on Linux PC`, isMobile: false, icon: '💻' }
  }

  // 5. GENERIC MOBILE FALLBACK (checked last, after specific iPhone/iPad/Android checks)
  if (d.includes('mobile') || d.includes('phone')) {
    return { label: `${browser} Mobile Device`, isMobile: true, icon: '📱' }
  }

  return { label: trimmed, isMobile: false, icon: '💻' }
}

export interface AuditNetworkInfo {
  ip: string
  label: string
  networkType: 'lan' | 'wan' | 'loopback'
  icon: string
  tooltip: string
}

export function formatAuditNetwork(rawIp?: string | null): AuditNetworkInfo {
  const ip = (rawIp || '').trim()

  if (!ip || ip === '127.0.0.1' || ip === 'localhost' || ip === '0.0.0.0' || ip === '::1') {
    return {
      ip: ip || '127.0.0.1',
      label: 'Local Host',
      networkType: 'loopback',
      icon: '💻',
      tooltip: 'Internal loopback / local system host',
    }
  }

  // Detect Private LAN / Wi-Fi subnets (192.168.x.x, 10.x.x.x, 172.16-31.x.x)
  if (
    ip.startsWith('192.168.') ||
    ip.startsWith('10.') ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(ip)
  ) {
    return {
      ip,
      label: `Local Network (${ip})`,
      networkType: 'lan',
      icon: '🏠',
      tooltip: `Office Wi-Fi / Local Area Network (${ip})`,
    }
  }

  // Public WAN / Cellular IP
  return {
    ip,
    label: ip,
    networkType: 'wan',
    icon: '🌐',
    tooltip: `Public Internet / Cellular Network: ${ip}`,
  }
}

export async function fetchClientIp(): Promise<string> {
  if (cachedIp) return cachedIp
  // Local network / browser determination without calling untrusted third-party services
  if (typeof window !== 'undefined' && window.location) {
    const host = window.location.hostname
    if (host && host !== 'localhost' && host !== '0.0.0.0' && host !== '127.0.0.1') {
      cachedIp = host
      return cachedIp
    }
  }
  cachedIp = '127.0.0.1'
  return cachedIp
}

export function getCachedClientIp(): string {
  if (cachedIp) return cachedIp
  if (typeof window !== 'undefined' && window.location?.hostname) {
    const h = window.location.hostname
    if (h && h !== 'localhost' && h !== '0.0.0.0') {
      cachedIp = h
      return h
    }
  }
  return '127.0.0.1'
}
