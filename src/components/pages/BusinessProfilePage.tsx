import React, { useEffect, useState } from 'react'
import { Org, Session } from '../../lib/api'
import {
  BusinessProfile,
  DEFAULT_BUSINESS_PROFILE,
  fetchBusinessProfileFromSupabase,
  getBusinessProfile,
  resetBusinessProfile,
  saveBusinessProfile,
} from '../../lib/businessProfile'
import { MYRIAD_LOGO_DATA_URL, MYRIAD_WATERMARK_DATA_URL } from '../../assets/companyAssets'
import { Icon, useUI } from '../ui'

interface Props {
  session: Session
  org: Org
}

export function BusinessProfilePage({ session, org }: Props) {
  const { toast, confirm } = useUI()
  const role = org?.role || 'member'
  const isAdminOrOwner = role === 'owner' || role === 'admin'

  const [profile, setProfile] = useState<BusinessProfile>(() => getBusinessProfile(org.id))
  const [hasChanges, setHasChanges] = useState(false)
  const [logoPreview, setLogoPreview] = useState<string>(profile.logo_url || MYRIAD_LOGO_DATA_URL)
  const [watermarkPreview, setWatermarkPreview] = useState<string>(profile.watermark_url || MYRIAD_WATERMARK_DATA_URL)

  useEffect(() => {
    fetchBusinessProfileFromSupabase(session.token, org.id).then((loaded) => {
      if (loaded) {
        setProfile(loaded)
        setLogoPreview(loaded.logo_url || MYRIAD_LOGO_DATA_URL)
        setWatermarkPreview(loaded.watermark_url || MYRIAD_WATERMARK_DATA_URL)
      }
    })
    const loaded = getBusinessProfile(org.id)
    setProfile(loaded)
    setLogoPreview(loaded.logo_url || MYRIAD_LOGO_DATA_URL)
    setWatermarkPreview(loaded.watermark_url || MYRIAD_WATERMARK_DATA_URL)
    setHasChanges(false)
  }, [session.token, org.id])

  if (!isAdminOrOwner) {
    return (
      <div className="page-container" style={{ padding: '32px 24px', maxWidth: 800, margin: '0 auto' }}>
        <div className="card" style={{ padding: 32, textAlign: 'center', border: '1px solid #e5e7eb' }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>🔒</div>
          <h2 style={{ fontSize: 20, fontWeight: 700, marginBottom: 8, color: '#111827' }}>
            Access Restricted
          </h2>
          <p style={{ color: '#4b5563', fontSize: 14, maxWidth: 480, margin: '0 auto' }}>
            Only Organisation <strong>Owners</strong> and <strong>Administrators</strong> are authorized to view and edit company business details, registered addresses, tax identifiers, and PDF quotation header branding.
          </p>
        </div>
      </div>
    )
  }

  const handleChange = (field: keyof BusinessProfile, value: any) => {
    setProfile((prev) => ({
      ...prev,
      [field]: value,
    }))
    setHasChanges(true)
  }

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return

    if (!file.type.startsWith('image/')) {
      toast('Please upload an image file (PNG, JPG, or SVG)', 'err')
      return
    }

    if (file.size > 2 * 1024 * 1024) {
      toast('Logo file size must be under 2MB', 'err')
      return
    }

    const reader = new FileReader()
    reader.onload = (event) => {
      const result = event.target?.result as string
      if (result) {
        setLogoPreview(result)
        handleChange('logo_url', result)
        toast('Custom logo uploaded successfully', 'ok')
      }
    }
    reader.readAsDataURL(file)
  }

  const handleWatermarkUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return

    if (!file.type.startsWith('image/')) {
      toast('Please upload an image file (PNG, JPG, or SVG)', 'err')
      return
    }

    if (file.size > 2 * 1024 * 1024) {
      toast('Watermark file size must be under 2MB', 'err')
      return
    }

    const reader = new FileReader()
    reader.onload = (event) => {
      const result = event.target?.result as string
      if (result) {
        setWatermarkPreview(result)
        handleChange('watermark_url', result)
        toast('Custom watermark uploaded successfully', 'ok')
      }
    }
    reader.readAsDataURL(file)
  }

  const handleResetLogo = () => {
    setLogoPreview(MYRIAD_LOGO_DATA_URL)
    handleChange('logo_url', MYRIAD_LOGO_DATA_URL)
    toast('Logo reset to default template asset', 'ok')
  }

  const handleResetWatermark = () => {
    setWatermarkPreview(MYRIAD_WATERMARK_DATA_URL)
    handleChange('watermark_url', MYRIAD_WATERMARK_DATA_URL)
    toast('Watermark reset to default template asset', 'ok')
  }

  const handleSave = () => {
    if (!profile.company_name?.trim()) {
      toast('Company Name is required', 'err')
      return
    }

    saveBusinessProfile(profile, org.id, session.username)
    setHasChanges(false)
    toast('Business profile and PDF branding saved successfully!', 'ok')
  }

  const handleResetDefaults = async () => {
    const ok = await confirm({
      title: 'Reset to Defaults?',
      body: 'Are you sure you want to reset all business and quotation header details to the default Myriad template values? This cannot be undone.',
      confirmLabel: 'Reset All',
      danger: true,
    })
    if (!ok) return

    const reset = resetBusinessProfile(org.id, session.username)
    setProfile(reset)
    setLogoPreview(reset.logo_url || MYRIAD_LOGO_DATA_URL)
    setWatermarkPreview(reset.watermark_url || MYRIAD_WATERMARK_DATA_URL)
    setHasChanges(false)
    toast('Business profile reset to default Myriad details.', 'ok')
  }

  return (
    <div className="page-container" style={{ padding: '24px 32px', maxWidth: 1100, margin: '0 auto' }}>
      {/* Top Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24, gap: 16 }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
            <h1 style={{ fontSize: 24, fontWeight: 700, margin: 0, color: '#111827' }}>
              Business Profile &amp; PDF Branding
            </h1>
            <span
              style={{
                fontSize: 11,
                fontWeight: 700,
                padding: '2px 8px',
                borderRadius: 4,
                border: '1px solid #111827',
                background: '#f9fafb',
                color: '#111827',
                textTransform: 'uppercase',
              }}
            >
              {role.toUpperCase()} ONLY
            </span>
          </div>
          <p style={{ margin: 0, fontSize: 13, color: '#4b5563' }}>
            Configure your organisation's company name, registered addresses, factory works, GSTIN, contact numbers, emails, logo, and watermark. All changes immediately reflect across your official PDF quotations.
          </p>
        </div>

        <div style={{ display: 'flex', gap: 10, flexShrink: 0 }}>
          <button
            type="button"
            className="btn ghost sm"
            onClick={handleResetDefaults}
            title="Reset back to default Myriad template"
            style={{ border: '1px solid #d1d5db' }}
          >
            Reset Defaults
          </button>
          <button
            type="button"
            className="btn primary sm"
            onClick={handleSave}
            style={{ fontWeight: 600, minWidth: 120 }}
          >
            {hasChanges ? 'Save Changes *' : 'Save Changes'}
          </button>
        </div>
      </div>

      {/* Live PDF Header Preview Section */}
      <div className="card" style={{ marginBottom: 24, padding: 20, border: '1px solid #d1d5db', background: '#ffffff' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14, borderBottom: '1px solid #e5e7eb', paddingBottom: 10 }}>
          <div>
            <strong style={{ fontSize: 13, textTransform: 'uppercase', letterSpacing: '0.05em', color: '#374151' }}>
              Live PDF Header Preview
            </strong>
            <span style={{ fontSize: 12, color: '#6b7280', marginLeft: 8 }}>
              (Exact layout as rendered on Page 1 &amp; Page 2 of official quotations)
            </span>
          </div>
          {profile.updated_at && (
            <span style={{ fontSize: 11, color: '#6b7280' }}>
              Last saved: {new Date(profile.updated_at).toLocaleDateString()} by {profile.updated_by || 'Admin'}
            </span>
          )}
        </div>

        <div
          style={{
            background: '#ffffff',
            padding: '16px 20px',
            border: '1px solid #9ca3af',
            borderRadius: 4,
            boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
          }}
        >
          <div className="pdf-header" style={{ paddingTop: 0, marginBottom: 0 }}>
            <div className="pdf-header-top">
              <div className="pdf-logo-box" style={{ maxWidth: 240 }}>
                {logoPreview ? (
                  <img
                    src={logoPreview}
                    alt={profile.company_name || 'Company Logo'}
                    className="pdf-logo-img"
                    style={{ maxHeight: 72, objectFit: 'contain' }}
                  />
                ) : (
                  <div style={{ fontWeight: 800, fontSize: 18, color: '#000' }}>
                    {profile.company_name || 'COMPANY NAME'}
                  </div>
                )}
                {profile.tagline && (
                  <div style={{ fontSize: 9.5, color: '#374151', marginTop: 4, fontWeight: 500 }}>
                    {profile.tagline}
                  </div>
                )}
              </div>
              <div className="pdf-company-info" style={{ textAlign: 'right', fontSize: 10.3, lineHeight: 1.36, color: '#000', fontWeight: 600 }}>
                <div>{profile.registered_address_line1 || 'Registered Office Address Line 1'}</div>
                <div>{profile.registered_address_line2 || 'City, State, Pin Code'}</div>
                <div>
                  Contact : {profile.contact_phone || 'Phone Number'}
                  {profile.alternate_phone ? ` / ${profile.alternate_phone}` : ''}
                </div>
                <div>
                  GST NO :{profile.gst_no || 'GSTIN'}
                  {profile.pan_no ? ` | PAN: ${profile.pan_no}` : ''}
                </div>
                {profile.info_email && (
                  <div>
                    <a href={`mailto:${profile.info_email}`} style={{ color: '#0000cc', textDecoration: 'underline' }}>
                      {profile.info_email}
                    </a>
                  </div>
                )}
                {profile.sales_email && (
                  <div>
                    <a href={`mailto:${profile.sales_email}`} style={{ color: '#0000cc', textDecoration: 'underline' }}>
                      {profile.sales_email}
                    </a>
                  </div>
                )}
                {profile.website && (
                  <div>
                    <a
                      href={profile.website.startsWith('http') ? profile.website : `https://${profile.website}`}
                      target="_blank"
                      rel="noreferrer"
                      style={{ color: '#0000cc', textDecoration: 'underline' }}
                    >
                      {profile.website}
                    </a>
                  </div>
                )}
              </div>
            </div>
            <div
              className="pdf-factory-address"
              style={{
                marginTop: 10,
                paddingTop: 6,
                borderTop: '1px solid #d1d5db',
                textAlign: 'center',
                fontSize: 10.2,
                lineHeight: 1.35,
                color: '#000',
              }}
            >
              <div>{profile.factory_address_line1 || 'Factory / Works Address Line 1'}</div>
              {profile.factory_address_line2 && <div>{profile.factory_address_line2}</div>}
            </div>
          </div>
        </div>
      </div>

      {/* Main Form Fields */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {/* Section 1: Company & Brand Identity */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            1. Company &amp; Brand Identity
          </h3>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: 16 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Company Legal Name *
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. MYRIAD STORAGE SYSTEM LLP"
                value={profile.company_name}
                onChange={(e) => handleChange('company_name', e.target.value)}
              />
              <span style={{ fontSize: 11, color: '#6b7280' }}>Displayed on header, terms, and authorized signatures</span>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Tagline / Subtitle
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. Manufacturer of Mobile Compactor & Storage Systems"
                value={profile.tagline || ''}
                onChange={(e) => handleChange('tagline', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Signatory Title
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. Authorized Signatory"
                value={profile.signatory_title || ''}
                onChange={(e) => handleChange('signatory_title', e.target.value)}
              />
            </div>
          </div>
        </div>

        {/* Section 2: Registered Office Address */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            2. Registered Office Address
          </h3>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Registered Address Line 1
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. 702/A, AL-HUSSAIN, MOMIN NAGAR, JOGESHWARI (W)"
                value={profile.registered_address_line1}
                onChange={(e) => handleChange('registered_address_line1', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Registered Address Line 2 (City, State, Pin Code &amp; State Code)
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. MUMBAI - 400102. State : Maharashtra, Code : 27"
                value={profile.registered_address_line2}
                onChange={(e) => handleChange('registered_address_line2', e.target.value)}
              />
            </div>
          </div>
        </div>

        {/* Section 3: Factory & Works Address */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            3. Factory &amp; Works Address
          </h3>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Factory Address Line 1
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. Factory Address : Gala No 5, Vakan Compound, Near Dynamic Co, Vakan Pada Raod, Vasai Phata, Vasai East Dis.,"
                value={profile.factory_address_line1}
                onChange={(e) => handleChange('factory_address_line1', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Factory Address Line 2
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. Palghar 401208."
                value={profile.factory_address_line2 || ''}
                onChange={(e) => handleChange('factory_address_line2', e.target.value)}
              />
            </div>
          </div>
        </div>

        {/* Section 4: Contact & Online Presence */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            4. Contact, Email &amp; Web Presence
          </h3>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 16 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Primary Contact Phone
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. 9769377043"
                value={profile.contact_phone}
                onChange={(e) => handleChange('contact_phone', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Alternate Phone / Mobile
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. +91 98765 43210 (optional)"
                value={profile.alternate_phone || ''}
                onChange={(e) => handleChange('alternate_phone', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Info Email
              </label>
              <input
                type="email"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. info@myriadstoragesystem.com"
                value={profile.info_email}
                onChange={(e) => handleChange('info_email', e.target.value)}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Sales Email
              </label>
              <input
                type="email"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. sales@myriadstoragesystem.com"
                value={profile.sales_email || ''}
                onChange={(e) => handleChange('sales_email', e.target.value)}
              />
            </div>

            <div style={{ gridColumn: 'span 2' }}>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                Official Website
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. www.myriadstoragesystem.com"
                value={profile.website}
                onChange={(e) => handleChange('website', e.target.value)}
              />
            </div>
          </div>
        </div>

        {/* Section 5: Tax & Identifiers */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            5. Legal &amp; Tax Identifiers
          </h3>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 16 }}>
            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                GSTIN / GST Number
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. 27ABJFM5468N1ZG"
                value={profile.gst_no}
                onChange={(e) => handleChange('gst_no', e.target.value.toUpperCase())}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: 12, fontWeight: 600, marginBottom: 4, color: '#374151' }}>
                PAN Number
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', boxSizing: 'border-box' }}
                placeholder="e.g. ABJFM5468N"
                value={profile.pan_no || ''}
                onChange={(e) => handleChange('pan_no', e.target.value.toUpperCase())}
              />
            </div>
          </div>
        </div>

        {/* Section 6: Company Logo & Watermark Assets */}
        <div className="card" style={{ padding: 20, border: '1px solid #e5e7eb' }}>
          <h3 style={{ fontSize: 14, fontWeight: 700, margin: '0 0 16px 0', textTransform: 'uppercase', color: '#111827', borderBottom: '1px solid #f3f4f6', paddingBottom: 8 }}>
            6. Company Logo &amp; Watermark Assets
          </h3>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 24 }}>
            {/* Logo Card */}
            <div style={{ border: '1px solid #e5e7eb', borderRadius: 6, padding: 16, background: '#fafafa' }}>
              <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8, color: '#111827' }}>
                Company Logo (Quotation Header)
              </div>
              <div style={{ minHeight: 90, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#fff', border: '1px dashed #d1d5db', borderRadius: 4, padding: 12, marginBottom: 12 }}>
                {logoPreview ? (
                  <img
                    src={logoPreview}
                    alt="Logo preview"
                    style={{ maxHeight: 70, maxWidth: '100%', objectFit: 'contain' }}
                  />
                ) : (
                  <span style={{ fontSize: 12, color: '#9ca3af' }}>No logo selected</span>
                )}
              </div>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <label className="btn ghost sm" style={{ cursor: 'pointer', border: '1px solid #d1d5db' }}>
                  Upload New Logo
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: 'none' }}
                    onChange={handleLogoUpload}
                  />
                </label>
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={handleResetLogo}
                  style={{ color: '#4b5563' }}
                >
                  Use Default Logo
                </button>
              </div>
              <span style={{ display: 'block', fontSize: 11, color: '#6b7280', marginTop: 6 }}>
                Recommended: Transparent PNG or SVG, ~220px width.
              </span>
            </div>

            {/* Watermark Card */}
            <div style={{ border: '1px solid #e5e7eb', borderRadius: 6, padding: 16, background: '#fafafa' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <div style={{ fontWeight: 600, fontSize: 13, color: '#111827' }}>
                  PDF Page Watermark
                </div>
                <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, cursor: 'pointer' }}>
                  <input
                    type="checkbox"
                    checked={profile.show_watermark !== false}
                    onChange={(e) => handleChange('show_watermark', e.target.checked)}
                  />
                  <span>Show Watermark</span>
                </label>
              </div>

              <div style={{ minHeight: 90, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#fff', border: '1px dashed #d1d5db', borderRadius: 4, padding: 12, marginBottom: 12 }}>
                {watermarkPreview && profile.show_watermark !== false ? (
                  <img
                    src={watermarkPreview}
                    alt="Watermark preview"
                    style={{ maxHeight: 70, maxWidth: '100%', objectFit: 'contain', opacity: 0.7 }}
                  />
                ) : (
                  <span style={{ fontSize: 12, color: '#9ca3af' }}>Watermark disabled or none</span>
                )}
              </div>

              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <label className="btn ghost sm" style={{ cursor: 'pointer', border: '1px solid #d1d5db' }}>
                  Upload Watermark
                  <input
                    type="file"
                    accept="image/*"
                    style={{ display: 'none' }}
                    onChange={handleWatermarkUpload}
                  />
                </label>
                <button
                  type="button"
                  className="btn ghost sm"
                  onClick={handleResetWatermark}
                  style={{ color: '#4b5563' }}
                >
                  Use Default Watermark
                </button>
              </div>
              <span style={{ display: 'block', fontSize: 11, color: '#6b7280', marginTop: 6 }}>
                Centered subtly in the background of quotation pages.
              </span>
            </div>
          </div>
        </div>

        {/* Action Footer */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            padding: '16px 20px',
            background: '#ffffff',
            border: '1px solid #d1d5db',
            borderRadius: 6,
            marginTop: 8,
          }}
        >
          <button
            type="button"
            className="btn ghost"
            onClick={handleResetDefaults}
            style={{ border: '1px solid #d1d5db' }}
          >
            Reset All to Default Myriad Template
          </button>

          <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
            {hasChanges && (
              <span style={{ fontSize: 12, color: '#d97706', fontWeight: 600 }}>
                ● Unsaved changes
              </span>
            )}
            <button
              type="button"
              className="btn primary"
              onClick={handleSave}
              style={{ fontWeight: 700, padding: '8px 24px' }}
            >
              Save Business Profile
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
