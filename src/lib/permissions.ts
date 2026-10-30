/**
 * Application Permission Enforcement Layer
 * Enforces strict role-based access control (RBAC) on the client side in addition to PostgreSQL RLS and RPC constraints.
 * Protects privileged actions: template publishing, rolling back/changing defaults, deleting projects, and migrating projects.
 */

import { api, Org, OrgRole, Session, Template } from './api'

export function isOwner(role: OrgRole | string): boolean {
  return role === 'owner'
}

export function isAdminOrOwner(role: OrgRole | string): boolean {
  return role === 'owner' || role === 'admin'
}

export function canManageTemplates(role: OrgRole | string): boolean {
  return isAdminOrOwner(role)
}

export function canDeleteProjects(role: OrgRole | string): boolean {
  return isAdminOrOwner(role)
}

export function canMigrateProjects(role: OrgRole | string): boolean {
  return isAdminOrOwner(role)
}

export function canManageMembers(role: OrgRole | string): boolean {
  return isAdminOrOwner(role)
}

export function assertPermission(role: OrgRole | string, allowed: Array<OrgRole | string>, action: string): void {
  if (!allowed.includes(role)) {
    throw new Error(`Permission denied: Action "${action}" requires role in [${allowed.join(', ')}], current role is "${role}".`)
  }
}

/**
 * Publishes a new template version with client-side role check.
 */
export async function publishTemplate(
  session: Session,
  org: Org,
  name: string,
  description: string,
  definition: any,
  draftId?: string
): Promise<{ id: string; version_num: number; status: string; is_default: boolean }> {
  assertPermission(org.role, ['owner', 'admin'], 'publishTemplate')
  return await api.templatePublish(session.token, org.id, name, description, definition, draftId)
}

/**
 * Updates an existing template version in-place without creating a new version.
 */
export async function updateTemplate(
  session: Session,
  org: Org,
  templateId: string,
  name: string,
  description: string,
  definition: any
): Promise<{ id: string; version_num: number; status: string; is_default: boolean }> {
  assertPermission(org.role, ['owner', 'admin'], 'updateTemplate')
  return await api.templateUpdate(session.token, org.id, templateId, name, description, definition)
}

/**
 * Rolls back or switches the default template version with client-side role check.
 */
export async function rollbackTemplateToVersion(
  session: Session,
  org: Org,
  templateId: string
): Promise<{ id: string; version_num: number; name: string; is_default: boolean }> {
  assertPermission(org.role, ['owner', 'admin'], 'rollbackTemplateToVersion')
  return await api.templateSetDefault(session.token, org.id, templateId)
}

/**
 * Permanently deletes a template version with client-side role check.
 */
export async function deleteTemplate(
  session: Session,
  org: Org,
  templateId: string
): Promise<{ id: string; version_num: number; deleted: boolean }> {
  assertPermission(org.role, ['owner', 'admin'], 'deleteTemplate')
  return await api.templateDelete(session.token, org.id, templateId)
}

/**
 * Permanently deletes a project with client-side role check.
 */
export async function deleteProject(
  session: Session,
  org: Org,
  projectId: string
): Promise<void> {
  assertPermission(org.role, ['owner', 'admin'], 'deleteProject')
  return await api.projectDelete(session.token, projectId)
}

/**
 * Migrates a project to a new template version with client-side role check.
 */
export async function migrateProjectToTemplate(
  session: Session,
  org: Org,
  projectId: string,
  targetTemplateId: string,
  inputs: any,
  formulaOverrides: any,
  total: number,
  summary: any
): Promise<void> {
  assertPermission(org.role, ['owner', 'admin'], 'migrateProjectToTemplate')
  return await api.projectRecalculateTemplate(session.token, projectId, targetTemplateId, inputs, formulaOverrides, total, summary)
}
