import {
  hashProtectionPassword,
  verifyProtectionPassword,
  type DocProtection,
} from '@genoffice/docx-engine'

export interface SetDocxProtectionInput {
  readonly enabled: boolean
  readonly password: string | null
}

export type SetDocxProtectionResult =
  | {
      readonly ok: true
      readonly enabled: boolean
      readonly passwordProtected: boolean
      readonly changed: boolean
      readonly protection: DocProtection | null
    }
  | {
      readonly ok: false
      readonly error: 'invalid_arguments' | 'execution_failed'
      readonly message: string
    }

/** Resolve document protection to one explicit final state without exposing password material. */
export async function setDocxProtection(
  current: DocProtection | null,
  input: SetDocxProtectionInput,
): Promise<SetDocxProtectionResult> {
  if (
    typeof input?.enabled !== 'boolean' ||
    (input?.password !== null &&
      (typeof input?.password !== 'string' ||
        Array.from(input.password).length < 1 ||
        Array.from(input.password).length > 255))
  ) {
    return {
      ok: false,
      error: 'invalid_arguments',
      message:
        'DOCX protection requires an enabled boolean and a null or 1–255 character password.',
    }
  }

  const active = current?.enforced === true && current.edit === 'readOnly'
  if (input.enabled && active) {
    return {
      ok: true,
      enabled: true,
      passwordProtected: Boolean(current.hash),
      changed: false,
      protection: current,
    }
  }
  if (!input.enabled && current === null) {
    return {
      ok: true,
      enabled: false,
      passwordProtected: false,
      changed: false,
      protection: null,
    }
  }

  try {
    if (input.enabled) {
      const credentials = input.password ? await hashProtectionPassword(input.password) : {}
      const protection: DocProtection = { edit: 'readOnly', enforced: true, ...credentials }
      return {
        ok: true,
        enabled: true,
        passwordProtected: Boolean(protection.hash),
        changed: true,
        protection,
      }
    }

    if (current?.hash) {
      if (!input.password || !(await verifyProtectionPassword(input.password, current))) {
        return {
          ok: false,
          error: 'execution_failed',
          message: 'The DOCX protection password is incorrect.',
        }
      }
    }
    return {
      ok: true,
      enabled: false,
      passwordProtected: false,
      changed: true,
      protection: null,
    }
  } catch (error) {
    return {
      ok: false,
      error: 'execution_failed',
      message: `The DOCX protection state could not be changed: ${
        error instanceof Error ? error.message : String(error)
      }`,
    }
  }
}
