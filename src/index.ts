/**
 * DeepSeek Harness Cordis plugin provider for the bundled `bio-thesis-formatter` skill.
 * Zero runtime dependencies: registers directly against `ctx.skills`.
 *
 * @module dsh-plugin-bio-thesis
 */

import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import type { Context } from '@deepseek-ai/cordis'
import type { SkillCandidate, SkillDefinition, SkillProvider } from '@deepseek-ai/dsh-skill'

const PROVIDER_NAME = 'bio-thesis-formatter'
const BUNDLED_SKILL_RANK = 600
const SKILL_DIR = fileURLToPath(new URL('../skills/bio-thesis-formatter/', import.meta.url))
const SKILL_MD_URL = new URL('../skills/bio-thesis-formatter/SKILL.md', import.meta.url)
const SKILL_MD_PATH = fileURLToPath(SKILL_MD_URL)
const RESOURCE_BASE = { kind: 'directory', path: SKILL_DIR } as const

interface ParsedSkill {
  name: string
  description: string
  whenToUse?: string
  modelInvocable: boolean
  userInvocable: boolean
  content: string
}

function parseBoolFlag(raw: string | undefined, fallback: boolean): boolean {
  if (raw === undefined) return fallback
  const v = raw.trim().toLowerCase()
  if (v === 'true' || v === 'yes' || v === 'on' || v === '1') return true
  if (v === 'false' || v === 'no' || v === 'off' || v === '0') return false
  return fallback
}

function parseSkillMarkdown(raw: string): ParsedSkill {
  const m = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/u.exec(raw)
  if (!m || m[1] === undefined) {
    throw new Error('dsh-plugin-bio-thesis: SKILL.md is missing YAML frontmatter')
  }
  const fields: Record<string, string> = {}
  for (const line of m[1].split(/\r?\n/u)) {
    if (!line.trim() || line.startsWith(' ') || line.startsWith('#')) continue
    const idx = line.indexOf(':')
    if (idx <= 0) continue
    const key = line.slice(0, idx).trim()
    let val = line.slice(idx + 1).trim()
    if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
      val = val.slice(1, -1)
    }
    fields[key] = val
  }
  const skillName = fields.name || 'bio-thesis-formatter'
  const description = fields.description || ''
  if (!description) {
    throw new Error('dsh-plugin-bio-thesis: SKILL.md frontmatter requires description')
  }
  const disableModel = parseBoolFlag(fields['disable-model-invocation'], false)
  const userInvocable = parseBoolFlag(fields['user-invocable'], true)
  return {
    name: skillName,
    description,
    whenToUse: fields.whenToUse,
    modelInvocable: !disableModel,
    userInvocable,
    content: raw.slice(m[0].length).trim(),
  }
}

const provider: SkillProvider = {
  name: PROVIDER_NAME,
  async list(): Promise<SkillCandidate[]> {
    const raw = await readFile(SKILL_MD_URL, 'utf8')
    const parsed = parseSkillMarkdown(raw)
    return [
      {
        name: parsed.name,
        description: parsed.description,
        ...(parsed.whenToUse ? { whenToUse: parsed.whenToUse } : {}),
        invocation: {
          modelInvocable: parsed.modelInvocable,
          userInvocable: parsed.userInvocable,
        },
        provider: PROVIDER_NAME,
        source: 'bundled',
        resourceBase: RESOURCE_BASE,
        path: SKILL_MD_PATH,
        rank: BUNDLED_SKILL_RANK,
        locator: SKILL_MD_URL,
      },
    ]
  },
  async get(): Promise<SkillDefinition> {
    const raw = await readFile(SKILL_MD_URL, 'utf8')
    const parsed = parseSkillMarkdown(raw)
    return {
      name: parsed.name,
      description: parsed.description,
      ...(parsed.whenToUse ? { whenToUse: parsed.whenToUse } : {}),
      invocation: {
        modelInvocable: parsed.modelInvocable,
        userInvocable: parsed.userInvocable,
      },
      provider: PROVIDER_NAME,
      source: 'bundled',
      resourceBase: RESOURCE_BASE,
      path: SKILL_MD_PATH,
      content: parsed.content,
    }
  },
}

/** Cordis plugin namespace. Function plugins must use named exports only (no default export). */
export const name = 'skill-bio-thesis'
/** Required Cordis service injection. */
export const inject = ['skills'] as const

/** Register the bundled bio-thesis-formatter skill provider on `ctx.skills`. */
export function apply(ctx: Context): void {
  ctx.skills.registerProvider(() => provider)
}
