import { buildSkillsPromptSection, type SkillMeta } from "@office-agents/core";
import rawSystemPrompt from "./system-prompt.md?raw";

export function buildExcelSystemPrompt(
  skills: SkillMeta[],
  commandSnippets: string[] = [],
): string {
  const customCommandsList = commandSnippets.map((s) => `  ${s}`).join("\n");
  return rawSystemPrompt
    .replace("{{CUSTOM_COMMANDS}}", customCommandsList)
    .replace("{{SKILLS_SECTION}}", buildSkillsPromptSection(skills));
}
