import { describe, expect, it, beforeAll } from "vitest";
import "fake-indexeddb/auto";
import { getInstalledSkills } from "../../src/skills";
import { StorageNamespace } from "../../src/context";
import { saveSkillFiles, deleteSkillFiles, listSkillNames } from "../../src/storage/db";

const ns: StorageNamespace = {
  dbName: "perf_test_db_baseline",
  dbVersion: 1,
  documentSettingsPrefix: "perf_test",
};

describe("getInstalledSkills performance", () => {
  beforeAll(async () => {
    // Add 1000 skills to measure performance
    for (let i = 0; i < 1000; i++) {
      const skillName = `skill-${i}`;
      const mdContent = `---
name: ${skillName}
description: description for ${skillName}
---
# ${skillName}`;
      await saveSkillFiles(ns, skillName, [
        { path: "SKILL.md", data: new TextEncoder().encode(mdContent) },
        { path: "other.txt", data: new TextEncoder().encode("other data") }
      ]);
    }
  });

  it("measures performance of getInstalledSkills", async () => {
    const start = performance.now();
    const skills = await getInstalledSkills(ns);
    const end = performance.now();

    console.log(`getInstalledSkills for 1000 skills took ${end - start} ms`);
    expect(skills.length).toBe(1000);
  });
});
