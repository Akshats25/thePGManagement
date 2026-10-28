import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const workspaceDir = 'D:/rental-main';
const skillDir = 'C:/Users/aksha/.codex/plugins/cache/openai-primary-runtime/presentations/26.1025.11240/skills/presentations';
const buildDir = path.join(workspaceDir, '.ppt-build');
const finalDir = path.join(workspaceDir, 'deliverables');
const candidatePath = path.join(buildDir, 'panesar-living-interview-draft.pptx');
const finalPath = path.join(finalDir, 'Panesar_Living_Interview_Presentation.pptx');
const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, 'container_tools/artifact_tool_utils.mjs')).href);
await fs.mkdir(finalDir, { recursive: true });
await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: 'C:/Users/aksha/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
  integrityValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(skillDir, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit'],
  requiredNativeTableOwnerSlides: [],
  fontPolicy: { basis: 'design', families: ['Aptos'] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, 'panesar-living-validation.json'),
});
console.log(finalPath);
