// Heuristic guard for explicitly staged public changes, not a confidentiality proof.
import { execFileSync } from "node:child_process";
const git = args => execFileSync("git",args,{encoding:"utf8",maxBuffer:20*1024*1024});
const files = git(["diff","--cached","--name-only","-z"]).split("\0").filter(Boolean);
const allowed = /^(README\.md|CHANGELOG\.md|package\.json|scripts\/(figure-style-test\.mjs|figure-typography-test\.py|package-preservation-test\.py|package-preservation-test\.mjs|compile-python\.mjs|check-public-diff\.mjs)|plugins\/scientific-illustrator\/(docs\/maintained-fork\/[^/]+\.md|styles\/figure-profiles\.json|scripts\/(figure-style\.mjs|pptx-package\.py|powerpoint-server\.mjs|powerpoint-mac-bridge\.py)|skills\/(edit-powerpoint-live|design-scientific-figure)\/SKILL\.md))$/;
const findings = files.filter(f=>!allowed.test(f)).map(f=>`${f}: outside the maintained-fork publish allowlist`);
const patch = git(["diff","--cached","--no-ext-diff","--unified=0","--no-color"]);
// Construct patterns in pieces so this guard's own source is not a false positive.
const patterns = [
  new RegExp("/(?:"+["Users","home"].join("|")+")/[^\\s/]+/"),
  new RegExp("(?:gh"+"p_|github_"+"pat_|sk-)[A-Za-z0-9_-]{20,}"),
  new RegExp("BEGIN "+"(?:RSA |EC |OPENSSH )?PRIVATE KEY"),
  new RegExp("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}"),
];
let file="";
for(const line of patch.split("\n")) {
  if(line.startsWith("+++ b/")){file=line.slice(6);continue;}
  if(!line.startsWith("+") || line.startsWith("+++"))continue;
  if(patterns.some(p=>p.test(line.slice(1)))) findings.push(`${file}: possible private path, email or credential in added content (value withheld)`);
}
if(findings.length){console.error(findings.join("\n"));process.exitCode=1;}
else console.log(`Staged-public-diff heuristic passed for ${files.length} explicitly allowed files. Human content and commit-metadata review still required.`);
