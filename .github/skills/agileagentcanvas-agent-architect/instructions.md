# agileagentcanvas-agent-architect instructions

This file explains how to create or update a skill-level `instructions.md` for the `agileagentcanvas-agent-architect` agent.

## Purpose
- Describe the agent's intended role and its behavior.
- Explain what users should expect when interacting with the skill.
- Capture any project-specific conventions or preferences for architecture guidance.

## When to create this file
Create or update this file when you want to:
- clarify what the architect agent should do
- document user-facing guidance for the skill
- preserve rules or preferences that should be enforced across conversations

## Recommended structure
Use these sections as a template:

1. `## Intent`
   - Briefly summarize the agent's role.
2. `## Activation`
   - Describe when this skill should be used.
   - Include any trigger phrases or types of requests.
3. `## Behavior`
   - Explain the agent's communication style.
   - State the output format expectations.
4. `## Project conventions`
   - Call out any repo-specific rules or coding preferences.
   - Mention file paths, naming schemes, or language choices.
5. `## Example prompts`
   - Give users sample questions that map to the skill.

## Example template
```md
## Intent
Explain architectural decisions, trade-offs, and solution structure for Agile Agent Canvas extensions.

## Activation
Use this skill when the user asks for architecture review, design patterns, or high-level technical direction.

## Behavior
- Answer in clear, structured sections.
- Prefer concise recommendations with rationale.
- Avoid low-level implementation details unless asked.

## Project conventions
- Follow the repository's existing naming and folder patterns.
- Keep PowerShell scripts and ARM/Bicep design guidance aligned with the current workspace.

## Example prompts
- "Review the architecture for this agent extension."
- "How should I structure a new BMAD skill?"
```

## Notes
- Keep the file short and focused.
- Use markdown headings and bullet lists for readability.
- If the repo evolves, update this file to preserve the current style and expectations.
