# AI Enablement Deck Trust Audit

This document checks whether the planned slide messages can be supported by reasonably strong public sources.

Where a statement was too broad or too absolute, it was replaced with a narrower, better-supported version.

## Assessment Scale

- hard-sourceable
- sourceable after narrowing
- avoid as factual claim

## Slide-by-Slide Review

### Slide 1

Original status:

- title only, no factual issue

Result:

- hard-sourceable not required

### Slide 2

Original claim:

- AI is already affecting everyday work, but value without understanding creates poor output and avoidable risk.

Assessment:

- sourceable after narrowing

Replacement:

- Generative AI is increasingly used for drafting, summarizing, and information support, so practical usage and review habits matter.

Rationale:

- the original wording was true in spirit but broad and hard to pin to one source without slipping into adoption claims

Sources:

- Microsoft Support: prompt writing and productivity tasks
- Anthropic: prompting guidance and importance of clear instructions

### Slide 3

Original claim:

- Modern AI is strong at language, pattern, structure, summarization, and drafting tasks.

Assessment:

- sourceable after narrowing

Replacement:

- Current generative AI systems are useful for drafting, summarizing, transforming, and brainstorming text-heavy work.

Sources:

- Microsoft Support prompt examples

### Slide 4

Original claim:

- AI does not truly understand, does not reliably know what is true, and should not be treated like an oracle.

Assessment:

- avoid as factual claim

Replacement:

- Generative AI can produce fluent but inaccurate output, so users should not assume correctness without verification.

Rationale:

- "does not truly understand" is philosophically loaded and not ideal as a hard factual claim in a public deck

Sources:

- Anthropic documentation on reducing hallucinations

### Slide 5

Original claim:

- Modern AI generates outputs from patterns and probabilities, not human-like understanding.

Assessment:

- sourceable after narrowing

Replacement:

- Large language models generate text by predicting likely token sequences from patterns learned during training data.

Sources:

- Hugging Face course: language modeling and next-token prediction
- Anthropic prompt engineering context

### Slide 6

Original claim:

- The biggest value often comes from supporting knowledge work before decisions or final delivery.

Assessment:

- sourceable after narrowing

Replacement:

- A practical starting point is support work around drafting, summarization, preparation, and restructuring in knowledge work.

Rationale:

- "biggest value" is too broad without organization-specific evidence

Sources:

- Microsoft Support prompt examples

### Slide 7

Original claim:

- Good AI use is a repeatable workflow, not a magic trick.

Assessment:

- sourceable after narrowing

Replacement:

- Using explicit context, role, goal, examples, and review tends to improve output quality.

Rationale:

- the original line is good rhetoric but not a strong factual statement

Sources:

- Microsoft Support prompt guidance
- Anthropic prompt engineering guide

### Slide 8

Original claim:

- Poor prompts, missing context, and blind trust create poor outcomes quickly.

Assessment:

- sourceable after narrowing

Replacement:

- Weak instructions and lack of review increase the chance of generic, irrelevant, or incorrect output.

Sources:

- Anthropic prompt engineering guide
- Anthropic hallucination guidance

### Slide 9

Original claim:

- AI systems that can act, trigger tools, or change state are more useful, but also more dangerous, than chat-only systems.

Assessment:

- sourceable after narrowing

Replacement:

- AI systems that can trigger tools or external actions require stronger safeguards, oversight, and limits than chat-only use.

Rationale:

- "more dangerous" is understandable but informal; the replacement is more precise and governance-compatible

Sources:

- NIST AI Risk Management Framework

### Slide 10

Original claim:

- Responsible AI use needs boundaries, review points, and explicit rules.

Assessment:

- hard-sourceable

Replacement:

- Responsible use benefits from defined boundaries, human review, and clear rules for higher-impact tasks.

Sources:

- NIST AI Risk Management Framework

### Slide 11

Original claim:

- AI is useful when treated as a tool with context, verification, and human accountability.

Assessment:

- sourceable after narrowing

Replacement:

- Useful everyday AI use depends on context, verification, and human accountability.

Sources:

- Anthropic prompt engineering guide
- NIST AI Risk Management Framework

### Slide 12

Original claim:

- AI creates value when used with method and guardrails, not hype or blind trust.

Assessment:

- sourceable after narrowing

Replacement:

- AI is most useful when applied with clear prompts, review, and appropriate guardrails.

Sources:

- Microsoft prompt guidance
- Anthropic prompt engineering guide
- NIST AI Risk Management Framework

## Net Result

Most of the deck is defensible after tightening a few broad formulations.

The main trust risk was not false content, but overly absolute wording.

That risk is now reduced by:

- replacing philosophy-like claims with operational claims
- avoiding unsupported superlatives
- grounding behavior and guardrails in public guidance
