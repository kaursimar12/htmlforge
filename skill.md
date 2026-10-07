Anti-Patterns Summary

PrincipleAnti-PatternFixThink Before CodingSilently assumes file format, fields, scopeList assumptions explicitly, ask for clarificationSimplicity FirstStrategy pattern for single discount calculationOne function until complexity is actually neededSurgical ChangesReformats quotes, adds type hints while fixing bugOnly change lines that fix the reported issueGoal-Driven"I'll review and improve the code""Write test for bug X → make it pass → verify no regressions"
Key Insight

The "overcomplicated" examples aren't obviously wrong—they follow design patterns and best practices. The problem is timing: they add complexity before it's needed, which:

Makes code harder to understand
Introduces more bugs
Takes longer to implement
Harder to test
The "simple" versions are:

Easier to understand
Faster to implement
Easier to test
Can be refactored later when complexity is actually needed
Good code is code that solves today's problem simply, not tomorrow's problem prematurely.