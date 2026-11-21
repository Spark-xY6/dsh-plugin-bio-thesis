# Campus benchmark v1.0.0

The corpus, answer key, task, scoring source and success criteria were frozen on 2026-11-21 before scoring any contestant or running the matched model arms. All documents, identities, numerical settings and control examples were created for this test. The fixtures have no connection to a real author's thesis or a school's actual submission requirements.

| Frozen item | SHA-256 |
| --- | --- |
| design.json | 1d5ef7aefeb01b401ea410125bba47fa4060c66465c9ebbc17aae37149b48a1e |
| inputs/TASK.md | 6b588bef823c902dd6d76b5f206f5e9af943bc1a6ba6e9ee83408be9b30d5e9c |
| build.mjs | 8a8d174dcb349259d69886ab9faf871fd6a6efefec12d4601775d4e5d5451b84 |
| input-sha256.json (hashes for 11 input files) | 371b464fbfb883e24235c7d4e0b3fa35af6a94738af921179e50afc0ccfbb44b |
| score.py, initially reviewed | 5a2d472be5a84d18ad07e233ddbdaa8f3be086d683939c9701bde007f2629fa5 |
| score.py, corrected before any contestant outputs were observed | b2fce3340624aa9891b4c3712d6ea8c5784c02b92f2a10d6e387fbe4fd1175eb |

During independent code review, before observing contestant outputs, the scorer was corrected to award partial credit when an otherwise readable DOCX omits optional styles or document relationships. The DOCX format item now checks the package-to-document relationship, and missing or unreadable generations are labelled explicitly. The corpus, task, gold labels, 12 generation criteria, decision thresholds and model instructions were unchanged.

Only the contents of inputs/ go into either solver's workspace. The arm output files are result.json (one set of labels for each of eight documents) and generated.docx. Run the scorer in the evaluator's directory with the arm output as its argument. The fixed matrix has 56 binary decisions, 14 positives, 42 negatives and two clean controls. The valid DOCX generation task has 12 independently counted checks. Missing or malformed audit rows count as incorrect for all seven decisions. A missing or unreadable generated file earns 0/12. Report malformed rows and completion separately.

Within a matched pair, keep the model, reasoning settings, prompt, corpus, resource limits and host constant. Alternate order when practical. For three pairs list all scores and counts. Compute audit micro-F1 as twice the true-positive count divided by twice that count plus false positives and false negatives. The quality and efficiency decision rules in design.json were fixed in advance. Time and tool calls, when available, are secondary endpoints. Separate the installable plugin's host from comparisons made in a different agent host.

For another tool, retain its raw output and identify which checks can be evaluated with its documented interface. An unimplemented or inapplicable rule is an abstention. Report supported-scope accuracy and end-to-end coverage separately. Compare product output with the same task only where the purpose and input overlap. No score in this benchmark establishes a worldwide novelty claim.
