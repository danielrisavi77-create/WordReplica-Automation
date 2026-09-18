"""WordReplica Document Fidelity Lab.

A continuously-running corpus lab: it ingests real and generated .docx
documents, fingerprints and scores them without Microsoft Word, schedules the
expensive Word-backed checks against the few that earn them, and minimizes any
document that finds a defect into a permanent regression fixture.

Nothing in this package may write Golden automation state or affect promotion
to main.
"""
