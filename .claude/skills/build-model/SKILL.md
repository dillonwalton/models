---
name: build-model
description: Create a new financial model for a public company from the repo's base model template. Use when the user asks to build, create, or add a model for a company or ticker (e.g. "build a model for CoreWeave", "add a NVDA model", "new model for Digital Realty").
---

# Build a model

Create a new company model by copying the base template.

## Steps

1. **Resolve the ticker.** The user may give a company name (e.g. "CoreWeave") or a ticker
   (e.g. "CRWV"). The filename must be the company's public exchange ticker, uppercase.
   If you are not confident of the ticker, ask the user rather than guessing.

2. **Check it doesn't already exist.** If `<TICKER>.xlsx` is already in the repo root, stop
   and tell the user — do not overwrite an existing model.

3. **Copy the template** to the repo root, named for the ticker:

   ```bash
   cp "base model.xlsx" "<TICKER>.xlsx"
   ```

4. **Confirm** to the user which file was created. Leave it uncommitted unless they ask
   you to commit.

## Conventions

- Models live in the repo root, one file per company, named `<TICKER>.xlsx`.
- The template is `base model.xlsx` — two sheets, `Main` and `Model`, with quarterly
  revenue rows. Never edit the template itself.
- Excel lock files (`~$*.xlsx`) are gitignored; ignore them.
