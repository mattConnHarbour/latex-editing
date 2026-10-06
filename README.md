# SuperDoc equation rendering demo

This standalone React demo opens `public/calculus-equations.docx`, a fixture made from native Word equation (OMML) objects. The document covers common calculus notation, including stacked quotients, derivatives, limits, integrals, summations, products, radicals, series, and multivariable formulas.

## Run it

Requires Node 22.12 or newer and pnpm 11.

```bash
pnpm install
pnpm dev
```

Open the printed URL and select any rendered equation. The anchored editor converts the existing OMML to LaTeX, shows a debounced KaTeX preview, converts the edited expression back to OMML, patches the DOCX in the browser, and reloads it in SuperDoc. Choose **Export DOCX** to download the edited document.

## Verify it

```bash
pnpm typecheck
pnpm build
pnpm test:e2e
```

Regenerate the fixture after changing the formula catalog:

```bash
python3 scripts/generate_fixture.py
```
