# RRF Calculator v2

Desktop application for macOS and Windows for chromatographic RRF calculations.

## What changed in v2

The old dictionary schema:

```text
class | H_formula | summary_group
```

was replaced by a chemistry-oriented schema:

```text
name | type | formula_rule | summary_group
```

Two entry types are supported:

- `class` — a homologous class that requires `C<number>` in `Group Name`.
- `compound` — an exact compound with a fixed molecular formula.

Examples:

```text
aromatic | class | C[n]H[2*n-6] | monocyclic aromatics
phenol class example | class | C[n]H[2*n-6]O | phenols
benzoic acid | compound | C7H6O2 | acids
```

The built-in defaults contain only the original validated hydrocarbon classes. Oxygenated or other new classes can be added through the GUI.

## Molecular formula and molecular weight

For a class entry:

```text
Group Name: aromatic C10
rule: C[n]H[2*n-6]
result: C10H14
```

For an exact compound:

```text
Group Name: benzoic acid
formula_rule: C7H6O2
result: C7H6O2
```

The program parses the final formula and automatically calculates:

- C
- H
- Formula
- Molecular Weight
- RRF
- RRF*V
- w, %

Supported elements in v2: H, B, C, N, O, F, Si, P, S, Cl, Br, I.

## RRF model

The calculation is isolated in `core.calculate_rrf()`:

```text
RRF = MolecularWeight * 0.83905 / (12.011 * number_of_C_atoms)
```

For the original hydrocarbon classes this is exactly equivalent to the previous implementation. For heteroatom-containing compounds the full molecular weight is used in the numerator. If the analytical methodology requires another model, only `calculate_rrf()` needs to be changed.

## Backward compatibility

- Existing input names such as `aromatic C10`, `naphthene C12`, etc. continue to work.
- `C19+` and `C14 ?` continue to use the numeric C value with a warning.
- A class without `C<number>` is still skipped with a warning.
- Old user `classes.json` dictionaries are automatically migrated to the v2 format when the app starts.
- Class summary and heatmap are retained.

## Run on macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## Run on Windows

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Build a single Windows EXE

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --name "RRF Calculator" --add-data "default_classes.json;." app.py
```

The result will be:

```text
dist\RRF Calculator.exe
```
