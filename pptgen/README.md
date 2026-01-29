pptgen — Weekly PowerPoint Generator

This folder contains a small Python generator that builds PowerPoint presentations from a template and a data file.

Usage

1. Install dependencies:

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r pptgen/requirements.txt
```

2. Run the generator:

```bash
python pptgen/src/generate_ppt.py --data pptgen/data/sample_data.csv --template template.pptx --output outputs/presentation.pptx
```

Place your `template.pptx` in the repository root or pass the `--template` path.

Place your Excel/CSV data file under `pptgen/data/` (or pass a path to `--data`). Use column names as placeholders in the template with the `{column_name}` syntax.

Example placeholder usage in a slide text box: `Quarter: {quarter}`

