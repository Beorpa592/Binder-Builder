# Default templates

The app renders generated documents from these templates unless the binder has
its own copy in its template folder (`.binder/templates/`, or
`binder.template_folder`). To restyle a binder, copy a template there, open it in
Word or Excel, and change fonts, sizes and layout. Keep the placeholders. On the
next sync, the plan shows every document built from that template as
`regenerate (template changed)`.

| Template | Used for | Placeholders |
| --- | --- | --- |
| `00_Cover.docx` | Binder cover | docxtpl: `{%p for line in lines %}{{ line }}{%p endfor %}`, `{{ binder_name }}` |
| `Section_Cover.docx` | Every section cover sheet | docxtpl: `{{ heading }}` (`A) Administration`), `{{ code }}`, `{{ label }}`, `{{ description }}`, loop `rows` with `r.mark` (`__X__` or `__ __`), `r.code`, `r.title` |
| `01_TOC.xlsx` | Table of Contents | Rows starting `{{#section}}`, `{{#header}}`, `{{#item}}`, `{{#subitem}}` are row templates (values and styles copied per row) with tokens `{{mark}}`, `{{code}}`, `{{title}}`, `{{label}}`, `{{page}}`. Elsewhere: `{{binder_name}}`, `{{page_heading}}` |
| `Tabs.docx` | Tab labels | The first table is the layout: its cells are filled in reading order, and copies of it are appended when there are more sections than cells. Each cell's first run keeps its formatting |
| `02_CKL.xlsx` | A new Master Checklist, only when none exists | Header row must contain `Reference` and `Description`. The app adds a hidden `bb_id` column |

Headers and footers are never part of a template. The app stamps them after rendering.

Rebuild the defaults with `python -m binder_builder.templates.build_defaults`.
