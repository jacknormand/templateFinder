# ADO Template Finder

Edge extension that finds template references in Azure DevOps YAML pipeline files and turns them into clickable links.

## Features
<img width="1301" height="779" alt="Screenshot 2026-03-05 091846" src="https://github.com/user-attachments/assets/6e146f81-276f-493a-9c89-ea6ce4d507c2" />




- Detects `template:` references in pipeline YAML files
- Resolves `@RepoAlias` using `resources.repositories` declarations
- Resolves `${{ variables.X }}` expressions by fetching external variable template files via the ADO REST API
- full resolution chain: where a variable is defined and what it resolves to
- Detects standalone variable template files (e.g., `EV2_Resources: path/to/file.yml@repo`)
- sections: Templates, Resolved via Variables, Available Template Variables, Unresolved
- Handles both `dev.azure.com` and `*.visualstudio.com` URL 
- re-analyzes on navigation without page reload

## How It Works

1. When you open a `.yml`/`.yaml` file in ADO, the extension reads the file path from the URL
2. It fetches the **full file content** via the ADO Git Items REST API (uses your session cookies)
3. It parses `resources.repositories` to build a map of repo aliases to actual repo names
4. It parses local variables (`- name: X` / `value: Y` and simple `Key: value` formats)
5. It finds all `- template:` lines that reference other YAML files with variables in the path (e.g., `${{variables.buildType}}`), resolves those variables, then **fetches the referenced files** to extract more variables
6. With the full merged variable map, it resolves `template: ${{ variables.EV2_Resources }}` to the actual file path and builds a clickable ADO URL
7. Results are shown in a floating sidebar panel, grouped by category

## Install (cus its not on edge store)

1. Go to `edge://extensions/`
2. Enable **Developer mode**
3. Click **Load unpacked** and select this folder (templateFinder)
4. DONE!

## Usage

1. Navigate to any YAML pipeline file in Azure DevOps
2. A **TF** button appears in the top-right corner
3. Click it to expand the panel with all template links
4. Click any link to open the template in a new tab

## Limitations

- Wont grab stuff in outer scope (e.g stuff that INHERITS the current resource. It wouldnt know where to look) 
- Variables defined in repos not listed in `resources.repositories` (e.g., OneBranch system repos) cannot be resolved
- The YAML parser is regex-based, not a full AST parser — works for standard pipeline patterns but may miss unusual formatting
- Only follows one level of variable template includes (does not recursively fetch templates referenced inside fetched templates)



