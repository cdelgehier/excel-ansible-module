[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
# Ansible Collection - cdelgehier.excel

This collection embeds a module allowing to manage basic actions with excel.

## Requirements

The only dependencies are:
- openpyxl
- os

## Usage

```yaml
- name: "Add a sheet in a non existant workbook"
  excel:
    operation: write
    data: "{{ data }}"
    table_name: pop_first_names
    path: "{{ playbook_dir }}"
    file: "test1.xlsx"
    worksheet: "my_title"
    create: true
```

## Changelog

See changelog.

## License

GNU GENERAL PUBLIC LICENSE v3
