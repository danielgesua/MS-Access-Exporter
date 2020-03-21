# Plaintext (Diffable) Exporter for Microsoft Access DB Modules

Wouldn't it be nice if we could use git differential commands on MS Access database files like *.accdb's?
The purpose of this project is to make this a reality by exporting all the *.accdb's modules, tables, 
queries and other objects as *.bas, *.cls and *.txt files.

# Requirements

Windows:

```
py -m pip install -r requirements.txt
```

Linux/OS X:

```
pip install -r requirements.txt
```

# Output

## Tables

Tables export as JSON dumps, and so can be read as clean syntax.

## Queries

Queries export as MSSQL and optionally can be prettified using sql-format.com by passing a second "True" parameter like so:

```
py access_db_exporter.py path/to/access.accdb True
```

## Modules

Modules are stored as cls/bas files which are plaintext and can be easily compared.

# Usage

## One-time

1. py access_db_exporter.py
2. Select an access DB vis File Select

## With all arguments including prettification of queries

```
py access_db_exporter.py path/to/access.accdb True
```

## Via Git Pre-Commit Hooks

1. copy sample_hook_scripts/pre-commit.sample path/to/.git/hooks/pre-commit
2. edit path/to/.git/hooks/pre-commit with proper paths to access DB and exporter binary/python file

