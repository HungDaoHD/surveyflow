"""Thin wrapper — delegates to surveyflow.cli.main().

Kept in project root for convenience during local development.
After ``pip install surveyflow``, use ``surveyflow-run`` instead.
"""
from surveyflow.cli import main

if __name__ == "__main__":
    main()
