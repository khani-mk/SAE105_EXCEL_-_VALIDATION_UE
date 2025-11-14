# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'SAE105_EXCEL_-_VALIDATION_UE'
copyright = '2025, Karim HANI Rémi Camelio'
author = 'Karim HANI Rémi Camelio'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = []

templates_path = ['_templates']
exclude_patterns = [
'sphinx.ext.todo',
'sphinx.ext.mathjax',
'sphinx.ext.ifconfig',
'sphinx.ext.autodoc',
'sphinx.ext.viewcode',
]



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = 'alabaster'
html_static_path = ['_static']
