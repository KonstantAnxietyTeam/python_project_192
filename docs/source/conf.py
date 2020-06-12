# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#
import os
import sys
sys.path.insert(0, os.path.abspath('../../Work/Scripts'))
sys.path.insert(0, os.path.abspath('../../Work/Library'))


# -- Project information -----------------------------------------------------

project = 'python_192'
copyright = '2020, Konstantinov, Sidorov, Berezutskiy'
author = 'Konstantinov, Sidorov, Berezutskiy'

# The full version, including alpha/beta/rc tags
release = '1.0'


# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom
# ones.
extensions = [
    'sphinx.ext.autodoc',
	'rst2pdf.pdfbuilder',
]
pdf_documents = [('index', u'rst2pdf', u'dev doc', u'192'),]
language = 'ru'
pdf_language = 'ru'

# Add any paths that contain templates here, relative to this directory.
templates_path = ['_templates']

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = []

# The name of the Pygments (syntax highlighting) style to use.
pygments_style = 'borland'



# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
html_theme = 'classic'
html_theme_options = {
    'stickysidebar': 'True',
    'footerbgcolor': '#560000',
    'sidebarbgcolor': '#3D3D3D',
    'relbarbgcolor': '#BF0000',
    'linkcolor': '#7D9EC0',
    'headbgcolor': '#F5B7B7',
    'headtextcolor': '#2F0000',
}


# Add any paths that contain custom static files (such as style sheets) here,
# relative to this directory. They are copied after the builtin static files,
# so a file named "default.css" will overwrite the builtin "default.css".
html_static_path = ['_static']