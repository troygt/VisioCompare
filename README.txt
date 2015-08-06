VisioCompare v1.0
Author: Troy Tucker

Purpose: Compare old and new versions of the same name Visio diagrams and report differences in a side-by-side html output file.

These scripts basically allow me to convert, compare, and report on differences between a multitude of network diagrams, with multiple pages, and store
reports and diff files on what changed.  For example, we require our network management company to provide quarterly updates on all network topology
diagram changes. These scripts allow me to compare Q1 diagrams with Q2 diagrams in bulk and show the differences via an html diff file and a report.  


Files:

1. visio2VSDX.py - takes a target directory, scans for .vsd files, renames any file with a space in the name, saves the files as a .vsdx file

2. visioVSDXCompare.py - takes a old source directory and a new source directory of .vsdx visio files, loads each file found, runs a diff compare between the old
and the new, at the page level, and then saves htmldiff output showing a side by side comparison of the differences.

*Both files use Tkinter dialogs to make folder selection more friendly... but you can easily pull that out in favor of command line args

Library Requirements: difflib, zipfile

Application Requirements: Visio 2013